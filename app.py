from datetime import date
from pathlib import Path

import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st

from core.excel_store import ExcelStore, to_numeric

st.set_page_config(page_title='ScrumMaster Governance Pro',page_icon='📊',layout='wide')
st.markdown('''<style>
:root{--navy:#123B63;--navy2:#0B2742;--red:#C62828;--green:#2E7D32;--bg:#F4F7FB;--text:#102A43}
.stApp{background:var(--bg)} [data-testid="stSidebar"]{background:linear-gradient(180deg,var(--navy2),var(--navy))}

[data-testid="stSidebar"] *{color:white}.block-container{padding-top:1rem;max-width:1800px}.hdr{background:linear-gradient(100deg,var(--navy2),var(--navy));color:white;padding:20px 30px;border-radius:16px;margin-bottom:18px;box-shadow:0 8px 24px #123b6328}.card{background:white;border-radius:12px;padding:16px;border-left:6px solid var(--navy);box-shadow:0 4px 14px #102a4314;min-height:110px}.red{border-left-color:var(--red)}.green{border-left-color:var(--green)}.amber{border-left-color:#F9A825}.lbl{color:#627D98;font-size:.8rem;text-transform:uppercase;font-weight:700}.val{color:var(--text);font-size:1.75rem;font-weight:800;margin-top:8px}.sub{color:#627D98;font-size:.85rem}.stDataFrame,.stDataEditor{background:white;border-radius:12px;box-shadow:0 3px 12px #102a4312;padding:4px}</style>''',unsafe_allow_html=True)
ROOT=Path(__file__).resolve().parent; TEMPLATE=ROOT/'templates'/'MPP_Upload_Template.xlsx'; store=ExcelStore(); store.ensure_app_sheets()
def hdr(t,s):st.markdown(f'<div class="hdr"><h2 style="margin:0">{t}</h2><div style="opacity:.86;margin-top:7px">{s}</div></div>',unsafe_allow_html=True)
def card(l,v,s='',tone=''):st.markdown(f'<div class="card {tone}"><div class="lbl">{l}</div><div class="val">{v}</div><div class="sub">{s}</div></div>',unsafe_allow_html=True)
def text(v):return '' if pd.isna(v) else str(v).strip()
def active(sheet,col):
    d=store.read_sheet(sheet)
    if d.empty or col not in d:return []
    if 'Active' in d:d=d[d['Active'].fillna(True).astype(bool)]
    return sorted([x for x in d[col].dropna().astype(str).str.strip().unique() if x])
def prep(d):
    if d.empty:return d
    d=d.copy()
    for c in ['Start Date','End Date','Last Updated']:
        if c in d:d[c]=pd.to_datetime(d[c],errors='coerce')
    for c in ['Duration Days','Progress %','Expected %','Variance %']:
        if c in d:d[c]=to_numeric(d[c])
    return d
def health(d):
    if d.empty:return d
    x=d.copy();today=pd.Timestamp(date.today());dur=(x['End Date']-x['Start Date']).dt.days.clip(lower=1);elapsed=(today-x['Start Date']).dt.days.clip(lower=0)
    x['Expected %']=((elapsed/dur)*100).clip(0,100).where(today>=x['Start Date'],0).where(today<x['End Date'],100).round(1);x['Variance %']=(x['Progress %']-x['Expected %']).round(1);x['Risk']='Low';x.loc[x['Variance %']<-10,'Risk']='Medium';x.loc[x['Variance %']<-25,'Risk']='High';x.loc[(x['End Date']<today)&(x['Progress %']<100),'Risk']='High';return x
def assign_task_ids(df,existing):
    used=set(existing.get('Task ID',pd.Series(dtype=str)).dropna().astype(str));n=0
    for tid in used:
        digits=''.join(ch for ch in tid if ch.isdigit());n=max(n,int(digits) if digits else 0)
    for i in df.index:
        if not text(df.at[i,'Task ID']):n+=1;df.at[i,'Task ID']=f'TASK-{n:04d}'
    return df
def normalize_upload(raw):
    cols=store.read_sheet('App Tasks').columns.tolist();raw=raw.copy();raw.columns=[str(c).strip() for c in raw.columns]
    lookup={c.lower():c for c in cols};raw=raw.rename(columns={c:lookup.get(c.lower(),c) for c in raw.columns})
    for c in cols:
        if c not in raw:raw[c]=''
    raw=raw[cols];raw=assign_task_ids(raw,store.read_sheet('App Tasks'))
    for c in ['Start Date','End Date','Last Updated']:raw[c]=pd.to_datetime(raw[c],errors='coerce')
    raw['Progress %']=pd.to_numeric(raw['Progress %'],errors='coerce').fillna(0).clip(0,100);raw['Duration Days']=(raw['End Date']-raw['Start Date']).dt.days+1;raw['Last Updated']=pd.Timestamp(date.today());return raw
def prepare_master_editor(df):
    d=df.copy()

    # Streamlit data_editor requires configured columns to use compatible dtypes.
    if 'Active' in d.columns:
        d['Active']=d['Active'].fillna(True).astype(bool)

    for c in ['Start Date','End Date']:
        if c in d.columns:
            d[c]=pd.to_datetime(d[c],errors='coerce')

    for c in ['Sort Order']:
        if c in d.columns:
            d[c]=pd.to_numeric(d[c],errors='coerce').astype('Int64')

    # IDs and other master-data text columns must remain strings, not mixed numeric/object.
    for c in d.columns:
        if c not in ['Active','Start Date','End Date','Sort Order']:
            d[c]=d[c].fillna('').astype(str)

    return d

def save_master(sheet,id_col,prefix,edited):
    d=edited.copy()
    d=d[~d['_Delete'].fillna(False)].drop(columns=['_Delete'])

    # Ignore completely blank rows created by the dynamic editor.
    business_cols=[c for c in d.columns if c not in [id_col,'Active']]
    if business_cols:
        non_blank=d[business_cols].apply(
            lambda row:any(text(v) for v in row),
            axis=1
        )
        d=d[non_blank].copy()

    existing=store.read_sheet(sheet)

    # Generate stable IDs without duplicating values during the same save.
    used=set(existing.get(id_col,pd.Series(dtype=str)).dropna().astype(str))
    next_number=0
    for value in used:
        digits=''.join(ch for ch in value if ch.isdigit())
        if digits:
            next_number=max(next_number,int(digits))

    for i in d.index:
        current=text(d.at[i,id_col])
        if not current:
            next_number+=1
            current=f'{prefix}-{next_number:03d}'
            while current in used:
                next_number+=1
                current=f'{prefix}-{next_number:03d}'
            d.at[i,id_col]=current
        used.add(current)

    if 'Active' in d.columns:
        d['Active']=d['Active'].fillna(True).astype(bool)

    store.write_sheet(sheet,d)
    st.success('Saved successfully.')
    st.rerun()

with st.sidebar:
    st.markdown('## ScrumMaster Pro');st.caption('Delivery Governance Workspace')
    page=st.radio('Navigation',['Executive Summary','Master Data Management','MPP / Delivery Plan','Sprint Board','Resource Loading','RAID Log','Governance','Reports & Export'],label_visibility='collapsed')
    st.divider();st.caption(f'Backend: {store.path.name}')
    if st.button('🔄 Refresh data',use_container_width=True):st.rerun()

tasks=health(prep(store.read_sheet('App Tasks')));raid=store.read_sheet('App RAID');resources=store.read_sheet('App Resources')
actual=float(tasks['Progress %'].mean()) if not tasks.empty else 0;expected=float(tasks['Expected %'].mean()) if not tasks.empty else 0;high=int((tasks['Risk']=='High').sum()) if not tasks.empty else 0;score=max(0,min(100,round(100-abs(min(0,actual-expected))*.65-high*4)));rag='Green' if score>=80 else 'Amber' if score>=60 else 'Red'
if page=='Executive Summary':
    hdr('Executive Project Health','Real-time delivery health, milestones and proactive recovery guidance.');cs=st.columns(4)
    with cs[0]:card('Health Score',f'{score}/100',rag,rag.lower())
    with cs[1]:card('Actual Progress',f'{actual:.1f}%','Average completion')
    with cs[2]:card('Expected Progress',f'{expected:.1f}%','Current date baseline')
    with cs[3]:card('High-Risk Tasks',high,'Requires intervention','red' if high else 'green')
    f=go.Figure([go.Bar(name='Expected',x=['Project'],y=[expected]),go.Bar(name='Actual',x=['Project'],y=[actual])]);f.update_layout(barmode='group',height=350,yaxis_range=[0,100]);st.plotly_chart(f,use_container_width=True)
elif page=='Master Data Management':
    hdr('Project and Planning Master Data','Add, update or delete Projects, Epics, Sprints, Owners, Statuses and Priorities.')
    tabs=st.tabs(['Projects','Epics','Sprints','Owners','Statuses','Priorities']);cfgs=[('App Projects','Project ID','PRJ'),('App Epics','Epic ID','EPC'),('App Sprints','Sprint ID','SPR'),('App Owners','Owner ID','OWN'),('App Statuses','Status ID','STS'),('App Priorities','Priority ID','PRI')]
    for tab,(sheet,id_col,prefix) in zip(tabs,cfgs):
        with tab:
            d=store.read_sheet(sheet)
            v=prepare_master_editor(d)
            v.insert(0,'_Delete',False)

            cc={
                '_Delete':st.column_config.CheckboxColumn('Delete',default=False),
                id_col:st.column_config.TextColumn(id_col,disabled=True),
            }

            if 'Active' in v.columns:
                cc['Active']=st.column_config.CheckboxColumn('Active',default=True)

            for c in ['Start Date','End Date']:
                if c in v.columns:
                    cc[c]=st.column_config.DateColumn(c,format='DD-MMM-YYYY')

            if 'Sort Order' in v.columns:
                cc['Sort Order']=st.column_config.NumberColumn(
                    'Sort Order',min_value=0,step=1,format='%d'
                )

            e=st.data_editor(
                v,
                use_container_width=True,
                height=470,
                hide_index=True,
                num_rows='dynamic',
                column_config=cc,
                key=f'master_{sheet}',
            )
            st.caption('Add a blank row to create. Edit a row to update. Tick Delete and save to remove.')
            if st.button('💾 Save '+sheet.replace('App ',''),type='primary',key='save'+sheet):save_master(sheet,id_col,prefix,e)
elif page=='MPP / Delivery Plan':
    hdr('MPP-Style Delivery Plan','Professional task CRUD, bulk upload, template download, validation and interactive Gantt.')
    projects=active('App Projects','Project Name');epics=active('App Epics','Epic Name');sprints=active('App Sprints','Sprint');owners=active('App Owners','Owner Name');statuses=active('App Statuses','Status Name') or ['Planned','In Progress','Blocked','On Hold','Complete'];priorities=active('App Priorities','Priority Name') or ['Critical','High','Medium','Low']
    up,plan,gantt=st.tabs(['Upload / Template','Task Create, Update & Delete','Gantt View'])
    with up:
        c1,c2=st.columns([1,1.5])
        with c1:st.markdown('### Download MPP Template');st.write('Task ID can be blank; the application generates it.');st.download_button('⬇️ Download MPP Upload Template',TEMPLATE.read_bytes(),'MPP_Upload_Template.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
        with c2:
            st.markdown('### Upload MPP Data');u=st.file_uploader('Upload completed template',type=['xlsx','xls','csv']);mode=st.radio('Upload action',['Append New Tasks','Update Existing + Add New','Replace All Tasks'],horizontal=True)
            if u:
                try:
                    raw=pd.read_csv(u) if u.name.lower().endswith('.csv') else pd.read_excel(u,sheet_name=0);n=normalize_upload(raw);st.success(f'{len(n)} row(s) ready.');st.dataframe(n.head(25),use_container_width=True,height=330,hide_index=True);bad=n[n['End Date']<n['Start Date']]
                    if not bad.empty:st.error('Some rows have End Date earlier than Start Date.')
                    if st.button('Import MPP Data',type='primary',disabled=not bad.empty):
                        cur=store.read_sheet('App Tasks')
                        if mode=='Replace All Tasks':res=n
                        elif mode=='Append New Tasks':res=pd.concat([cur,n],ignore_index=True)
                        else:
                            rows={text(r['Task ID']):r for _,r in cur.iterrows()}
                            for _,r in n.iterrows():rows[text(r['Task ID'])]=r
                            res=pd.DataFrame(rows.values())
                        store.write_sheet('App Tasks',res);st.success('Imported successfully.');st.rerun()
                except Exception as ex:st.error(f'Upload could not be processed: {ex}')
    with plan:
        v=tasks.copy();v.insert(0,'Delete',False);cc={'Delete':st.column_config.CheckboxColumn('Delete'),'Task ID':st.column_config.TextColumn('Task ID',disabled=True,width='small'),'Task':st.column_config.TextColumn('Task',required=True,width='large'),'Status':st.column_config.SelectboxColumn('Status',options=statuses),'Priority':st.column_config.SelectboxColumn('Priority',options=priorities),'Start Date':st.column_config.DateColumn('Start Date',format='DD-MMM-YYYY'),'End Date':st.column_config.DateColumn('End Date',format='DD-MMM-YYYY'),'Duration Days':st.column_config.NumberColumn('Duration',disabled=True),'Progress %':st.column_config.ProgressColumn('Progress %',min_value=0,max_value=100,format='%d%%'),'Expected %':st.column_config.NumberColumn('Expected %',disabled=True),'Variance %':st.column_config.NumberColumn('Variance %',disabled=True),'Risk':st.column_config.TextColumn('Risk',disabled=True),'Comments':st.column_config.TextColumn('Comments',width='large')}
        if projects:cc['Project']=st.column_config.SelectboxColumn('Project',options=projects)
        if epics:cc['Epic']=st.column_config.SelectboxColumn('Epic',options=epics)
        if sprints:cc['Sprint']=st.column_config.SelectboxColumn('Sprint',options=sprints)
        if owners:cc['Owner']=st.column_config.SelectboxColumn('Owner',options=owners);cc['Peer QA']=st.column_config.SelectboxColumn('Peer QA',options=['']+owners)
        e=st.data_editor(v,use_container_width=True,height=620,hide_index=True,num_rows='dynamic',column_config=cc,key='mpp');work=e[~e['Delete'].fillna(False)].drop(columns=['Delete']);work=work[work['Task'].fillna('').astype(str).str.strip()!=''];bad=work[pd.to_datetime(work['End Date'],errors='coerce')<pd.to_datetime(work['Start Date'],errors='coerce')]
        st.caption('Add a row to create; Task ID auto-generates. Edit to update. Tick Delete to remove.')
        if not bad.empty:st.error('End Date cannot be earlier than Start Date.')
        if st.button('💾 Save MPP Plan',type='primary',disabled=not bad.empty):
            work=assign_task_ids(work,store.read_sheet('App Tasks'));work['Start Date']=pd.to_datetime(work['Start Date'],errors='coerce');work['End Date']=pd.to_datetime(work['End Date'],errors='coerce');work['Duration Days']=(work['End Date']-work['Start Date']).dt.days+1;work['Last Updated']=pd.Timestamp(date.today());store.write_sheet('App Tasks',work);st.success('MPP plan saved.');st.rerun()
    with gantt:
        g=tasks.dropna(subset=['Start Date','End Date'])
        if g.empty:st.info('Add dated tasks to display the Gantt.')
        else:
            fig=px.timeline(g,x_start='Start Date',x_end='End Date',y='Task',color='Status',hover_data=['Task ID','Project','Epic','Sprint','Owner','Progress %','Priority','Dependency']);fig.update_yaxes(autorange='reversed');fig.update_layout(height=max(560,len(g)*34));st.plotly_chart(fig,use_container_width=True)
elif page=='Sprint Board':
    hdr('Sprint Execution Board','Planned, active, blocked, on-hold and completed work.');sts=['Planned','In Progress','Blocked','On Hold','Complete'];cols=st.columns(5)
    for col,s in zip(cols,sts):
        with col:
            st.markdown('### '+s)
            for _,r in tasks[tasks['Status'].astype(str)==s].iterrows():
                with st.container(border=True):st.markdown(f"**{r['Task ID']} — {r['Task']}**");st.caption(f"{r['Owner']} · {r['Progress %']:.0f}% · {r['Risk']} risk")
elif page=='Resource Loading':
    hdr('Resource Loading and Capacity','Resource workload and allocation view.');st.dataframe(resources,use_container_width=True,hide_index=True)
elif page=='RAID Log':
    hdr('RAID Management','Track risks, assumptions, issues and dependencies.');e=st.data_editor(raid,use_container_width=True,height=540,hide_index=True,num_rows='dynamic');
    if st.button('💾 Save RAID Log',type='primary'):store.write_sheet('App RAID',e);st.rerun()
elif page=='Governance':
    hdr('Governance and Controls','Ceremonies, decisions and milestone gates.');st.dataframe(pd.DataFrame([['Daily Pod Stand-up','Daily'],['Scrum of Scrums','Mon/Wed/Fri'],['Steering Committee','Weekly'],['Milestone Gate','At milestone']],columns=['Ceremony','Frequency']),use_container_width=True,hide_index=True)
elif page=='Reports & Export':
    hdr('Reports and Export','Download backend, MPP template and tasks CSV.');c1,c2,c3=st.columns(3)
    with c1:st.download_button('Download Excel Backend',store.path.read_bytes(),'scrum_master_backend.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
    with c2:st.download_button('Download MPP Template',TEMPLATE.read_bytes(),'MPP_Upload_Template.xlsx','application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',use_container_width=True)
    with c3:st.download_button('Download Tasks CSV',tasks.to_csv(index=False).encode(),'mpp_tasks.csv','text/csv',use_container_width=True)