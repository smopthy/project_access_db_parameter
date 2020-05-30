# Project_work_获取數據与导出excel 1.0 

`import pandas as pd 
import prodbc
`

### 首先建立与数据库连接
* 假设合作单位的数据库为本地連接

`conn = pyodbc.connect(
    'Driver={SQL Server Native Client 10.0};'
    'Server=.;'
    'Database=gshm_jx;'
    'Trusted_Connection=yes;'
)
`

### 获取含有 SQL QUERY 的表单,合并

`month = '2020-04'
data_month ='and data_month =\'month\''
df_parameter['para_data']=df_parameter['remark']+'and data_month ='+'\'2020-04\''
df_parameter = df_parameter.dropna(axis = 0, subset = ['remark'] )
`

### iterate execute SQL QUERY 并把数据结果返回成pd.DataFrame 

`df= pd.DataFrame() 
listcolum = []
df_parameter=df_parameter.reset_index(drop=True)

for para in df_parameter['para_data'] : 
        a = pd.read_sql(para , conn)
        df= df.append(a).reset_index(drop=True)

df.columns = ['para_value']
`

### 合并 

`
df_parameter = pd.merge(df_parameter,df,left_index=True, right_index=True, how='outer')
`
### 导出excel 档 

`
with pd.ExcelWriter(r'C:\path\filename.xlsx') as writer:
    df_parameter.to_excel(writer, sheet_name='df1')
    df_parameter.to_excel(writer, sheet_name='df2')
`



