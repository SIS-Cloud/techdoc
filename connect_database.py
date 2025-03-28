import pyodbc
import pandas as pd

# Connect Database Server โดยการกำหนด Parameter ให้ถูกต้องกับการเชื่อมต่อกับ Server
con_string = "driver=SQL SERVER;server=27.254.204.72;database=Cloud;UID=cloudapi;PWD=Cloudapi1234!"

# ใช้ SQL Statement ในการ Query ข้อมูลที่เป็น Raw Data ทั้งหมด สำหรับการทำ Techdoc
sql1 = """select a.NumAtCard as CusID, a.DocNum as SO, h.DocNum as 'SQ',
a.U_m_status as Status,
CONVERT (varchar, FORMAT(GETDATE(), 'dd/MM/yyyy'), 120) as today,
isnull(b.firstName,'') + ' ' +isnull(b.lastName,'') 'Owner', e.SlpName 'Sale',
a.CardName as Partner,
a.U_m_enduser as Enduser,
FORMAT(a.U_poc_start_date, 'dd/MM/yyyy') as poc_startdate,
FORMAT(a.U_poc_end_date, 'dd/MM/yyyy') as poc_enddate,
FORMAT(a.U_prod_start_date, 'dd/MM/yyyy') as prod_startdate,
FORMAT(a.U_prod_end_date, 'dd/MM/yyyy') as prod_enddate,
--c.Name 'partner name1'
isnull(c.firstName,'') + ' ' +isnull(c.lastName,'') 'partner_name',c.E_MailL as partner_email,
isnull(c.Tel1,'') + ' ' +isnull(c.Cellolar,'') 'mobile',
a.U_m_ip,
'"'+ convert(nvarchar(max), a.U_m_password) +'"' As password,
a.U_m_accsskey,
a.U_m_secretkey,
a.U_m_wsb_rootacc
from ORDR a
left outer join OHEM b on a.OwnerCode = b.empID
left outer join OCPR c on a.CntctCode = c.CntctCode and a.CardCode = c.CardCode
left outer join (select a.CardCode,a.name,a.E_MailL from OCPR a)d on a.CardCode = d.CardCode and a.U_m_technician_partner = d.Name
left outer join OSLP e on a.SlpCode = e.SlpCode
left outer join (select a.CardCode,a.name,a.E_MailL from OCPR a)f on a.CardCode = f.CardCode and a.U_m_technician_partner2 = f.Name
left outer join (select a.CardCode,a.name,a.E_MailL from OCPR a)g on a.CardCode = g.CardCode and a.U_m_technician_partner3 = g.Name
left outer join (select distinct h1.DocEntry,h2.DocNum from RDR1 h1
left outer join OQUT h2 on h1.BaseType = h2.ObjType and h1.BaseRef = h2.DocNum)h on a.DocEntry = h.DocEntry
where
a.DocStatus = 'O'
and a.U_m_status not in ('DEL', 'DIS')
order by
a.U_m_status,
a.NumAtCard"""

# ใช้ SQL Statement ในการ Query ข้อมูลสำหรับการทำตาราง Item Description ใน Techdoc
sql2 = """select
a.NumAtCard 'Tenant No.',d.Dscription, d.U_m_sizing
from Cloud.dbo.ORDR a
left outer join Cloud.dbo.RDR1 d on a.DocEntry = d.DocEntry
left outer join Cloud.dbo.OHEM b on a.OwnerCode = b.empID
left outer join Cloud.dbo.OSLP c on a.SlpCode = c.SlpCode
left outer join Cloud.dbo.[@TNNO] e on a.NumAtCard = e.Code
where a.DocStatus = 'O'
and a.U_m_status in ('PROD','POC','INT')
order by a.U_m_status, a.NumAtCard"""


with pyodbc.connect(con_string, autocommit= True) as con:
    con.execute(sql1)

    df = pd.read_sql_query(sql1, con)
    df.to_excel('query_info.xlsx',index=False,sheet_name='Sheet1')

    con.execute(sql2)

    df = pd.read_sql_query(sql2, con)
    df.to_excel('query_item.xlsx',index=False,sheet_name='Sheet1')

    con.close()


