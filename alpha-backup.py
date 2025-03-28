from flask import Flask, render_template, request, session, redirect, url_for
from flask_wtf import FlaskForm
from wtforms import StringField, SubmitField, SelectField
import pyodbc
import pandas as pd
from docxtpl import DocxTemplate
from docx.shared import Pt
from datetime import datetime


app = Flask(__name__)
app.config['SECRET_KEY'] = ',kp8upN'
# กำหนด Secret Key สำหรับการทำ Security ที่เกี่ยวข้องกับ CSRF Protection

class MyForm(FlaskForm):
# สร้าง class ชื่อ MyForm สำหรับสร้าง Form บนหน้าเว็บ
    tenantid = StringField("ป้อนเลข Tenant ID")
    # สร้างตัวแปร Tenantid สำหรับสร้างช่องกรอกข้อมูลบน Form
    submit = SubmitField("Submit")
    # สร้างตัวแปร submit สำหรับสร้างปุ่มกด submit บน Form
    # checkbox = BooleanField("Checkbox")
    # radio_template = RadioField("Template Techdoc",
    #                               choices=[('1','1101-XXXX S3 Storage'),
    #                                        ('2','1102-XXXX Veeam Agent')])
    select_template = SelectField("เลือก Template Techdoc",
                                  choices=[('Default','SiS Template Techdoc'),
                                           ('1','1101-XXXX S3 Storage'),
                                           ('2','1102-XXXX Veeam Agent'),
                                           ('3','1102-XXXX Veeam Backup Standard'),
                                           ('4','1103-XXXX Vmware IaaS'),
                                           ('5','1103-XXXXDR Vmware DRaaS (VCC01)'),
                                           ('6','1103-XXXXDR Vmware DRaaS (VCC04)'),
                                           ('7','1105-XXXX Nutanix IaaS'),
                                           ('8','1107-XXXX Magicbox'),
                                           ('9','1401-XXXX Wasabi Storage'),
                                           ('10','1002-XXXX Veeam Agent IDC'),
                                           ('11','1002-XXXX Veeam Backup Standard IDC'),
                                           ('12','1003-XXXX Vmware IaaS IDC'),
                                           ('13','1003-XXXXDR Vmware DRaaS IDC'),
                                           ('14','1005-XXXX Nutanix IaaS IDC')])

@app.route('/',methods=['get','post'])
# สร้าง route ไปยังหน้า index
def index():
    try:
        form = MyForm()
    # สร้างตัวแปร form เท่ากับ class MyForm

        if form.validate_on_submit():
        # สร้างเงื่อนไข ถ้ามีการกด Submit บน Form ให้ทำอย่างไร
            session['tenantid'] = form.tenantid.data
            session['select_template'] = form.select_template.data
            # session['checkbox'] = form.checkbox.data
            # session['radio_template'] = form.radio_template.data
            # ดึงข้อมูลที่อยู่ใน Field แต่ละแบบ จาก class MyForm มาใช้งาน
            # ใช้ session ในการกระจายค่าตัวแปรต่างๆ ไปยังแต่ละหน้าเว็บที่ต้องการเรียกใช้งาน

            form.tenantid.data = ""
            form.select_template.data = ""

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

            # อ่านไฟล์ Excel "query_info" กำหนดให้คอลั่ม CusID เป็นคอลั่ม Index
            df = pd.read_excel("query_info.xlsx",sheet_name="Sheet1",index_col="CusID")

            # ใช้ While Loop กำหนดให้เป็นจริงเสมอจนกว่าจะ Break

            # Input ค่าเลข Tenant และอ่านค่า Dataframe ที่ตำแหน่ง Index = X (เลข Tenant)
            x = session['tenantid']
            test = df.loc[[x]]
            print(test)


            # เก็บค่าแต่ละ Cell ตามคอลั่มที่กำหนดของแถว Index = X ลงในตัวแปรต่าง ๆ
            today_date = datetime.today().strftime("%d.%m.%y") # เปลี่ยน Format Datetime ตามที่ต้องการ
            tenantid = x
            status = df.loc[[x],'Status'].values[0]
            sonumber = df.loc[[x],'SO'].values[0]
            sqnumber = df.loc[[x],'SQ'].values[0]
            pocstartdate = df.loc[[x],'poc_startdate'].values[0]
            pocenddate = df.loc[[x],'poc_enddate'].values[0]
            prodstartdate = df.loc[[x],'prod_startdate'].values[0]
            prodenddate = df.loc[[x],'prod_enddate'].values[0]
            partnercom = df.loc[[x],'Partner'].values[0]
            partnername = df.loc[[x],'partner_name'].values[0]
            partneremail = df.loc[[x],'partner_email'].values[0]
            partnerphone = df.loc[[x],'mobile'].values[0]
            endusercom = df.loc[[x],'Enduser'].values[0]
            password = df.loc[[x],'password'].values[0].strip(f'"') # ทำการ Trim ตัวอักษรหัวท้ายที่มีเครื่องหมาย "
            publicip = df.loc[[x],'U_m_ip'].values[0]
            accesskey = df.loc[[x],'U_m_accsskey'].values[0]
            secretkey = df.loc[[x],'U_m_secretkey'].values[0]
            wsbrootacc = df.loc[[x],'U_m_wsb_rootacc'].values[0]


            # เปลี่ยนข้อมูล String เป็น Format Datetime ตามที่ต้องการ
            if status == 'POC':
                pocstartdate = pd.to_datetime(df.loc[[x],'poc_startdate'].values[0]).strftime("%d.%m.%y")
                pocenddate = pd.to_datetime(df.loc[[x],'poc_enddate'].values[0]).strftime("%d.%m.%y")
            elif status == 'PROD':
                prodstartdate = pd.to_datetime(df.loc[[x],'prod_startdate'].values[0]).strftime("%d.%m.%y")
                prodenddate = pd.to_datetime(df.loc[[x],'prod_enddate'].values[0]).strftime("%d.%m.%y")
            elif status == 'Int':
                pocstartdate = pd.to_datetime(df.loc[[x],'poc_startdate'].values[0]).strftime("%d.%m.%y")
                pocenddate = pd.to_datetime(df.loc[[x],'poc_enddate'].values[0]).strftime("%d.%m.%y")

            # ทำ Dictionary บน Techdoc กับตัวแปรที่เก็บข้อมูลจาก Dataframe
            context = {'tenantid': x,
                'today_date': today_date,
                'status': status,
                'sonumber': sonumber,
                'sqnumber': sqnumber,
                'poc_startdate': pocstartdate,
                'poc_enddate': pocenddate,
                'prod_startdate': prodstartdate,
                'prod_enddate': prodenddate,
                'partner_com': partnercom,
                'partner_name': partnername,
                'partner_email': partneremail,
                'partner_mobile': partnerphone,
                'enduser_com' : endusercom,
                'password': password,
                'publicip': publicip,
                'accesskey': accesskey,
                'secretkey': secretkey,
                'wsbrootacc': wsbrootacc
                }
            print(context)

            def format_template():
                if status == 'POC':
                    techdoc = DocxTemplate(f"SiS Cloud Techdoc {y} poc.docx")
                elif status == 'PROD':
                    techdoc = DocxTemplate(f"SiS Cloud Techdoc {y} prod.docx")
                else:
                    techdoc = DocxTemplate(f"SiS Cloud Techdoc {y} poc.docx")
                context.update(context)
                techdoc.render(context)

                df = pd.read_excel("query_item.xlsx",sheet_name="Sheet1")

                    # เปลี่ยนชื่อคอลัมน์
                df = df.rename(columns={"Dscription": "รายละเอียดสินค้า", "U_m_sizing": "หน่วย"})

                # เปลี่ยน format style font ในไฟล์ word
                style = techdoc.styles['Normal']
                font = style.font
                font.name = 'Cordia New'
                font.size = Pt(11)

                # สร้างตารางใน Word โดยกำหนดจำนวนคอลัมน์เป็น 2 (สำหรับคอลัมน์ที่ 2 และ 3)
                table = techdoc.add_table(rows=1, cols=2)
                hdr_cells = table.rows[0].cells

                # กำหนดหัวตารางจากชื่อคอลัมน์ที่ 2 และ 3 (ปรับตามชื่อจริงในไฟล์ Excel)
                hdr_cells[0].text = str(df.columns[1])
                hdr_cells[1].text = str(df.columns[2])

                # ตั้งค่าข้อมูลตารางทำเป็น Grid
                table.style = 'TableGrid'

                # ตั้งค่าให้ตารางทำ AutoFit
                table.autofit = True
                table.allow_autofit = True

                    # วนลูปข้อมูลใน DataFrame และคัดลอกเฉพาะแถวที่คอลัมน์แรกมีค่าเป็น ตัวแปร x (เก็บค่าเลข Tenant)
                for index, row in df.iterrows():
                    if row[df.columns[0]] == x:
                        new_row_cells = table.add_row().cells
                        new_row_cells[0].text = str(row[df.columns[1]])
                        new_row_cells[1].text = str(row[df.columns[2]])
                # techdoc.save(f"SiS Cloud Techdoc {x}.docx")

                # Save File โดยระบุ Path ไปที่ Folder Downlaod และ Format ชื่อตามเลข Tenant
                techdoc.save("C:/Users/Administrator/Downloads/" + f"SiS Cloud Techdoc {x}.docx")

            if session['select_template'] == '1':
                print("1101 Template : " + session['select_template'])
                y = "1101-xxxx"
                format_template()

            elif session['select_template'] == '2':
                print("1102 Template : " + session['select_template'])
                y = "1102-XXXX (Veeam Agent)"
                format_template()

            elif session['select_template'] == '3':
                print("1102 Template : " + session['select_template'])
                y = "1102-XXXX (Veeam Backup Standard)"
                format_template()

            elif session['select_template']== '4':
                print("1103 Template : " + session['select_template'])
                y = "1103-xxxx"
                format_template()

            elif session['select_template'] == '5':
                print("1103DR Template : " + session['select_template'])
                y = "1103-XXXXDR (vcc01)"
                format_template()

            elif session['select_template'] == '6':
                print("1103DR Template : " + session['select_template'])
                y = "1103-XXXXDR (vcc04)"
                format_template()

            elif session['select_template'] == '7':
                print("1105 Template : " + session['select_template'])
                y = "1105-xxxx"
                format_template()

            elif session['select_template'] == '8':
                print("1107 Template : " + session['select_template'])
                y = "1107-xxxx"
                format_template()

            elif session['select_template'] == '9':
                print("1401 Template : " + session['select_template'])
                y = "1401-xxxx"
                format_template()

            elif session['select_template'] == '10':
                print("1001 Template : " + session['select_template'])
                y = "1001-xxxx"
                format_template()
    except KeyError:
    # กำหนดให้ยกเว้น Code Error เมื่อเจอ KeyError และแสดงข้อความ "No Tenant ID แทน
        print("No Tenant ID : " + x)
        return render_template('test.html',form=form)



app.route('/about')
def about():
        if request.method == 'POST':
        # รับข้อมูลจากฟอร์ม
            session['tenantid'] = request.form[session['tenantid']]
            return render_template('success.html', x=session['tenantid'])
        return render_template("test2.html")


if __name__ == "__main__":
    app.run(debug=True,use_reloader=True)
    # สั่ง app run โดยเปิด debug mode บนหน้าเว็บและทำการ auto reload ทุกครั้งที่กด save
