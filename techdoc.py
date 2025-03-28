import pandas as pd
from docxtpl import DocxTemplate
from docx.shared import Pt
from datetime import datetime

# สร้าง function format_template ในการสร้าง Techdoc จาก MS. Word
def format_template():
    if status == 'POC':
        # techdoc = DocxTemplate(f"SiS Cloud Techdoc {y} poc.docx")
        techdoc = DocxTemplate("C:/Python Project/Techdoc/Files MS Word Template/" + f"SiS Cloud Techdoc {y} poc.docx")
    elif status == 'PROD':
        # techdoc = DocxTemplate(f"SiS Cloud Techdoc {y} prod.docx")
        techdoc = DocxTemplate("C:/Python Project/Techdoc/Files MS Word Template/" + f"SiS Cloud Techdoc {y} prod.docx")
    else:
        # techdoc = DocxTemplate(f"SiS Cloud Techdoc {y} poc.docx")
        techdoc = DocxTemplate("C:/Python Project/Techdoc/Files MS Word Template/" + f"SiS Cloud Techdoc {y} poc.docx")
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

# อ่านไฟล์ Excel "query_info" กำหนดให้คอลั่ม CusID เป็นคอลั่ม Index
df = pd.read_excel("query_info.xlsx",sheet_name="Sheet1",index_col="CusID")

# ใช้ While Loop กำหนดให้เป็นจริงเสมอจนกว่าจะ Break
while (True):
    try:
        # Input ค่าเลข Tenant และอ่านค่า Dataframe ที่ตำแหน่ง Index = X (เลข Tenant)
        print('Enter Tenant ID:')
        x = input()
        test = df.loc[[x]]
        print(test)

        # Input Template Techdoc ที่ต้องการ
        print(f"Template Techdoc :\n\
              1 > 1101-XXXX S3 Storage\n\
              2 > 1102-XXXX Veeam Agent\n\
              3 > 1102-XXXX Veeam Backup Standard\n\
              4 > 1103-XXXX Vmware IaaS\n\
              5 > 1103-XXXXDR Vmware DRaaS (VCC01)\n\
              6 > 1103-XXXXDR Vmware DRaaS (VCC04)\n\
              7 > 1105-XXXX Nutanix IaaS\n\
              8 > 1107-XXXX Magicbox\n\
              9 > 1401-XXXX Wasabi Storage\n\
              10 > 1002-XXXX Veeam Agent IDC\n\
              11 > 1002-XXXX Veeam Backup Standard IDC\n\
              12 > 1003-XXXX Vmware IaaS IDC\n\
              13 > 1003-XXXXDR Vmware DRaaS IDC\n\
              14 > 1005-XXXX Nutanix IaaS IDC")
        template = input("Choose Number 1-14 : ")

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

        # สร้างเงื่อนไขจาก Input เลข Template เรียกใช้งาน function format_template
        if template == '1':
            print("1101 Status : " + status)
            y = "1101-xxxx"
            format_template()
            break

        elif template == '2':
            print("1102 Status : " + status)
            y = "1102-XXXX (Veeam Agent)"
            format_template()
            break

        elif template == '3':
            print("1102 Status : " + status)
            y = "1102-XXXX (Veeam Backup Standard)"
            format_template()
            break

        elif template == '4':
            print("1103 Status : " + status)
            y = "1103-xxxx"
            format_template()
            break

        elif template == '5':
            print("1103DR Status : " + status)
            y = "1103-XXXXDR (vcc01)"
            format_template()
            break

        elif template == '6':
            print("1103DR Status : " + status)
            y = "1103-XXXXDR (vcc04)"
            format_template()
            break

        elif template == '7':
            print("1105 Status : " + status)
            y = "1105-xxxx"
            format_template()
            break

        elif template == '8':
            print("1107 Status : " + status)
            y = "1107-xxxx"
            format_template()
            break

        elif template == '9':
            print("1401 Status : " + status)
            y = "1401-xxxx"
            format_template()
            break

        elif template == '10':
            print("1001 Status : " + status)
            y = "1001-xxxx"
            format_template()
            break


        else:
            print("Input Invalid Number Template : " + template)
            pass

    # กำหนดให้ยกเว้น Code Error เมื่อเจอ KeyError และแสดงข้อความ "No Tenant ID แทน
    except KeyError:
        print("No Tenant ID : " + x)

