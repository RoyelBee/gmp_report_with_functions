import datetime
import os
import smtplib
from _datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import pandas as pd
import Functions.design_report_layout as layout
import Functions.read_gpm_info as gpm
import path as d


def send_mail(gpm_name):
    import Functions.banner_code as b
    # b.banner()

    import Functions.generate_data as gdata
    # gdata.GenerateReport(gpm_name)

    import Functions.dashboard as dash
    # dash.dash_kpi_generator(gpm_name)

    import Functions.cumulative_target_sales as cm
    # cm.cumulative_target_sales(gpm_name)

    import Functions.executive_wise_sales_target as ex
    ex.executive_sales_target(gpm_name)

    import Functions.brand_wise_target_sales as b
    # b.brand_wise_target_sales()

    # # --------- Join Executive and Cumulative
    from PIL import Image
    kpi1 = Image.open("./Images/executive_wise_target_vs_sold_quantity.png")
    kpi2 = Image.open("./Images/Cumulative_Day_Wise_Target_vs_Sales.png")

    imageSize = Image.new('RGB', (960, 480))  # 2110
    imageSize.paste(kpi1, (0, 0))
    imageSize.paste(kpi2, (401, 0))

    imageSize.save("./Images/marge_2_into_1.png")

    data = pd.read_excel('./Data/html_data_Sales_and_Stock.xlsx')

    if data.empty:
        print('No data for print')
    else:
        print('Else working')
        layout.generate_layout(gpm_name)

    to = gpm.getGPMEmail(gpm_name)

    if (to == ['nawajesh@skf.transcombd.com', '']):
        to = ['rejaul.islam@transcombd.com', '']
        cc = ['', '']
        bcc = ['', '']
        print('Report Sending to = ', to)

    to = to
    cc = cc
    bcc = bcc

    msgRoot = MIMEMultipart('related')
    me = 'erp-bi.service@transcombd.com'
    # to = to
    # cc = ['biswascma@yahoo.com', 'yakub@transcombd.com', 'zubair.transcom@gmail.com']
    # bcc = ['rejaul.islam@transcombd.com', 'aftab.uddin@transcombd.com', 'fazle.rabby@transcombd.com']

    recipient = to + cc + bcc

    date = datetime.today()
    today = str(date.day) + '/' + str(date.month) + '/' + str(date.year) + ' ' + date.strftime("%I:%M %p")

    # # ------------ Group email --------------------
    subject = "Brand Wise Sales and Stock Report " + today
    email_server_host = 'mail.transcombd.com'
    port = 25

    msgRoot['From'] = me
    msgRoot['To'] = ', '.join(to)
    msgRoot['Cc'] = ', '.join(cc)
    msgRoot['Bcc'] = ', '.join(bcc)
    msgRoot['Subject'] = subject

    # msgRoot['to'] = recipient
    # msgRoot['from'] = me
    # msgRoot['subject'] = subject

    msgAlternative = MIMEMultipart('alternative')
    msgRoot.attach(msgAlternative)
    msgText = MIMEText('This is the alternative plain text message.')
    msgAlternative.attach(msgText)

    # # We reference the image in the IMG SRC attribute by the ID we give it below
    msgText = MIMEText("""
                            """ + layout.generate_layout(gpm_name) + """
                        """, 'html')
    msgAlternative.attach(msgText)

    # --- Read Banner Images
    fp = open(d.get_directory() + '/images/banner_ai.png', 'rb')
    banner_ai = MIMEImage(fp.read())
    fp.close()
    banner_ai.add_header('Content-ID', '<banner_ai>')
    msgRoot.attach(banner_ai)

    # --- Read Dashboard KPI Images
    fp = open(d.get_directory() + '/images/dashboard.png', 'rb')
    dash = MIMEImage(fp.read())
    fp.close()
    dash.add_header('Content-ID', '<dash>')
    msgRoot.attach(dash)

    # --- Read Cumulative Target & Sales Images
    fp = open(d.get_directory() + '/images/marge_2_into_1.png', 'rb')
    cm = MIMEImage(fp.read())
    fp.close()
    cm.add_header('Content-ID', '<cm>')
    msgRoot.attach(cm)

    # --- Read Cumulative Target & Sales Images
    fp = open(d.get_directory() + '/images/brand_wise_target_vs_sold_quantity.png', 'rb')
    brand = MIMEImage(fp.read())
    fp.close()
    brand.add_header('Content-ID', '<brand>')
    msgRoot.attach(brand)

    # Add GPM sales and stock dataset
    part = MIMEBase('application', "octet-stream")
    file_location = d.get_directory() + '/Data/Sales_and_Stock.xlsx'
    filename = os.path.basename(file_location)
    attachment = open(file_location, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msgRoot.attach(part)

    # Add GPM No Sales dataset
    part = MIMEBase('application', "octet-stream")
    file_location = d.get_directory() + '/Data/NoSales.xlsx'
    filename = os.path.basename(file_location)
    attachment = open(file_location, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msgRoot.attach(part)

    # Add GPM No Stock dataset
    part = MIMEBase('application', "octet-stream")
    file_location = d.get_directory() + '/Data/NoStock.xlsx'
    filename = os.path.basename(file_location)
    attachment = open(file_location, "rb")
    part = MIMEBase('application', 'octet-stream')
    part.set_payload(attachment.read())
    encoders.encode_base64(part)
    part.add_header('Content-Disposition', "attachment; filename= %s" % filename)
    msgRoot.attach(part)

    # #----------- Finally send mail and close server connection -----
    server = smtplib.SMTP(email_server_host, port)
    server.ehlo()
    print('\n-----------------')
    print('Sending Mail')
    server.sendmail(me, recipient, msgRoot.as_string())
    print('Mail Send')
    print('-------------------')
    server.close()

    # Html_file = open("testinghtml.html", "w")
    # Html_file.write(layout.generate_layout(gpm_name))
    # Html_file.close()
