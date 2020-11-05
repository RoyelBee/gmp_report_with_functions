import datetime
import os
import smtplib
from _datetime import datetime
from email import encoders
from email.mime.base import MIMEBase
from email.mime.image import MIMEImage
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

import Functions.design_report_layout as layout
import pandas as pd
import numpy as np
import path as d
import xlsxwriter

def send_mail(gpm_name):
    import Functions.banner_code as ban
    import Functions.generate_data as gdata
    import Functions.dashboard as dash
    import Functions.cumulative_target_sales as cm
    import Functions.executive_wise_sales_target as ex
    import Functions.brand_wise_target_sales as b

    ban.banner()
    # gdata.GenerateReport(gpm_name)
    # dash.dash_kpi_generator(gpm_name)
    # cm.cumulative_target_sales(gpm_name)
    # ex.executive_sales_target(gpm_name)
    # b.brand_wise_target_sales()
    #
    # # # --------- Add Image Border ---------------------------------------
    from PIL import Image
    # da = Image.open("./Images/dashboard.png")
    # imageSize = Image.new('RGB', (962, 297))
    # imageSize.paste(da, (1, 0))
    # imageSize.save("./Images/dashboard.png")
    #
    # kpi1 = Image.open("./Images/Cumulative_Day_Wise_Target_vs_Sales.png")
    # imageSize = Image.new('RGB', (962, 481))
    # imageSize.paste(kpi1, (1, 0))
    # imageSize.save("./Images/Cumulative_Day_Wise_Target_vs_Sales.png")
    #
    # kpi2 = Image.open("./Images/executive_wise_target_vs_sold_quantity.png")
    # imageSize = Image.new('RGB', (962, 481))
    # imageSize.paste(kpi2, (1, 0))
    # imageSize.save("./Images/executive_wise_target_vs_sold_quantity.png")
    # #
    # kpi3 = Image.open("./Images/brand_wise_target_vs_sold_quantity.png")
    # imageSize = Image.new('RGB', (962, 481))
    # imageSize.paste(kpi3, (1, 0))
    # imageSize.save("./Images/brand_wise_target_vs_sold_quantity.png")

    # # -------------------------------------------------------------------
    # # ------ Branch wise stock summery -----------------------------------

    import sys

    # df = pd.read_excel('Data/gpm_data.xlsx')
    #
    # cols = range(52, 83)
    # df2 = pd.read_excel('Data/gpm_data.xlsx', usecols=cols)
    # l = list(df2.columns)
    #
    # cols2 = range(19, 50)
    # df3 = pd.read_excel('Data/gpm_data.xlsx', usecols=cols2)
    # l2 = list(df3.columns)
    #
    # final_list = []
    #
    # for branches, avgq in zip(l, l2):
    #     branch_current_stock = df[branches].to_list()
    #
    #     branch_avg_sell = df[avgq].to_list()
    #
    #     for each, each2 in zip(branch_current_stock, branch_avg_sell):
    #         if each == 0:
    #             index = branch_current_stock.index(each)
    #             branch_current_stock[index] = 1
    #         if each2 == 0:
    #             index2 = branch_avg_sell.index(each2)
    #             branch_avg_sell[index2] = 1
    #
    #     np.seterr(divide='ignore', invalid='ignore')
    #     branch_stock_limit = np.divide(branch_current_stock, branch_avg_sell)
    #     branch_stock_limit_list = list(branch_stock_limit)
    #
    #     new_list_nill = []
    #     new_list_super_under_stock = []
    #     new_list_under_stock = []
    #     new_list_normal_stock = []
    #     new_list_over_stock = []
    #     new_list_super_over_stock = []
    #
    #     for i in branch_stock_limit_list:
    #         # Color for NIL
    #         if i <= 0:
    #             new_list_nill.append(i)
    #         # Super Under Stock
    #         elif i <= 15:
    #             new_list_super_under_stock.append(i)
    #
    #         # Under Stock
    #         elif i <= 35:
    #             new_list_under_stock.append(i)
    #
    #         # Color for Normal
    #         elif i <= 45:
    #             new_list_normal_stock.append(i)
    #
    #         # Color for Over Stock
    #         elif i <= 60:
    #             new_list_over_stock.append(i)
    #         else:
    #             # Super over stock
    #             new_list_super_over_stock.append(i)
    #
    #     sum_new_list_nill = len(new_list_nill)
    #     sum_new_list_super_under_stock = len(new_list_super_under_stock)
    #     sum_new_list_under_stock = len(new_list_under_stock)
    #     sum_new_list_normal_stock = len(new_list_normal_stock)
    #     sum_new_list_over_stock = len(new_list_over_stock)
    #     sum_new_list_super_over_stock = len(new_list_super_over_stock)
    #
    #     sum_list = [sum_new_list_nill, sum_new_list_super_under_stock, sum_new_list_under_stock,
    #                 sum_new_list_normal_stock, sum_new_list_over_stock, sum_new_list_super_over_stock]
    #     final_list.append(sum_list)
    #
    # workbook = xlsxwriter.Workbook('Data/branch_wise_nil_us_ss.xlsx')
    # worksheet = workbook.add_worksheet()
    #
    # len = range(1, len(l) + 1)
    #
    # for i, j in zip(l, len):
    #     worksheet.write(j, 0, i)
    #
    # stock_list = ['Branch', 'Nill', 'Super Under Stock', 'Under Stock', 'Normal Stock', 'Over Stock',
    #               'Super Over Stock']
    #
    # len2 = range(0, 7)
    # for i, j in zip(stock_list, len2):
    #     worksheet.write(0, j, i)
    #
    # # Branch wise value set
    # all_row_length = range(0, 31)
    # all_len = range(0, 6)
    #
    # for a in all_row_length:
    #     for b in all_len:
    #         c = final_list[a][b]
    #         # print(c)
    #         worksheet.write(a + 1, b + 1, c)
    # workbook.close()
    # print('Branch wise stock summery generated')

    # # ------------- HTML generating section ------------------------------
    data = pd.read_excel('./Data/html_data_Sales_and_Stock.xlsx')

    if data.empty:
        print('No data for print')
    else:
        print('Layout Generating')
        layout.generate_layout(gpm_name)

    # to = gpm.getGPMEmail(gpm_name)
    #
    # if (to == ['nawajesh@skf.transcombd.com', '']):
    #     to = ['rejaul.islam@transcombd.com', '']
    #     cc = ['', '']
    #     bcc = ['', '']
    #     print('Report Sending to = ', to)

    to = ['rejaul.islam@transcombd.com', '']
    cc = ['', '']
    bcc = ['', '']

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
    fp = open(d.get_directory() + '/images/Cumulative_Day_Wise_Target_vs_Sales.png', 'rb')
    cm = MIMEImage(fp.read())
    fp.close()
    cm.add_header('Content-ID', '<cm>')
    msgRoot.attach(cm)

    # --- Read Cumulative Target & Sales Images
    fp = open(d.get_directory() + '/images/executive_wise_target_vs_sold_quantity.png', 'rb')
    executive = MIMEImage(fp.read())
    fp.close()
    executive.add_header('Content-ID', '<executive>')
    msgRoot.attach(executive)

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
    #
    # # Html_file = open("testinghtml.html", "w")
    # # Html_file.write(layout.generate_layout(gpm_name))
    # # Html_file.close()
