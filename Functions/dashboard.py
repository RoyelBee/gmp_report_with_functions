
from PIL import Image, ImageFont, ImageDraw
import pyodbc
import pandas as pd
from calendar import monthrange
from datetime import date, datetime
import sys
import path as dir

conn = pyodbc.connect('DRIVER={SQL Server};'
                      'SERVER=137.116.139.217;'
                      'DATABASE=ARCHIVESKF;'
                      'UID=sa;'
                      'PWD=erp@123;')
# Total SKU
def dash_kpi_generator(name):
    total_sku = pd.read_sql_query(
        """select gpmname,count(distinct BRAND) as 'total brand',count(itemno) as 'total_SKU' from PRINFOSKF
       where status=1
       and gpmname like ?
       group by gpmname """, conn, params={name})
    # Total SKU List
    total_sku_list = total_sku['total_SKU'].to_list()
    # Total brand list
    total_brand_list = total_sku['total brand'].tolist()
    # Sold SKU
    sold_sku = pd.read_sql_query(
        """select count(distinct item) as 'Sold_SKU' from OESalesDetails
       where item in(select itemno from PRINFOSKF
       where status=1
       and gpmname like ? )
       and transtype = 1
       and left(transdate,6)=CONVERT(varchar(6), dateAdd(day,0,getdate()), 112)""", conn, params={name})
    # Sold SKU List
    sold_sku_list = sold_sku['Sold_SKU'].to_list()
    # No sale SKU
    no_sales_sku = total_sku_list[0] - sold_sku_list[0]
    # No Stock SKU
    no_stock_sku = pd.read_sql_query("""select count(a.itemno) 'no stock item' from
                   (select itemno from PRINFOSKF
                   where status=1
                   and gpmname like ? ) as a
                   left join
                   (select itemno,isnull(sum(QTYONHAND),0) as stock from ICStockStatusCurrentLOT
                   group by itemno) as b
                   on a.itemno = b.itemno
                   where stock = 0 """, conn, params={name})
    # No Stock SKU list
    no_stock_sku_list = no_stock_sku['no stock item'].to_list()
    # Read excel file
    read_file_for_all_data = pd.read_excel('./Data/gpm_data.xlsx')
    # Target
    total_target = read_file_for_all_data['MTD Sales Target'].to_list()
    total_sales = read_file_for_all_data['Actual Sales MTD'].to_list()
    # Achivement

    if sum(total_target) == 0:
        achivemet = 0
    else:
        achivemet = (sum(total_sales) / sum(total_target)) * 100
    # Trend
    today = str(date.today())
    datee = datetime.strptime(today, "%Y-%m-%d")
    date_val = int(datee.day)
    num_days = monthrange(datee.year, datee.month)[1]
    trend = (sum(total_sales) / date_val) * num_days
    # Trend Achivement
    trend_achivement = (sum(total_sales) / trend) * 100

    def currency_converter(num):
        num_size = len(str(num))
        if num_size >= 8:
            number = str(round((num / 10000000), 2)) + 'Cr'
        elif num_size >= 7:
            number = str(round(num / 1000000, 2)) + 'M'
        elif num_size >= 4:
            number = str(int(num / 1000)) + 'K'
        else:
            number = num
        return number

    # Open Image file

    image = Image.open(dir.get_directory() + "/images/dash_kpi.png")
    print('Image read')
    draw = ImageDraw.Draw(image)
    font = ImageFont.truetype(dir.get_directory() + '/Images/FrancoisOne-Regular.ttf', 30)
    # Draw on the image
    print('Font loaded ')
    draw.text((80, 80), str(total_brand_list[0]), font=font, fill=(39, 98, 236))
    draw.text((270, 80), str(total_sku_list[0]), font=font, fill=(39, 98, 236))
    draw.text((460, 80), str(sold_sku_list[0]), font=font, fill=(39, 98, 236))
    draw.text((650, 80), str(no_sales_sku), font=font, fill=(39, 98, 236))
    draw.text((840, 80), str(no_stock_sku_list[0]), font=font, fill=(39, 98, 236))
    #
    draw.text((70, 220), str(currency_converter(sum(total_target))), font=font, fill=(255, 255, 255))
    draw.text((260, 220), str(currency_converter(sum(total_sales))), font=font, fill=(255, 255, 255))
    draw.text((435, 220), str(round(achivemet, 1)) + '%', font=font, fill=(255, 255, 255))
    draw.text((625, 220), str(currency_converter(round(trend, 1))), font=font, fill=(255, 255, 255))
    draw.text((830, 220), str(round(trend_achivement, 1)) + '%', font=font, fill=(255, 255, 255))
    # image.show()
    image.save('./Images/dashboard.png')
    print('Dash generated')



