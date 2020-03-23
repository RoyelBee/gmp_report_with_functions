from PIL import Image, ImageDraw, ImageFont
import pytz
import datetime as dd
from PIL import Image
from datetime import datetime

import path as dir
print(dir.get_directory())

def banner():
    date = datetime.today()
    day = str(date.day) + '/' + str(date.month) + '/' + str(date.year)
    tz_NY = pytz.timezone('Asia/Dhaka')
    datetime_BD = datetime.now(tz_NY)
    time = datetime_BD.strftime("%I:%M %p")
    date = datetime.today()
    x = dd.datetime.now()
    day = str(date.day) + '-' + str(x.strftime("%b")) + '-' + str(date.year)
    tz_NY = pytz.timezone('Asia/Dhaka')
    datetime_BD = datetime.now(tz_NY)
    time = datetime_BD.strftime("%I:%M %p")

    img = Image.open(dir.get_directory() + "/images/new_ai.png")
    title = ImageDraw.Draw(img)
    timestore = ImageDraw.Draw(img)
    tag = ImageDraw.Draw(img)
    branch = ImageDraw.Draw(img)
    font = ImageFont.truetype(dir.get_directory() + "/images/Stencil_Regular.ttf", 40, encoding="unic")
    font1 = ImageFont.truetype(dir.get_directory() + "/images/ROCK.ttf", 30, encoding="unic")
    font2 = ImageFont.truetype(dir.get_directory() + "/images/ROCK.ttf", 22, encoding="unic")
    report_name = 'GPM '
    #
    tag.text((25, 8), 'SK+F', (255, 255, 255), font=font)
    branch.text((25, 130), report_name + " Stock Report", (255, 209, 0), font=font1)
    timestore.text((25, 178), time + "\n" + day, (255, 255, 255), font=font2)
    img.save(dir.get_directory() + "/images/banner_ai.png")
    # img.show()
    print("Banner Created")