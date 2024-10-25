
import pandas as pd
from PIL import Image, ImageFont, ImageDraw
import os


FONT_FILE = ImageFont.truetype('ShipporiMincho-Medium.ttf', 36)#chage the font and size here
FONT_COLOR = "#000000"
WIDTH, HEIGHT  = 528,448 # define certificate actual length/2 and bridth/2 or point where you wanna start writing

def make_cert(name):
    os.makedirs('out', exist_ok=True)
    image_source = Image.open('2.png')
    draw = ImageDraw.Draw(image_source)
    bbox = draw.textbbox((0, 0), name, font=FONT_FILE)
    #print(bbox)
    name_width = bbox[2] - bbox[0] 
    #print(name_width)
    name_height = bbox[3] - bbox[1]

    draw.text((WIDTH-name_width/2, HEIGHT-name_height), name, fill=FONT_COLOR, font=FONT_FILE)#adjust the width / value according to the postion keep it in /2 to make it center
    path = os.path.join('out', name+'.png')
    image_source.save(path)
    print('printing certificate: '+name)

data=pd.read_excel('data.xlsx')
names = list(data.name)
for x in names:
    make_cert(x)
