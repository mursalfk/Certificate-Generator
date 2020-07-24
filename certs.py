#Imports Required Packages from PIL
from PIL import Image, ImageDraw, ImageFont

#Import Pandas for better access of Data and .xlsx File
import pandas as pd

#Import the file that contains all the details
data = pd.read_excel("file.xlsx")

#Import 'Name' List from the imported .xlsx file
name_list = data['Name'].to_list()

#Determining the length of the list
max_no = len(name_list)

#The Loops for creating certificates in bulk
# for i in name_list:
for idx, i in enumerate(name_list):

    im = Image.open("cert.jpg")
    d = ImageDraw.Draw(im)
    location = (275, 1050)
    text_color = (0, 137, 209)
    font = ImageFont.truetype("AlexBrush-Regular.ttf", 250, encoding="unic")
    d.text(location, i.title(), fill=text_color,font=font)
    im.save("certificate_"+i+".pdf")
    print("(%d/%d) Certificate Created for:  %s" % (idx+1, max_no, i.title()))
    
print("""\n*************************
All Certificates Created! \nPlease follow me on Github for more awesome codes\n\tGitHub: https://github.com/mursalfk
*************************
""")
#Read readme.md for further instructions