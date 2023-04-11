#!usr/bin/env

# import libraries
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont

# Import the xlrd module
import xlrd

# Open the Workbook
workbook = xlrd.open_workbook("name_list.xlsx")

# Open the worksheet
worksheet = workbook.sheet_by_index(0)

img = Image.open('final_invitations/final_homecoming.jpg')

# draw image object
I1 = ImageDraw.Draw(img)

width = img.width
height = img.height
  
# display width and height
print("The height of the image is: ", height)
print("The width of the image is: ", width)


# Iterate the rows and columns
for i in range(0, 16):
    img = Image.open('final_invitations/final_homecoming.jpg')

    # draw image object
    I1 = ImageDraw.Draw(img)
    # Print the cell values with tab space
    only_name = worksheet.cell_value(i, 1)
    name = worksheet.cell_value(i, 0) +" "+ worksheet.cell_value(i, 1)
    print(name)
    
    # open image
    

    myFont = ImageFont.truetype('Fonts/Demo_ConeriaScript.ttf', 35)
    _, _, w, h = I1.textbbox((0, 0), name, font=myFont)
    #I1.text(((width-w)/2, (height-h)/2), name, font=ont, fill=fontColor)
    # Add Text to an image
    #I1.text((700, 420), name , font=myFont, fill =(0, 0, 0))
    I1.text(((width-w)/2, 180), name , font=myFont, fill =(255, 255, 255))

    # save image
    img.save("images2/"+only_name+".png")