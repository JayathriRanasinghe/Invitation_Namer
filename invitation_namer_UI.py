#!usr/bin/env

# import libraries
from PIL import Image
from PIL import ImageDraw
from PIL import ImageFont

# Import the xlrd module
import xlrd



#-----------------------------------------------------------------------------
import tkinter as tk
from tkinter import messagebox

class App:
    def __init__(self, master):
        self.master = master
        self.master.title("Invitation Namer")

        # Create the text fields
        tk.Label(master, text="Excel file path").grid(row=0)
        tk.Label(master, text="Invitation path").grid(row=1)
        tk.Label(master, text="Height").grid(row=2)
        tk.Label(master, text="Width").grid(row=3)
        tk.Label(master, text="Font").grid(row=4)
        tk.Label(master, text="Font size").grid(row=5)
        tk.Label(master, text="Font color").grid(row=6)
        tk.Label(master, text="keep text field empty to get the center alignment to height and width and \n other default settings").grid(row=7)
        self.field1 = tk.Entry(master)
        self.field2 = tk.Entry(master)
        self.field3 = tk.Entry(master)
        self.field4 = tk.Entry(master)
        self.field5 = tk.Entry(master)
        self.field6 = tk.Entry(master)
        self.field7 = tk.Entry(master)
        self.field1.grid(row=0, column=1)
        self.field2.grid(row=1, column=1)
        self.field3.grid(row=2, column=1)
        self.field4.grid(row=3, column=1)
        self.field5.grid(row=4, column=1)
        self.field6.grid(row=5, column=1)
        self.field7.grid(row=6, column=1)

        # Create the create button
        tk.Button(master, text="Create", command=self.create).grid(row=7, column=1)

    def create(self):
        # Get the values from the text fields
        value1 = self.field1.get()
        value2 = self.field2.get()
        value3 = self.field3.get()
        value4 = self.field4.get()
        value5 = self.field5.get()
        value6 = self.field6.get()
        value7 = self.field7.get()

        # Print the values to the console
        print("Excel path:", value1)
        print("Invitation path:", value2)
        print("Height:", value3)
        print("Width:", value4)
        print("Font:", value5)
        print("Font size:", value6)
        print("Font color:", value7)
        
        if value1 == '':
            value1 = "name_list.xlsx"
        if value2 == '':
            value2 = "final_invitations/final_invitation.jpg"
        # Open the Workbook
        workbook = xlrd.open_workbook(value1)

        # Open the worksheet
        worksheet = workbook.sheet_by_index(0)

        img = Image.open(value2)

        # draw image object
        I1 = ImageDraw.Draw(img)

        width = img.width
        height = img.height
        
        # display width and height
        print("The height of the image is: ", height)
        print("The width of the image is: ", width)
        

            
        if value5 == '':
            value5 = "Fonts/Demo_ConeriaScript.ttf"
        if value6 == '':
            value6 = 35
        if value7 == '':
            value7 = (255,255,255)
        else:
            value7 = eval(value7)
        # Iterate over rows and columns until the last entry
        for row in range(worksheet.nrows+1):
            img = Image.open(value2)

            # draw image object
            I1 = ImageDraw.Draw(img)
            if row == worksheet.nrows:
                # Stop iterating at the last entry
                break
            else:
                # Print the value of the current cell
                only_name = worksheet.cell_value(row, 1)
                name = worksheet.cell_value(row, 0) +" "+ worksheet.cell_value(row, 1)
                print(name)
        
                myFont = ImageFont.truetype(value5, int(value6))
                _, _, w, h = I1.textbbox((0, 0), name, font=myFont)
                
                
                if value3 == '':
                    value3 = (height-h)/2
                else:
                    value3 = int(value3)
                        
                if value4 == '':
                    value4 = (width-w)/2
                else:
                    value4 = int(value4)
                #I1.text(((width-w)/2, (height-h)/2), name, font=ont, fill=fontColor)
                # Add Text to an image
                #I1.text((700, 420), name , font=myFont, fill =(0, 0, 0))
                
                
                I1.text((value4, value3), name , font=myFont, fill =value7)

                # save image
                img.save("images/"+only_name+".png")

        # Show a success message in a pop-up window
        messagebox.showinfo("Success", "Invitations are created successfully!")

root = tk.Tk()
app = App(root)
root.mainloop()


