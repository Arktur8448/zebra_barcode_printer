import os
import barcode
from barcode.writer import ImageWriter
from PIL import Image, ImageDraw, ImageFont, ImageWin
import win32print
import win32ui
from PIL import Image, ImageWin
import tkinter as tk
from tkinter import ttk

try:
    f = open("code.txt", "r")
    code = int(f.readline())
    f.close()
except FileNotFoundError:
    f = open("code.txt", "w")
    f.write("1234567890")
    f.close()
    code = 1234567890


def generate_bar():
    global code
    textcode = str(code)
    text = "•   " + textcode[:-1] + "   " + textcode[-1]

    ean = barcode.get_barcode_class('code128')

    writer = ImageWriter()
    writer.margin_top = 5


    barcode_obj = ean(textcode, writer=writer)
    barcode_obj.save("barcode_output", {'write_text': False, 'module_width': 0.4})


    img = Image.open("barcode_output" + '.png').convert("RGBA")

    new_height = img.height + 250
    new_img = Image.new('RGBA', (img.width, new_height), (255, 255, 255, 0))  # Fully transparent
    new_img.paste(img, (0, 0), img)

    draw = ImageDraw.Draw(new_img)
    font = ImageFont.truetype("arial.ttf", 50)
    text_bbox = draw.textbbox((0, 0), text, font=font)
    text_width = text_bbox[2] - text_bbox[0]
    text_x = (new_img.width - text_width) // 2
    text_y = img.height + 5
    draw.text((text_x, text_y), text, fill='black', font=font)

    new_img.save("print.png")

    os.remove("barcode_output.png")
    code += 1

def print_imgs():
    for _ in range(seriesVAR.get()):
        printer_name = win32print.GetDefaultPrinter()

        # Utwórz kontekst drukowania
        hDC = win32ui.CreateDC()
        hDC.CreatePrinterDC(printerVAR.get())
        hDC.StartDoc("Printing Images")

        for _ in range(int(amoutVAR.get())):
            generate_bar()
            image = Image.open("print.png")

            hDC.StartPage()

            dib = ImageWin.Dib(image)
            dib.draw(hDC.GetHandleOutput(), (110, 10, image.width - 170, image.height - 280))

            hDC.EndPage()


        hDC.EndDoc()
        hDC.DeleteDC()
        global code
        code += 10
        code -= code % 10
        f = open("code.txt", "w")
        f.write(str(code))
        f.close()
        global startingNumberVAR
        startingNumberVAR.set(str(code)[:-1])

root = tk.Tk()
root.title("Bar code printer")
root.resizable(False, False)

tk.Label(root, text="Starting Number:").grid(row=0, column=0, padx=20, pady=2)
tk.Label(root, text="Series:").grid(row=0, column=1, padx=20, pady=2)
tk.Label(root, text="Amount:").grid(row=0, column=2, padx=20, pady=2)

def on_number_change(*args):
    global code, startingNumberVAR
    code = int(f"{startingNumberVAR.get()}0")


startingNumberVAR = tk.IntVar()
startingNumberVAR.set(str(code)[:-1])
startingNumber = tk.Entry(root, textvariable=startingNumberVAR)
startingNumber.grid(row=1, column=0, padx=20, pady=5)
startingNumberVAR.trace_add("write", on_number_change)

seriesVAR = tk.IntVar()
seriesVAR.set(1)
series = tk.Entry(root, textvariable=seriesVAR)
series.grid(row=1, column=1, padx=20, pady=5)

amoutVAR = tk.StringVar()
amount = ttk.Combobox(root, textvariable=amoutVAR, values=[str(i) for i in range(1, 11)], state="readonly")
amount.grid(row=1, column=2, padx=20, pady=5)
amount.current(0)


printerVAR = tk.StringVar()
printerList = ttk.Combobox(root, textvariable=printerVAR, values=[printer[2] for printer in win32print.EnumPrinters(2)], state="readonly")
printerList.grid(row=2, column=0, columnspan=2, padx=5, pady=5)
count = 0
for d in [printer[2] for printer in win32print.EnumPrinters(2)]:
    if d == win32print.GetDefaultPrinter():
        printerList.current(count)
    count += 1


sendprint = tk.Button(root, text="Print", command=print_imgs)
sendprint.grid(row=2, column=2, padx=5, pady=5, sticky="ew", )

root.mainloop()
