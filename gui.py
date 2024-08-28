import main
import os
from pathlib import *
from tkinter import Tk, Canvas, Entry, Button, PhotoImage, filedialog, messagebox
import tkinter as tk
import xlsx_searcher

OUTPUT_PATH = Path(__file__).parent
ASSETS_PATH = OUTPUT_PATH / Path(r"assets\frame0")

def chooseInputDir():
    main.directory1_path = filedialog.askdirectory()
    entry_1.configure(state='normal')
    entry_1.delete(0, tk.END)
    entry_1.insert(tk.END, main.directory1_path)
    entry_1.configure(state='disabled')

def chooseOutputDir():
    xlsx_searcher.directory2_path = filedialog.askdirectory()
    entry_2.configure(state='normal')
    entry_2.delete(0, tk.END)
    entry_2.insert(tk.END, xlsx_searcher.directory2_path)
    entry_2.configure(state='disabled')

def verifyDir():
    if main.directory1_path is not None:
        if os.path.exists(main.directory1_path):
            if not os.listdir(main.directory1_path):
                return "empty"
            else:
                return "ok"
        else:
            return "non-existent"
    else:
        return "non-existent"
    
def startReading():
    dir_status = verifyDir()
    if dir_status == "empty":
        messagebox.showerror("Error", "Please select a directory that is not empty.")
    elif dir_status == "ok":
        main.main()
        messagebox.showinfo("Done", "All PDFs have been converted, sorted and extracted as CSVs")
    elif dir_status == "non-existent":
        messagebox.showerror("Error","This directory does not exist or is not valid!")
    
def relative_to_assets(path: str) -> Path:
    return ASSETS_PATH / Path(path)

window = Tk()

window.geometry("700x550")
window.configure(bg = "#FFFFFF")
window.title("PDF Sorter")

canvas = Canvas(
    window,
    bg = "#FFFFFF",
    height = 550,
    width = 700,
    bd = 0,
    highlightthickness = 0,
    relief = "ridge"
)

canvas.place(x = 0, y = 0)
image_image_1 = PhotoImage(
    file=relative_to_assets("image_1.png"))
image_1 = canvas.create_image(
    350.0,
    2.0,
    image=image_image_1
)

canvas.create_rectangle(
    0.0,
    4.0,
    700.0,
    61.0,
    fill="#FFFFFF",
    outline="")

image_image_2 = PhotoImage(
    file=relative_to_assets("image_2.png"))
image_2 = canvas.create_image(
    69.0,
    32.0,
    image=image_image_2
)

entry_image_1 = PhotoImage(
    file=relative_to_assets("entry_1.png"))
entry_bg_1 = canvas.create_image(
    316.5,
    238.5,
    image=entry_image_1
)
entry_1 = Entry(
    bd=0,
    bg="#EFF1F2",
    fg="#000716",
    highlightthickness=0,
    cursor="arrow",
    state="disabled"
)

entry_1.place(
    x=78.0,
    y=218.0,
    width=477.0,
    height=39.0
)

entry_image_2 = PhotoImage(
    file=relative_to_assets("entry_2.png")
)

entry_bg_2 = canvas.create_image(
    316.5,
    341.5,
    image=entry_image_2
)
entry_2 = Entry(
    bd=0,
    bg="#EFF1F2",
    fg="#000716",
    highlightthickness=0,
    cursor="arrow",
    state="disabled"
)
entry_2.place(
    x=78.0,
    y=321.0,
    width=477.0,
    height=39.0
)

canvas.create_text(
    78.0,
    186.0,
    anchor="nw",
    text="Selecione o diretório de origem:",
    fill="#000000",
    font=("BoschSansGlobal Regular", 18 * -1)
)

canvas.create_text(
    78.0,
    290.0,
    anchor="nw",
    text="Selecione o diretório de destino:",
    fill="#000000",
    font=("BoschSansGlobal Regular", 18 * -1)
)

button_image_1 = PhotoImage(
    file=relative_to_assets("button_1.png"))
button_1 = Button(
    image=button_image_1,
    borderwidth=0,
    highlightthickness=0,
    command= chooseInputDir,
    relief="flat"
)
button_1.place(
    x=555.0,
    y=218.0,
    width=62.0,
    height=42.0
)

button_image_2 = PhotoImage(
    file=relative_to_assets("button_2.png"))
button_2 = Button(
    image=button_image_2,
    borderwidth=0,
    highlightthickness=0,
    command=chooseOutputDir,
    relief="flat"
)
button_2.place(
    x=555.0,
    y=321.0,
    width=62.0,
    height=42.0
)

button_image_3 = PhotoImage(
    file=relative_to_assets("button_3.png"))
button_3 = Button(
    image=button_image_3,
    borderwidth=0,
    highlightthickness=0,
    command=startReading,
    relief="flat"
)
button_3.place(
    x=266.0,
    y=424.0,
    width=170.0,
    height=32.0
)
window.resizable(False, False)
window.mainloop()