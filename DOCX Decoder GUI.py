from tkinter import *
from tkinter.filedialog import askopenfilename
from tkinter import messagebox
from tkinter.ttk import Progressbar
from tkinter import Menu
import xml.etree.ElementTree as ET
import zipfile
import os
import webbrowser

inFilePath = ""
logOutput = ""
outFilePath = ""
absolute_path = ""
extract_path = "\\extract\\word\\"
media_path = "\\extract\\word\\media\\"
start_tag = "<html>"
end_tag = " </html>"
break_tag = "<br>"
space = " "
image_tag = "<img src=$>"
image_src = "image$.png"
bold_tag = "<b>"
bold_end = "</b>"
p_tag = "<p>"
p_end = "</p>"
font_start = "<font style=\"font-size:$px;\">"
font_end = "</font>"
color_start = "<font color=$>"
color_end = "</font>"
table_start = "<table border=1>"
table_end = "</table>"
row_start = "<tr>"
row_end = "</tr>"
column_start = "<td>"
column_end = "</td>"
new_list = "<ul><li>"
list_start = "<li>"
list_end = "</li>"
list_stop = "</ul>"
pid = ""
list_started = False

new = 2


# use this function to unzip a docx file
def unzip_file(file):
    zip_ref = zipfile.ZipFile(file, 'r')
    zip_ref.extractall(absolute_path + "\\extract")
    zip_ref.close()


# extract text from xml
def extract_text(parent):
    global logOutput
    global pid
    global list_started
    url = '{http://schemas.openxmlformats.org/wordprocessingml/2006/main}'
    for child in parent:
        if child.tag == url+"t":
            log(child.text, False)
        if child.tag == url+"tbl":
            logOutput = logOutput + table_start
            extract_text(child)
            logOutput = logOutput + table_end
        elif child.tag == url+"tr":
            logOutput = logOutput + row_start
            extract_text(child)
            logOutput = logOutput + row_end
        elif child.tag == url+"tc":
            logOutput = logOutput + column_start
            extract_text(child)
            logOutput = logOutput + column_end
        elif child.tag == url+"p":
            if list_started and (pid is not child.attrib[url+"rsidP"]):
                list_started = False
                logOutput = logOutput + list_stop
            if child.find('./'+url+'pPr') is not None:
                if child.find('./'+url+'pPr').find('./'+url+'pStyle') is not None:
                    if child.find('./'+url+'pPr').find('./'+url+'pStyle').attrib[url+"val"] == "ListParagraph":
                        if pid is not child.attrib[url+"rsidP"]:
                            pid = child.attrib[url+"rsidP"]
                            list_started = True
                            logOutput = logOutput+new_list
                        else:
                            logOutput = logOutput + list_start

            logOutput = logOutput + p_tag
            extract_text(child)
            logOutput = logOutput + p_end
            if child.find('./'+url+'pPr') is not None:
                if child.find('./'+url+'pPr').find('./'+url+'pStyle') is not None:
                    if child.find('./'+url+'pPr').find('./'+url+'pStyle').attrib[url+"val"] == "ListParagraph":
                        logOutput = logOutput + list_end
        elif child.tag == url+"r":
            if child.find('./'+url+'rPr') is not None:
                if child.find('./'+url+'rPr').find('./'+url+'b') is not None:
                    logOutput = logOutput + bold_tag
                if child.find('./'+url+'rPr').find('./'+url+'color') is not None:
                    font_color = child.find('./'+url+'rPr').find('./'+url+'color').attrib[url+"val"]
                    logOutput = logOutput + color_start.replace("$", font_color)
                if child.find('./'+url+'rPr').find('./'+url+'sz') is not None:
                    val = (int(child.find('./'+url+'rPr').find('./'+url+'sz').attrib[url+"val"]))
                    if val > 0:
                        logOutput = logOutput + font_start.replace("$", str(val))
                extract_text(child)
                if child.find('./'+url+'rPr').find('./'+url+'sz') is not None:
                    val = (int(child.find('./'+url+'rPr').find('./'+url+'sz').attrib[url+"val"]))
                    if val > 0:
                        logOutput = logOutput + font_end
                if child.find('./'+url+'rPr').find('./'+url+'color') is not None:
                    logOutput = logOutput + color_end
                if child.find('./'+url+'rPr').find('./'+url+'b') is not None:
                    logOutput = logOutput + bold_end
            else:
                extract_text(child)
        elif child.tag == "{http://schemas.openxmlformats.org/drawingml/2006/wordprocessingDrawing}docPr":
            logOutput = logOutput + space + image_tag.replace("$", absolute_path + media_path + image_src.replace("$", child.attrib["id"]))
        elif len(child):
            extract_text(child)


def log(output, newline):
    global logOutput
    logOutput = logOutput + output
    if newline:
        logOutput = logOutput + break_tag
    return output


def log_to_file():
    global logOutput
    try:
        with open(outFilePath, "w") as writeFile:
            writeFile.write(start_tag)
            logOutput=logOutput.replace("<font style=\"font-size:28px;\"> </font>", "")
            writeFile.write(logOutput)
            writeFile.write(end_tag)
    except:
        print(log("Unexpected Error, sorry for the inconvenience!!!", True))
        exit(0)


def extract_xml():
    global logOutput
    tree = ET.parse(absolute_path + extract_path + "document.xml")
    root = tree.getroot()
    extract_text(root)
    log_to_file()
    logOutput = ""
    open_html()


def convert():
    global outFilePath
    global absolute_path
    global pid
    global list_started
    if entry.get() and inFilePath:
        absolute_path = os.path.dirname(inFilePath)
        outFilePath = absolute_path + "\\" + entry.get()
        progress_bar()
        unzip_file(inFilePath)
        extract_xml()
        pid = ""
        list_started = False
    else:
        input_file.config(text='No file Selected', fg='red')


def progress_bar():
    progress.pack(fill=BOTH, expand=0)
    for i in range(1500):
        progress.step()
        master.update()
    progress.pack_forget()


def open_file():
    global inFilePath
    inFilePath = askopenfilename(initialdir=".", filetypes=(("docx files", "*.docx"), ("all files", "*.*")),
                                 title="Choose a file.")
    input_file.config(text=inFilePath, fg=blk)


def open_html():
    msgbx = messagebox.askquestion('Success', 'File converted successfully! Do you want to open the HTML?', icon='info')
    if msgbx == 'yes':
        webbrowser.open('file://' + os.path.realpath(outFilePath))
    else:
        win_exit()


def reset():
    global inFilePath
    inFilePath = ""
    input_file.config(text='')
    entry.delete(0, END)
    entry.insert(0, "output.html")


def win_exit():
    master.destroy()


def footer(win):
    lbl_1 = Label(win, text=text_1, font=("Lato", 11), anchor=CENTER, fg=wht, bg=blue, height=4, padx=20)
    lbl_1.pack(fill=BOTH, expand=0)
    lbl_2 = Label(win, text=text_2, font=("Lato", 9), anchor=E, fg='#ffffff', bg=blue, height=1)
    lbl_2.pack(fill=BOTH, expand=0)


def new_window():
    newwin = Toplevel(master)
    newwin.title("Welcome")
    newwin.geometry('550x550')
    newwin.configure(bg=grey)
    Label(newwin, text="About Us", font=bold14, anchor=CENTER, fg=wht, bg=blue, height=2).pack(fill=BOTH, expand=0)
    Label(newwin, text="TEAM", font=bold14, anchor=CENTER, fg=blk, bg=grey, height=3).pack()
    Label(newwin, image=image_he).pack()
    Label(newwin, text="Kiran Balla", bg=grey, fg=blue, font=L12).pack(pady=5)
    Label(newwin, image=image_she).pack(pady=5)
    Label(newwin, text="Rachana Bandapalle", bg=grey, fg=blue, font=L12).pack(pady=5)
    foot = Frame(newwin)
    foot.pack(side="bottom", fill="x")
    footer(foot)
    newwin.mainloop()


def btn(txt, cmnd):
    b1 = Button(master, text=txt, fg='white', bg='#f06e2c', font=L12, relief=RAISED, command=cmnd)
    b1.pack(side="left", ipadx=15, ipady=3, padx=(115, 0), pady=(5, 0))


blk = '#333333'
wht = '#f7f7f7'
blue = '#26354A'
grey = '#bcbcbc'

bold14 = 'Lato 14 bold'
L12 = 'Lato 12'
text_1 = "The primary objective of this tool is to open and manipulate a DOCX file\n " \
         "without using Microsoft Word with the help of python libraries."
text_2 = "© 2019 MyProject.com | All rights reserved"


master = Tk()
master.title("Welcome")
master.geometry('550x550')
master.configure(bg=grey)

image_she = PhotoImage(file="images/she.gif")
image_up = PhotoImage(file="images/upload.gif")
image_he = PhotoImage(file="images/he.gif")

menubar = Menu(master)

filemenu = Menu(menubar, tearoff=0)
filemenu.add_command(label="Choose File", command=open_file)
filemenu.add_command(label="Convert", command=convert)
filemenu.add_command(label="Reset", command=reset)
filemenu.add_separator()
filemenu.add_command(label="Exit", command=win_exit)
menubar.add_cascade(label="File", menu=filemenu)

filemenu2 = Menu(menubar, tearoff=0)
filemenu2.add_command(label="About Us", command=new_window)
menubar.add_cascade(label="Help", menu=filemenu2)
master.config(menu=menubar)

Label(master, text="Convert DOCX files", font=bold14, anchor=CENTER, fg=wht, bg=blue, height=2).pack(fill=BOTH, expand=0)
Label(master, text="Let's Get Started", font=bold14, anchor=CENTER, fg=blk, bg=grey, height=3).pack()
Label(master, text="File Name", font=L12, anchor=W, justify=LEFT, fg=blk, bg=grey).pack()

entry = Entry(master)
entry.pack(ipady=3, ipadx=3)
entry.delete(0, END)
entry.insert(0, "output.html")
entry.config(background='white', font=L12, width=50)

lb4 = Label(master, text="Select files to convert", font=L12, anchor=W, justify=LEFT, fg=blk, bg=grey).pack(pady=(30, 0))


b = Button(master, text="Choose File  ", image=image_up, compound="right", fg='white', bg='#f06e2c',
           font=L12, relief=RAISED, command=open_file)
b.pack(side="top", ipadx=3, ipady=3)

input_file = Label(master, font=("Lato", 10), anchor=W, justify=LEFT, background=grey)
input_file.pack(pady=(0, 5))

foot_frame = Frame(master)
foot_frame.pack(side="bottom", fill="x")

progress_frame = Frame(master)
progress_frame.configure(background=grey)
progress_frame.pack()

btn('Convert', convert)
btn('Reset', reset)

progress = Progressbar(progress_frame, orient=HORIZONTAL, length=200, mode="determinate", takefocus=True, maximum=1500)
progress.pack_forget()
footer(foot_frame)
mainloop()
