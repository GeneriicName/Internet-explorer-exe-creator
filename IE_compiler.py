import sys
from validators import url as check_url
import tkinter
import tkinter as tk
from tkinter import filedialog as fd
from tkinter.messagebox import showerror
from tkinter import simpledialog
from os import environ, path, unlink
from subprocess import run, DEVNULL
from random import random
from icoextract import IconExtractor

compiler = r"C:\Windows\Microsoft.NET\Framework\v4.0.30319\vbc.exe"
if not path.exists(compiler):
    showerror("No vbc compiler found!", message=f"No VBC compiler found at {compiler}\nPlease "
                                                f"install microsoft .NET SDK"
                                                f"https://dotnet.microsoft.com/en-us/download/dotnet-framework")
    sys.exit(1)


def reset():
    ready.icon = False
    ready.output = False
    ready.url = False
    ready.url_actual = None
    ready.icon_path = None
    ready.output_path = None
    convert_but.configure(state="disabled", cursor="arrow")
    url_entry.configure(text="URL: None")
    icon_path_lab.configure(text="Icon path: None")
    output_path_lab.configure(text="Output: None")


class IsReady:
    def __init__(self):
        self.icon = False
        self.output = False
        self.url = False
        self.url_actual = None
        self.icon_path = None
        self.output_path = None
        self.desktop = environ["USERPROFILE"] + "\\desktop"

    def check(self):
        return all([self.output, self.icon, self.url])


ready = IsReady()


def select_icon():
    try:
        icon_path = fd.askopenfile(filetypes=[('icon files', '*.ico')], initialdir=ready.desktop).name
        ready.icon_path = icon_path
        ready.icon = True
        icon_path_lab.configure(text=f"Icon path: {icon_path}")
    except AttributeError:
        pass
    if ready.check():
        convert_but.configure(state="normal", cursor="hand2")
    else:
        convert_but.configure(state="disabled", cursor="arrow")


def select_output():
    try:
        output_path = fd.asksaveasfile(mode='w', defaultextension=".exe", initialdir=ready.desktop).name.strip()
        if not output_path.lower().endswith(".exe"):
            output_path += ".exe"
        ready.output_path = output_path
        ready.output = True
        output_path_lab.configure(text=f"Output path: {output_path}")
        unlink(output_path)
    except AttributeError:
        pass
    if ready.check():
        convert_but.configure(state="normal", cursor="hand2")
    else:
        convert_but.configure(state="disabled", cursor="arrow")


def select_url():
    url = simpledialog.askstring(title="URL", prompt="URL for the website:\t\t\t\t\t\t")
    if url is None:
        return
    elif not check_url(url):
        if not tkinter.messagebox.askyesno(title=f'Invalid URL',
                                           message=f'Invalid URL "{url}"\nDo you want to continue?'):
            return
    ready.url = True
    ready.url_actual = url
    url_entry.configure(text=f"URL: {url}")
    if ready.check():
        convert_but.configure(state="normal", cursor="hand2")
    else:
        convert_but.configure(state="disabled", cursor="arrow")



def extract():
    try:
        extract_path = fd.askopenfile(filetypes=[('executables', '*.exe')], initialdir=ready.desktop,
                                      title="Select a file to extract it's icon").name.strip()
        output_icon_path = fd.asksaveasfile(mode='w', defaultextension=".ico", initialdir=ready.desktop,
                                            title="Select location to save the icon").name.strip()
        output_icon_path.strip("ico")
    except AttributeError:
        return
    try:
        extractor = IconExtractor(extract_path)
        extractor.export_icon(output_icon_path, num=0)
        tk.messagebox.showinfo(title="Success!", message=f"Extracted the icon to {output_icon_path}")
    except Exception as e:
        unlink(output_icon_path)
        showerror("Error", f"Didnt extract.\n{e}")


def compile_():
    script = f"""
    Module Module1
        Dim url As String = "{ready.url_actual}"
        Dim wWidth As Integer
        Dim wHeight As Integer
        Dim objApp As Object
        Dim ie As Object

        Sub Main()
            objApp = CreateObject("Shell.Application")
            ie = Nothing

            For Each Window In objApp.Windows
                If InStr(Window.Name, "Internet Explorer") Then
                    ie = Window
                End If
            Next

            If ie Is Nothing Then
                NewIE()
            Else
                OpenIE()
            End If

            ie = Nothing
        End Sub

        Sub NewIE()
            ie = CreateObject("InternetExplorer.Application")
            MaximizeWindow(ie.hwnd, 3)
            CreateObject("WScript.Shell").AppActivate("Internet Explorer")
            ie.Navigate(url)
        End Sub

        Sub OpenIE()
            CreateObject("WScript.Shell").AppActivate("Internet Explorer")
            ie.Visible = True
            MaximizeWindow(ie.hwnd, 3)
            ie.Navigate2(url, 2048)
        End Sub

        Declare Function MaximizeWindow Lib "user32" Alias "ShowWindow" (ByVal hwnd As Long, ByVal nCmdShow As Long) As Long

    End Module
    """

    will_comp = fr"{environ['temp']}\{random()}{random()}".replace(".", "") + ".vb"
    with open(will_comp, "w") as to_comp:
        to_comp.write(script)
    run([compiler, "/target:winexe", f'/out:{ready.output_path}',
         f'/win32icon:{ready.icon_path}', will_comp], stdout=DEVNULL, stderr=DEVNULL)

    unlink(will_comp)
    tkinter.messagebox.showinfo(title="Success", message=f'Successfully compiled to "{ready.output_path}')
    reset()


font = ("Ariel", 12, "bold")
root = tk.Tk()
root.geometry("761x340")
root.configure(background="#DDD1D1")
root.resizable(False, False)

extract_but = tk.Button(root, text="Extract", background="#0F9BC8", font=("Ariel", 20, "bold"), command=extract,
                        cursor="hand2")
extract_but.place(x=15, y=17, width=730, height=41)

url_but = tk.Button(root, text="Select URL", background="#ADAEBB", font=("Ariel", 10, "bold"), command=select_url,
                    cursor="hand2")
url_but.place(x=15, y=74, width=100, height=41)
url_entry = tk.Label(root, background="#D9D9D9", text="URL: None", font=font)
url_entry.place(x=150, y=74, width=550, height=41)

icon_but = tk.Button(root, text="Pick an icon", background="#ADAEBB", font=font, command=select_icon, cursor="hand2")
icon_but.place(x=15, y=131, width=100, height=41)
icon_path_lab = tk.Label(root, background="#D9D9D9", text="Icon path: None", font=font)
icon_path_lab.place(x=150, y=131, width=550, height=41)

output_but = tk.Button(root, text="Output path", background="#ADAEBB", font=font, command=select_output, cursor="hand2")
output_but.place(x=15, y=188, width=100, height=41)
output_path_lab = tk.Label(root, background="#D9D9D9", text="Output: None", font=font)
output_path_lab.place(x=150, y=188, width=550, height=41)

convert_but = tk.Button(root, text="Convert", background="#0F9BC8", font=("Ariel", 20, "bold"), command=compile_)
convert_but.configure(state="disabled")
convert_but.place(x=15, y=268, width=730, height=41)
root.mainloop()
