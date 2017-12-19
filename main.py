from tkinter import *
from tkinter import messagebox

import os


if __name__ == "__main__":

    app = Tk()
    app.title("pdf2cla")

    loadxrdp = False
    try:
        import xlrd
        loadxrdp = True
    except ModuleNotFoundError:
        if messagebox.askokcancel("import Error", "Unbable to import xlrp\n Try and install it?:\n pip install xlrp", icon='error'):
            os.system("pip install xlrd")
            try:
                import xlrd
                loadxrdp = True
            except ModuleNotFoundError:
                messagebox.showerror("Warning", " Still unable to load xlrd \n Without xlrd you will be unable to load .xlsx files \n Try and run \n\"pip install xlrd\"\n with admin rights?")
        else:
            messagebox.showerror("Warning", "Without xlrd you will be\nunable to load .xlsx files")

    loadgAPI = False
    try:
        import apiclient
        loadgAPI = True
    except ModuleNotFoundError:
        if messagebox.askokcancel("import Error", "Unbable to import goole-api\n Try and install it?:\n pip install google-api-python-client", icon='error'):
            os.system("pip install google-api-python-client")
            try:
                import apiclient
                loadgAPI = True
            except ModuleNotFoundError:
                messagebox.showerror("Warning", " Still unable to load google-api \n Without it you cant add Calendar Events \n Try and run \n\"pip install google-api-python-client\"\n with admin rights")
        else:
            messagebox.showerror("Warning", "Without this lib you can't add google Calendar Events")

    print(loadxrdp)
    print(loadgAPI)

    import xlsx2name
    if loadgAPI:
        import gCalendar