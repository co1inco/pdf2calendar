from tkinter import *
from tkinter import messagebox

import os

global input

class Checkbar(Frame): #https://www.python-kurs.eu/tkinter_checkboxes.php
    def __init__(self, parent=None, picks=[], anchor=W):
        Frame.__init__(self, parent)
        self.vars = []
        count = 0
        for pick in picks:
            var = IntVar()
            chk = Checkbutton(self, text=pick, variable=var, bg = bgMain)
            chk.grid(row=count, sticky=W)
 
#           chk.pack( anchor=anchor, expand=YES, fill=X)
            self.vars.append(var)
            count = count + 1
    def state(self):
        return map((lambda var: var.get()), self.vars)


def xlsxProcess(app):

    text = Label(text="Filename:").pack()
    inputText = Entry(app, width = 40)
    inputText.pack()
    button = Button(app, text="OK", command=lambda: useInput(app, inputText), width=10).pack()

def useInput(app, inputText):
    input = inputText.get()
    app.destroy()

    if os.path.isfile(input):

        if input.endswith(".xlsx") or endswith(".xls"):
             print("xlsx")
             t = xlsx2name.readXlsx(input)
             classes = t.getAllPages()
             xlsx2name.writeToFile("classes.txt", classes)
        else:
            print("txt")
            classes = getFileContent(input)

        print(classes)

        if loadgAPI:
            createGoogleEntrys()

    else:
        messagebox.showerror("Warning", "unaple to locate file")

def createGoogleEntrys():
    pass



def main():
    
    app = Tk()
    app.title("pdf2cla")
    app.geometry("250x80")

    x = xlsxProcess(app)

    app.mainloop()






if __name__ == "__main__":

    global loadxrdp
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

    global loadgAPI
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
    
    main()