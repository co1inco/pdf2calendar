from tkinter import *
from tkinter import messagebox
from tkinter import ttk

import os

global input

# exitcode -5 => exit and interupting creation of entrys by closing the progress window

class LoadingScreen():
    def __init__(self, work, length):
        self.length = length
        self.lScreen = Toplevel()
#        self.lScreen = Tk()
        Label(self.lScreen, text="Working").pack()
        self.progress = 0

        self.work  = Label(self.lScreen, text=work).pack()
        
        self.progress_str = StringVar()
        self.progress_str.set(str(self.progress) + ' / ' + str(self.length))
        self.propgressLabel = Label(self.lScreen, text=self.progress_str, textvariable=self.progress_str).pack()

        self.progress_var = DoubleVar()
        self.prog_bar = ttk.Progressbar(self.lScreen, variable=self.progress_var, maximum=length)
        self.prog_bar.pack()

        self.lScreen.pack_slaves()

        self.lScreen.protocol("WM_DELETE_WINDOW", self.destroyed)

    def increaseProgress(self):
        self.progress = self.progress + 1
        if self.progress >= self.length:
            progress = 0
        self.progress_var.set(self.progress)
        self.progress_str.set(str(self.progress) + ' / ' + str(self.length))

    def update(self):
        self.lScreen.update()

    def destroyed(self):
        os._exit(-5)
        pass

    def __del__(self):
        self.lScreen.destroy()
        del self

#------------------------------------------------

def main():
    
    app = Tk()
    app.title("pdf2cla")
    app.geometry("250x80")
    app.withdraw()
    app.resizable(False, False)
#    messagebox.showinfo("pdf to xlsx", "Use THIS online converter to \nconvert the .pdf to .xlsx \n https://www.ilovepdf.com/pdf_to_excel")

    infobox = Tk()
    infobox.title("pdf to xlsx")
    infobox.resizable(False, False)

    Label(infobox, text='Use THIS online converter to \nconvert the .pdf to .xlsx').pack()
    T = Text(infobox, height=1, width=37, relief='flat')
    T.insert(END, "https://www.ilovepdf.com/pdf_to_excel")
    T.config(state=DISABLED)
    T.pack()

    Button(infobox, command=lambda: xlsxProcess(infobox, app), text="OK", width=10).pack()

    app.mainloop()
    infobox.mainloop()


def xlsxProcess(infobox, app):
    infobox.destroy()

    app.deiconify()
    text = Label(text="Filename: \n (emty for entrys.txt or timetable.pdf)").pack()
    inputText = Entry(app, width = 40)
    inputText.pack()
    button = Button(app, text="OK", command=lambda: useInput(app, inputText), width=10).pack()


def useInput(app, inputText):
    input = inputText.get()
    app.destroy()

    if len(input) == 0:
        if os.path.isfile("entrys.txt"):
            input = "entrys.txt" 
        elif os.path.isfile("timetable.pdf"):
            input = "timetable.xlsx"
        else:
            input = "error"

    if os.path.isfile(input):

        if input.endswith(".xlsx") or input.endswith(".xls"):
             print("xlsx")
             t = xlsx2name.readXlsx(input)
             classes = t.getAllPages()
             xlsx2name.writeToFile("entrys.txt", classes)
        else:
            print("txt")
            classes = xlsx2name.getFileContent(input)

        if loadgAPI:
            preGoogleEntrys(classes)

    else:
        messagebox.showerror("Warning", "unaple to locate file")
        os._exit(-3)


def preGoogleEntrys(classes):

    app = Tk()
    app.title("pdf2cla")
    app.resizable(False, False)

    calendar = gCalendar.gCalendar()

#    calendarId = 'bjo0233a5f7clkofr5khtt8608@group.calendar.google.com'

    calendarList = calendar.getCalendarList()

    selected = IntVar()
    for i, j in enumerate(calendarList):
        print(j['summary'])
        Radiobutton(app, text=j['summary'], variable=selected, value=i).grid(row=i, sticky='W')

    but = Button(app, text='OK', command=lambda: createGoogleEntrys(app,calendar, calendarList[selected.get()]['id'], classes), width=20).grid()

    app.mainloop()

def createGoogleEntrys(app,calendar, calendarId, classes):
    app.withdraw()
    
    loading = LoadingScreen("Creating Entys", len(classes))
    try:
#    if True:
        for i in classes:
            startTime = i[0] + "T" + i[1][0:2] + ":" + i[1][3:5]
            endTime   = i[0] + "T" + i[1][6:8] + ":" + i[1][9:11]

            if len(i) < 4:
                calendar.createEvent(calendarId, startTime, endTime, eventName=i[2])
                print(calendarId + " " + startTime + " " + endTime + i[2])
            elif len(i) > 3:

                name = ""
                for k in range(2, len(i)-1):
                    name = name + " " + i[k]

                calendar.createEvent(calendarId, startTime, endTime, eventName=name, location=i[-1])
                print(calendarId + " " + startTime + " " + endTime + name + "\t:\t" + i[-1])
            loading.increaseProgress()
            loading.update()

    except:
        messagebox.showerror("Warning", "Error while creating Calendar Entrys")
        loading.__del__()
        os._exit(-3)

    loading.__del__()
    messagebox.showinfo("Ready!", "Ready!\n  :-) ")
    os._exit(1)


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