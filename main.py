from tkinter import *
from tkinter import messagebox
from tkinter import ttk
from tkinter import filedialog

import os

#global input

timezones = ["Europe/Berlin", "Europe/Dublin"]

# exitcode 5 => exit and interupting creation of entrys by closing the progress window

class LoadingScreen():
    def __init__(self, work, length):
        self.length = length
        self.lScreen = Toplevel()
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
        os._exit(5)
        pass

    def __del__(self):
        self.lScreen.destroy()
        del self

#------------------------------------------------

def main():
    
    app = Tk()
    app.title("pdf2cla")
    app.resizable(False, False)

    app.protocol("WM_DELETE_WINDOW", on_closing)

    Label(app, text='Use THIS online converter to \nconvert the .pdf to .xlsx').grid(row=0)
    T = Text(app, height=1, width=37, relief='flat')
    T.insert(END, "https://www.ilovepdf.com/pdf_to_excel")
    T.config(state=DISABLED)
    T.grid(row=1)

    Button(app, command=lambda: xlsxProcess(app), text="OK", width=10).grid(row=2)

    app.mainloop()
    infobox.mainloop()


def xlsxProcess(app):
    filename = filedialog.askopenfilename(initialdir = os.getcwd, title = "Select file", filetypes = (("Excel or Text file", "*.xls *.xlsx *.txt"), ("all files","*.*")))
    useInput(app, filename)



def useInput(app, inputText):


    print(inputText)
    
    if os.path.isfile(inputText):

        if inputText.endswith(".xlsx") or inputText.endswith(".xls"):
             print("xlsx")
             t = xlsx2name.readXlsx(inputText)
             classes = t.getAllPages()
             txtName = inputText+".txt"
             xlsx2name.writeToFile(txtName, classes)
             print("Entrys saved as: " + txtName)
        else:
            print("txt")
            classes = xlsx2name.getFileContent(inputText)

        if loadgAPI:
            preGoogleEntrys(classes, app)

    else:
        messagebox.showerror("Warning", "unaple to locate file")
        os._exit(-3)


def preGoogleEntrys(classes, app):

    app.title("pdf2cal")
    app.resizable(False, True)
    rowOffset = 4

    calendar = gCalendar.gCalendar()

#    calendarId = 'bjo0233a5f7clkofr5khtt8608@group.calendar.google.com'

    calendarList = calendar.getCalendarList()

    Label(app, text="------------------------").grid(row=rowOffset)
    
    timezoneVar = StringVar()
    timezoneVar.set(timezones[0])
    tzMenu = OptionMenu(app, timezoneVar, *timezones)
    tzMenu.config(width=len(timezoneVar.get()))
    tzMenu.grid(row=rowOffset+1)

    
    selected = IntVar()
    selected.set(1)
    for i, j in enumerate(calendarList):
        print(j['summary'])
        Radiobutton(app, text=j['summary'], variable=selected, value=i).grid(row=i+2+rowOffset, sticky='W')

    but = Button(app, text='OK', command=lambda: createGoogleEntrys(app,calendar, calendarList[selected.get()]['id'], classes, timezoneVar.get()), width=20).grid()

    app.mainloop()

def createGoogleEntrys(app,calendar, calendarId, classes, timezone):
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
                
                if i[-1] == "\n": # if a tab comes after the room \n is the room
                    i = i[:-1]

                name = ""
                for k in range(3, len(i)):
                    name = name + "  " + i[k-1]

#                print(i)
#                print(name)

                calendar.createEvent(calendarId, startTime, endTime, eventName=name, location=i[-1], timezone=timezone)
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


def on_closing():
    os._exit(1)


if __name__ == "__main__":

    global loadxrdp
    loadxrdp = False
    try:
        import xlrd
        loadxrdp = True
    except ModuleNotFoundError:
        if messagebox.askokcancel("import Error", "Unbable to import xlrp\n Try and install it?:\n pip install xlrp", icon='error'):
            os.system("py -m pip install xlrd")
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
            os.system("py -m pip install google-api-python-client")
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
