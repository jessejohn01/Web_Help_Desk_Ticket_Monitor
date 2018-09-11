import urllib.request
import requests
import time
import sys
import re
import lxml
import os
import ctypes
import winsound
import threading
import queue

#Test Comment v2

try:
    from Tkinter import *
except ImportError:
    from tkinter import *

try:
    import ttk
    py3 = False
except ImportError:
    import tkinter.ttk as ttk
    py3 = True

import Gui_support

from bs4 import BeautifulSoup

name = os.path.basename(__file__)

websiteLogin = "https://helpdesk.msu.montana.edu/helpdesk/WebObjects/Helpdesk.woa"
website = "https://helpdesk.msu.montana.edu/helpdesk/WebObjects/Helpdesk.woa/wa/TicketActions/view?tab=group"
startupTime = time.time()
refreshToggle = False
progressVariable = 0



class Help_Desk_Login: # This class is the structure of the login GUI.
    def __init__(self, top=None):
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        _bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#d9d9d9' # X11 color: 'gray85'
        _ana1color = '#d9d9d9' # X11 color: 'gray85'
        _ana2color = '#d9d9d9' # X11 color: 'gray85'
        font9 = "-family {Segoe UI} -size 10 -weight normal -slant "  \
            "roman -underline 0 -overstrike 0"

        top.geometry("375x381+649+219")
        top.title("Help Desk Login")
        top.configure(background="#d9d9d9")



        self.Frame1 = Frame(top)
        self.Frame1.place(relx=0.03, rely=0.03, relheight=0.75, relwidth=0.95)
        self.Frame1.configure(relief=GROOVE)
        self.Frame1.configure(borderwidth="2")
        self.Frame1.configure(relief=GROOVE)
        self.Frame1.configure(background="#d9d9d9")
        self.Frame1.configure(width=355)

        self.Entry1 = Entry(self.Frame1)
        self.Entry1.place(relx=0.51, rely=0.35,height=20, relwidth=0.35)
        self.Entry1.configure(background="white")
        self.Entry1.configure(disabledforeground="#a3a3a3")
        self.Entry1.configure(font="TkFixedFont")
        self.Entry1.configure(foreground="#000000")
        self.Entry1.configure(insertbackground="black")
        self.Entry1.configure(width=124)

        self.Entry2 = Entry(self.Frame1)
        self.Entry2.place(relx=0.51, rely=0.49,height=20, relwidth=0.35)
        self.Entry2.configure(background="white")
        self.Entry2.configure(disabledforeground="#a3a3a3")
        self.Entry2.configure(font="TkFixedFont")
        self.Entry2.configure(foreground="#000000")
        self.Entry2.configure(insertbackground="black")
        self.Entry2.configure(show="*")
        self.Entry2.configure(width=124)

        self.Label1 = Label(self.Frame1)
        self.Label1.place(relx=0.17, rely=0.32, height=31, width=74)
        self.Label1.configure(background="#d9d9d9")
        self.Label1.configure(disabledforeground="#a3a3a3")
        self.Label1.configure(font=font9)
        self.Label1.configure(foreground="#000000")
        self.Label1.configure(text='''Username :''')
        self.Label1.configure(width=74)

        self.Label2 = Label(self.Frame1)
        self.Label2.place(relx=0.17, rely=0.46, height=31, width=74)
        self.Label2.configure(background="#d9d9d9")
        self.Label2.configure(disabledforeground="#a3a3a3")
        self.Label2.configure(font=font9)
        self.Label2.configure(foreground="#000000")
        self.Label2.configure(text='''Password :''')
        self.Label2.configure(width=74)

        self.but43 = Button(top)
        self.but43.place(relx=0.35, rely=0.84, height=34, width=107)
        self.but43.configure(activebackground="#d9d9d9")
        self.but43.configure(activeforeground="#000000")
        self.but43.configure(background="#d9d9d9")
        self.but43.configure(command= lambda: loginPassThrough(top,self.Entry1,self.Entry2) )
        self.but43.configure(disabledforeground="#a3a3a3")
        self.but43.configure(foreground="#000000")
        self.but43.configure(highlightbackground="#d9d9d9")
        self.but43.configure(highlightcolor="black")
        self.but43.configure(pady="0")
        self.but43.configure(text='''Login''')
        self.but43.configure(width=107)

        self.menubar = Menu(top,font="TkMenuFont",bg=_bgcolor,fg=_fgcolor)
        top.configure(menu = self.menubar)


class Ticket_Monitor(threading.Thread): # This class is the structure of the main GUI window that runs. This runs on the main thread and will end the program if everything is closed.

    def __init__(self, top=None):
        global refreshToggle
        global progressVariable
        '''This class configures and populates the toplevel window.
           top is the toplevel containing window.'''
        _bgcolor = '#d9d9d9'  # X11 color: 'gray85'
        _fgcolor = '#000000'  # X11 color: 'black'
        _compcolor = '#d9d9d9' # X11 color: 'gray85'
        _ana1color = '#d9d9d9' # X11 color: 'gray85'
        _ana2color = '#d9d9d9' # X11 color: 'gray85'
        self.style = ttk.Style()
        if sys.platform == "win32":
            self.style.theme_use('winnative')
        self.style.configure('.',background=_bgcolor)
        self.style.configure('.',foreground=_fgcolor)
        self.style.configure('.',font="TkDefaultFont")
        self.style.map('.',background=
            [('selected', _compcolor), ('active',_ana2color)])

        top.geometry("600x450+537+258")
        top.title("Ticket Monitor")
        top.configure(background="#d9d9d9")



        self.Frame1 = Frame(top)
        self.Frame1.place(relx=0.02, rely=0.02, relheight=0.94, relwidth=0.96)
        self.Frame1.configure(relief=GROOVE)
        self.Frame1.configure(borderwidth="2")
        self.Frame1.configure(relief=GROOVE)
        self.Frame1.configure(background="#d9d9d9")
        self.Frame1.configure(width=575)

        self.style.configure('TNotebook.Tab', background=_bgcolor)
        self.style.configure('TNotebook.Tab', foreground=_fgcolor)
        self.style.map('TNotebook.Tab', background=
            [('selected', _compcolor), ('active',_ana2color)])
        self.TNotebook1 = ttk.Notebook(self.Frame1)
        self.TNotebook1.place(relx=0.02, rely=0.02, relheight=0.96
                , relwidth=0.96)
        self.TNotebook1.configure(width=554)
        self.TNotebook1.configure(takefocus="")
        self.TNotebook1_t0 = Frame(self.TNotebook1)
        self.TNotebook1.add(self.TNotebook1_t0, padding=3)
        self.TNotebook1.tab(0, text="Monitor",compound="left",underline="-1",)
        self.TNotebook1_t0.configure(background="#d9d9d9")
        self.TNotebook1_t0.configure(highlightbackground="#d9d9d9")
        self.TNotebook1_t0.configure(highlightcolor="black")
        self.TNotebook1_t0.configure(width=300)
        self.TNotebook1_t1 = Frame(self.TNotebook1)
        self.TNotebook1.add(self.TNotebook1_t1, padding=3)
        self.TNotebook1.tab(1, text="Commands",compound="left",underline="-1",)
        self.TNotebook1_t1.configure(background="#d9d9d9")
        self.TNotebook1_t1.configure(highlightbackground="#d9d9d9")
        self.TNotebook1_t1.configure(highlightcolor="black")

        self.TProgressbar1 = ttk.Progressbar(self.TNotebook1_t0, variable = progressVariable, mode = "determinate")
        self.TProgressbar1.place(relx=0.04, rely=0.08, relwidth=0.93
                , relheight=0.0, height=22)
        self.TProgressbar1.configure(length="510")






        self.Text1 = Text(self.TNotebook1_t0)
        self.Text1.place(relx=0.04, rely=0.16, relheight=0.8, relwidth=0.92)
        self.Text1.configure(background="white")
        self.Text1.configure(font="TkTextFont")
        self.Text1.configure(foreground="black")
        self.Text1.configure(highlightbackground="#d9d9d9")
        self.Text1.configure(highlightcolor="black")
        self.Text1.configure(insertbackground="black")
        self.Text1.configure(selectbackground="#c4c4c4")
        self.Text1.configure(selectforeground="black")
        self.Text1.configure(width=504)
        self.Text1.configure(wrap=WORD)





        self.Button1 = Button(self.TNotebook1_t1,command = lambda : setStartupTime())
        self.Button1.place(relx=0.04, rely=0.08, height=34, width=217)
        self.Button1.configure(activebackground="#d9d9d9")
        self.Button1.configure(activeforeground="#000000")
        self.Button1.configure(background="#d9d9d9")
        self.Button1.configure(disabledforeground="#a3a3a3")
        self.Button1.configure(foreground="#000000")
        self.Button1.configure(highlightbackground="#d9d9d9")
        self.Button1.configure(highlightcolor="black")
        self.Button1.configure(pady="0")
        self.Button1.configure(text='''Mark All Client Responses As Finished''')
        self.Button1.configure(width=217)

        self.Button2 = Button(self.TNotebook1_t1,command = lambda : setRefreshToggle(True))
        self.Button2.place(relx=0.51, rely=0.08, height=34, width=217)
        self.Button2.configure(activebackground="#d9d9d9")
        self.Button2.configure(activeforeground="#000000")
        self.Button2.configure(background="#d9d9d9")
        self.Button2.configure(disabledforeground="#a3a3a3")
        self.Button2.configure(foreground="#000000")
        self.Button2.configure(highlightbackground="#d9d9d9")
        self.Button2.configure(highlightcolor="black")
        self.Button2.configure(pady="0")
        self.Button2.configure(text='''Manual Refresh''')
        self.Button2.configure(width=217)
        threading.Thread.__init__(self)
        self.start()

    def start(self):
        sys.stdout = StdoutRedirector(self.Text1)
        self.TProgressbar1["value"] = 0
        self.updateBar()
        self.TProgressbar1["maximum"] = 300

    def updateBar(self):
        self.TProgressbar1["value"] = progressVariable
        if (progressVariable < 300):
            root.after(1, self.updateBar)

class loginAndMonitorLogic(threading.Thread): # Class that deals with all the logic that runs in the ticket monitor. It is structured to run on its own thread.


    def __init__(self):
        threading.Thread.__init__(self)
        self.daemon = True
        self.start()
    def run(self):
        global userName
        global password
        global refreshToggle
        global progressVariable
        with requests.Session() as sesh:
            print("Logging in with ", userName, " as username...")
            sesh.post(websiteLogin, data={'userName': userName, 'password': password})

            response = sesh.get(websiteLogin)
            html = response.content
            loginCheck = response.text

        soup = BeautifulSoup(html, 'lxml')
        textSoup = soup.get_text()
        if 'You can find your NetID username' in textSoup:
            messageBox('Ticket Monitor', 'Login failed please try again.', 0)

            os.system(name)

            # Main()
        else:
            messageBox('Ticket Monitor', 'Login was successful.', 0)
            starttime = time.time()

            while True:

                print("Loading...")

                response = sesh.get(website)

                html = response.text
                soup = BeautifulSoup(html, 'html.parser')
                textSoup = soup.get_text()
                os.system('cls')
                print("Refreshed")
                getNewTicketResponse(soup)

                # getTicketResponse(soup)

                # results = re.findAll('Ticket has a new client note, needs response',html)
                if 'Ticket is new, needs response' in html:
                    print('New ticket found!')
                    messageBox('Ticket Monitor', 'A new ticket has been found!', 0)
                    for i in range(0, 300):
                        time.sleep(1)
                        progressVariable += 1
                        if refreshToggle == True:
                            setRefreshToggle(False)
                            progressVariable = 0
                            break
                    progressVariable = 0


                else:
                    print('No new tickets found.')
                    for i in range(0, 300):
                        time.sleep(1)
                        progressVariable = progressVariable + 1
                        if refreshToggle == True:
                            setRefreshToggle(False)
                            progressVariable = 0
                            break

                    progressVariable = 0
        def stop(self):
            self.running = False



class IORedirector(object): # Default classes to help redirect console output to GUI window
    def __init__(self,console):
        self.console = console
class StdoutRedirector(IORedirector):
    def write(self,str):
        self.console.insert("end", str)
        self.console.see("end")



def printStartupTime(): #Function to get the current time.
    time.ctime()
    print("Program Startup Time: " + time.strftime('%I:%M%p %m/%d'))

def setStartupTime(): #Function that sets the startup time to the current time. Used for manually marking done tickets.
    global startupTime
    startupTime = time.time()
    printStartupTime()


def vp_start_gui_login(): # Starts the login GUI window.
    '''Starting point when module is the main routine.'''
    global val, w, root
    root = Tk()
    top = Help_Desk_Login (root)
    Gui_support.init(root, top)
    root.mainloop()

w = None
def create_Help_Desk_Login(root, *args, **kwargs): # Creates login window if accessed from another program.
    '''Starting point when module is imported by another program.'''
    global w, w_win, rt
    rt = root
    w = Toplevel (root)
    top = Help_Desk_Login (w)
    Gui_support.init(w, top, *args, **kwargs)
    return (w, top)

def destroy_Help_Desk_Login(): # Destroys login page.
    global w
    w.destroy()
    w = None

import MainWindow_Support_support

def vp_start_gui_MainWindow(): # Starter function fo main GUI window.
    '''Starting point when module is the main routine.'''
    global val, w, root
    root = Tk()
    top = Ticket_Monitor (root)
    MainWindow_Support_support.init(root, top)
    root.after(1, loginAndMonitorLogic)
    root.mainloop()


w = None
def create_Ticket_Monitor(root, *args, **kwargs): # If main window needs to be created through another program this helps out.
    '''Starting point when module is imported by another program.'''
    global w, w_win, rt
    rt = root
    w = Toplevel (root)
    top = Ticket_Monitor (w)
    MainWindow_Support_support.init(w, top, *args, **kwargs)
    return (w, top)

def destroy_Ticket_Monitor(): # Destroys main window.
    global w
    w.destroy()
    w = None





def loginPassThrough(master,entry1,entry2):#Helper Function for Login. This passes the login information and closes the login window once you login.
    global userName
    global password
    userName = entry1.get()
    password = entry2.get()
    master.destroy()






def messageBox(title,text,style): #This function is an easy call for message boxes.
    winsound.PlaySound('C:\Windows\media\notify.wav',winsound.SND_FILENAME)
    return ctypes.windll.user32.MessageBoxW(0,text,title,style)


def printTicketsInTable(soup): #Function to print all tickets.
    ticketRows = soup.select('div > table > tbody')
    for tag in soup.find_all("tbody", id = lambda value: value and value.startswith("row_")):

        print (tag.get('id').replace('row_',''))

def getTicketResponse(soup): # This function gets all tickets that have responses.
    print("These tickets have a client response: ")
    for ticket in soup.find_all("img"): # For every ticket with the html tag img.
        if ticket.get('title') == 'Ticket has a new client note, needs response':#Check for response in img tag through title
            #Print statement that traverses the html tree and prints the correct data.
            print("Ticket: " + ticket.parent.parent.parent.get('id').replace('row_','') + " Tech: " + ticket.parent.parent.find_next("td",{"class": 'leftAlignTop'}).find_next_sibling("td",{"class": 'leftAlignTop'}).text.replace(' ','').replace('\n',''))

def getNewTicketResponse(soup): # This function grabs any responses that have happened since the start of this program
    ticketList = [] # Creates a list
    ticketTimeDate= []
    ticketTimeTime = []
    print("These tickets have a client response: ")
    for ticket in soup.find_all("img"): # for all images


        #ticketTime = convertTime(ticketTime) #converts the time to a floating point
        if ticket.get('title') == 'Ticket has a new client note, needs response':
            temp = ticket.parent.find_next_sibling("td",{'class':'rightAlignTop'}).find_next_sibling("td",{'class':'rightAlignTop'}).find_next_sibling("td",{'class':'rightAlignTop'}).find_next("div").get_text().replace("\n","").strip()
            temp2 = ticket.parent.find_next_sibling("td",{'class':'rightAlignTop'}).find_next_sibling("td",{'class':'rightAlignTop'}).find_next_sibling("td",{'class':'rightAlignTop'}).find_next("div").find_next_sibling("div").get_text().replace("\n","").strip()
            ticketTimeDate.append(temp)
            ticketTimeTime.append(temp2)
            currentTicketTime = convertTime(ticketTimeDate,ticketTimeTime)
            if currentTicketTime > startupTime:
                messageBox('Ticket Monitor', 'A new client response has been found!', 0)
                print("Ticket: " + ticket.parent.parent.parent.get('id').replace('row_',
                                                                                 '') + " Tech: " + ticket.parent.parent.find_next(
                    "td", {"class": 'leftAlignTop'}).find_next_sibling("td", {"class": 'leftAlignTop'}).text.replace(
                    ' ', '').replace('\n', ''))



def convertTime(ticketTimeDate,ticketTimeTime): # Converts the date and time into an easy to manage unix time.
    convertedTime = 0
    for x in range(0,len(ticketTimeDate)):
        dateTime = ticketTimeDate[x] + " " + ticketTimeTime[x]
        pattern = '%m/%d/%y %I:%M %p'
        convertedTime = time.mktime(time.strptime(dateTime,pattern))
    return convertedTime


def setRefreshToggle(booleanValue): #Helper function to manually refresh.
    global refreshToggle
    refreshToggle = booleanValue


def Main(): #Main function that runs everything and establishes connection with the ticket system.
    printStartupTime()
    vp_start_gui_login()

    ##Multi-Threading Begins here
    vp_start_gui_MainWindow()
    logic = loginAndMonitorLogic()






















    ##login()





Main()








