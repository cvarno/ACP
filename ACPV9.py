""" 
Welcome, I hope you visit this code only by curiosity, and not by necessity.
If, by some stroke of terrible destiny, you do find yourself here against your will,
then I pity you, for the code below is a terrible thing, and should be avoided at all costs.

In all seriousness, this code is horrible. I didn't know how classes worked or how to separate files.
Good luck.
"""


#region Overhead
import random                           #Allow for random generation, used to create a random password
import string                           #Used with random to generate random characters
import tkinter as tk                    #Import tkinter, shortened to tk, for all visual interfaces
from tkinter import *                   #IDK why but this breaks if removed, so its here twice
from tkinter.ttk import *               #ttk is tkinter toolkit, which adds some extra boxes
import subprocess                       #Allows me to run windows programs, used to run powershell windows
import time                             #Gets the current time, allows for automatic polling of user and group lists.
import os                               #Allows some special manipulation of the file system.
from tkinter.messagebox import *        #Allows popup messages to prompt the user
from tkinter import simpledialog        #Popup Messages to display basic info
from idlelib.tooltip import Hovertip    #Allows more information to be displayed when hovering over an element
import pathlib                          #Lets me check the time a file was last modified for file checking purposes
import webbrowser                       #Doesn't actually open a web browser, but is used to open a notepad file with the account info.
from sys import exit                    #Closes the program

#
#region Classes
class Create():
    #Move groups from the Inactive to the Active sides.
    def Inactive_to_Active(event, selected_group=''):
        if selected_group=='': #Let the program directly tell the function what to change. If nothing is provided, default to the original behavior.
            selected_group=Create.Inactive_Visible[event.widget.curselection()[0]]

        Create.Active_Master.append(selected_group) #Add the group to the other list
        Create.Active_Master.sort() #Sort the other list alphabetically
        Create.Active_Update(event=None,value=Create.Active_Search.get()) #Refresh the other column visually

        Create.Inactive_Master.remove(selected_group) #Remove from this list
        Create.Inactive_Update(event=None,value=Create.Inactive_Search.get()) #Refresh this column visually

    def Active_to_Inactive(event, selected_group=''):
        if selected_group=='': #Let the program directly tell the function what to change. If nothing is provided, default to the original behavior.
            selected_group=Create.Active_Visible[event.widget.curselection()[0]]
        Create.Active_Selection=event.widget.curselection()[0] #Using the event that is passed to the function, determine what list value is selected
        if not selected_group in Create.Default_Groups: #Don't allow the user to remove the default groups. 
            Create.Inactive_Master.append(selected_group) #Add the group to the other list
            Create.Inactive_Master.sort() #Sort the other list alphabetically 
            Create.Inactive_Update(event=None,value=Create.Inactive_Search.get()) #Refresh the other column visually

            Create.Active_Master.remove(selected_group) #Remove from this list
            Create.Active_Update(event=None,value=Create.Active_Search.get()) #Refresh this column visually

    def Inactive_Update(event='', value=''):
        try:Create_Description_Update()
        except:pass
        if Create.Advanced_Mode.get() == True: #Different behavior if advanced mode is enabled or not
            try:
                val = event.widget.get()
            except:
                val=value
            if val == '':
                Create.Inactive_Visible = Create.Inactive_Master
            else:
                Create.Inactive_Visible=[]
                for item in Create.Inactive_Master:
                    if val.lower() in item.lower():
                        Create.Inactive_Visible.append(item)
            Create.Inactive_Listbox.delete(0, 'end')
            Create.Inactive_Visible.sort()
            for item in Create.Inactive_Visible:
                Create.Inactive_Listbox.insert('end', item)

        else: #Different behavior if advanced mode is enabled or not
            Create.Presets.Inactive_Master=[]
            Create.License_User_Button.config(state='normal')
            Create.License_User_Var.set(value=1)
            Create.License_User_Button.config(state='disabled')
            for group in Create.Inactive_Master:
                if group in Create.Presets.Master and not group in Create.Presets.Inactive_Master:
                    Create.Presets.Inactive_Master.append(group)
            try:
                val = event.widget.get()
            except:
                val=value
            if val == '':
                Create.Inactive_Visible = Create.Presets.Inactive_Master

            else:
                Create.Inactive_Visible=[]
                for item in Create.Presets.Inactive_Master:
                    if val.lower() in item.lower():
                        Create.Inactive_Visible.append(item)
            Create.Inactive_Listbox.delete(0, 'end')
            Create.Inactive_Visible.sort()
            for item in Create.Inactive_Visible:
                Create.Inactive_Listbox.insert('end', item)
        if Create.Advanced_Mode.get()==True:
            Create.License_User_Button.config(state='normal')
            Create.Enable_MFA_Button.config(state='normal')
        if Create.Advanced_Mode.get()==False:
            Create.License_User_Button.config(state='disabled')
            Create.Enable_MFA_Button.config(state='disabled')


    def Active_Update(event='', value=''):
        try:
            Create_Description_Update()
        except:
            pass
        try:
            val = event.widget.get()
        except:
            val=value
        if val == '':
            Create.Active_Visible = Create.Active_Master
        else:
            Create.Active_Visible=[]
            for item in Create.Active_Master:
                if val.lower() in item.lower():
                    Create.Active_Visible.append(item)
        Create.Active_Listbox.delete(0, 'end')
        Create.Active_Visible.sort()
        for item in Create.Active_Visible:
            Create.Active_Listbox.insert('end', item)


    class Presets():
        pass

class Modify():
    #Move groups from the Inactive to the Active sides.
    def Inactive_to_Active(event='', selected_group=''):
        if selected_group=='':
            selected_group=Modify.Inactive_Visible[event.widget.curselection()[0]] #Pull the currently selected list item's number from the listbox, then get the group that is assigned that number currently.
        Modify.Active_Master.append(selected_group) #Add the group to the other list
        Modify.Active_Master.sort() #Sort the other list alphabetically
        Modify.Active_Update(event=None,value=Modify.Active_Search.get()) #Refresh the other column visually

        Modify.Inactive_Master.remove(selected_group) #Remove from this list
        Modify.Inactive_Update(event=None,value=Modify.Inactive_Search.get()) #Refresh this column visually

    def Active_to_Inactive(event, selected_group=''):
        if selected_group=='':
            selected_group=Modify.Active_Visible[event.widget.curselection()[0]] #Pull the currently selected list item's number from the listbox, then get the group that is assigned that number currently.
        if not selected_group in Modify.Default_Groups: #Don't allow the user to remove the default groups. 
            Modify.Inactive_Master.append(selected_group) #Add the group to the other list
            Modify.Inactive_Master.sort() #Sort the other list alphabetically 
            Modify.Inactive_Update(event=None,value=Modify.Inactive_Search.get()) #Refresh the other column visually

            Modify.Active_Master.remove(selected_group) #Remove from this list
            Modify.Active_Update(event=None,value=Modify.Active_Search.get()) #Refresh this column visually
        Modify_Description_Update()

    def Inactive_Update(event='', value=''):
        if Modify.Advanced_Mode.get() == True:
            try:
                val = event.widget.get()
            except:
                val=value
            if val == '':
                Modify.Inactive_Visible = Modify.Inactive_Master
            else:
                Modify.Inactive_Visible=[]
                for item in Modify.Inactive_Master:
                    if val.lower() in item.lower():
                        Modify.Inactive_Visible.append(item)
            Modify.Inactive_Listbox.delete(0, 'end')
            Modify.Inactive_Visible.sort()
            for item in Modify.Inactive_Visible:
                Modify.Inactive_Listbox.insert('end', item)
        else:
            Modify.Presets.Inactive_Master=[]
            for group in Modify.Inactive_Master:
                
                if group in Modify.Presets.Master and not group in Modify.Presets.Inactive_Master:
                    Modify.Presets.Inactive_Master.append(group)
            try:
                val = event.widget.get()
            except:
                val=value
            if val == '':
                Modify.Inactive_Visible = Modify.Presets.Inactive_Master

            else:
                Modify.Inactive_Visible=[]
                for item in Modify.Presets.Inactive_Master:
                    if val.lower() in item.lower():
                        Modify.Inactive_Visible.append(item)
            Modify.Inactive_Listbox.delete(0, 'end')
            Modify.Inactive_Visible.sort()
            for item in Modify.Inactive_Visible:
                Modify.Inactive_Listbox.insert('end', item)
        Modify_Description_Update()

    def Active_Update(event='', value=''):
        try:
            val = event.widget.get()
        except:
            val=value
        if val == '':
            Modify.Active_Visible = Modify.Active_Master
        else:
            Modify.Active_Visible=[]
            for item in Modify.Active_Master:
                if val.lower() in item.lower():
                    Modify.Active_Visible.append(item)
        Modify.Active_Listbox.delete(0, 'end')
        Modify.Active_Visible.sort()
        for item in Modify.Active_Visible:
            Modify.Active_Listbox.insert('end', item)
    class Presets():
        pass

#region Class Variables
DG_Groups=[] #Distribution Groups. A specific type of group pulled from Active Directory.
SG_Groups=[] #Security Groups. A specific type of group pulled from Active Directory.


Modify.Description=''
Modify.Inactive_Selection=int()
Modify.Active_Selection=int()

Modify.Master=[] #List of all groups assigned to the Modify page. 
Modify.Inactive_Master=[] #List of all of the groups currently designated as inactive.
Modify.Inactive_Visible=[] #List of all the inactive groups currently being displayed. Same if search is empty, different if something is searched.

Modify.Presets.Master=[] #List of all presets assigned to the Modify page. 
Modify.Presets.Inactive_Master=[] #List of all of the presets currently designated as inactive.
Modify.Presets.Inactive_Visible=[] #List of all the inactive presets currently being displayed. Same if search is empty, different if something is searched.

Modify.Active_Master=[] #List of all of the groups currently designated as active.
Modify.Active_Visible=[] #List of all the active groups currently being displayed. Same if search is empty, different if something is searched.


Create.Master=[] #List of all groups assigned to the Create page. 
Create.Inactive_Master=[] #List of all of the groups currently designated as inactive.
Create.Inactive_Visible=[] #List of all the inactive groups currently being displayed. Same if search is empty, different if something is searched.

Create.Presets.Master=[] #List of all presets assigned to the Create page. 
Create.Presets.Inactive_Master=[] #List of all of the presets currently designated as inactive.
Create.Presets.Inactive_Visible=[] #List of all the inactive presets currently being displayed. Same if search is empty, different if something is searched.

Create.Inactive_Selection=int()
Create.Active_Selection=int()

Create.Active_Master=[] #List of all of the groups currently designated as active.
Create.Active_Visible=[] #List of all the active groups currently being displayed. Same if search is empty, different if something is searched.

Modify.Presets.Corresponding=[  ['SG-Accounting','Accounting'], ['SG-Lifeguards'],  [''],               [''],   ['Cardiopulmonary-Rehab'],  ['Development','SG-Development','Donation-Reconciliation','Leadership Development Group','CatalogEditors'], ['GroupEx'],    ['Maintenance','SG-Maintenance'],   ['SG-FamilyPALZone-NW','PlayAndLearn-NW'],  ['SG-FamilyPALZone-SE','PlayAndLearn-SE'],  ['PersonalTrainers-NW','SG-Personal-Trainers'], ['PersonalTrainers-SE','SG-Personal-Trainers'], [''],       ['SG-WelcomeCenter','WelcomeCenter-NW'],    ['SG-WelcomeCenter','WelcomeCenter-SE'],    ['SG-Wellness-Coaches','WellnessCoaches-NW'],   ['SG-Wellness-Coaches','WellnessCoaches-SE'],   ['Yoga'],   ['IT-Test']]
Modify.Presets.Descriptions=[   "Accounting",                   "Lifeguard",        "Swim Instructor",  "Camp", "Cardiac Rehab",            "Development",                                                                                              "Group Ex",     "Maintenance",                      "Play and Learn NW",                        "Play and Learn SE",                        "Personal Trainer NW",                          "Personal Trainer SE",                          "Sports",   "Welcome Center NW",                        "Welcome Center SE",                        "Wellness Coach NW",                            "Wellness Coach SE",                            "Yoga",     "IT Test Account"]

Create.Presets.Corresponding=[['SG-Accounting','Accounting'],['SG-Lifeguards'],[''],[''],['Cardiopulmonary-Rehab'],['Development','SG-Development','Donation-Reconciliation','Leadership Development Group','CatalogEditors'],['GroupEx'],['Maintenance','SG-Maintenance'],['SG-FamilyPALZone-NW','PlayAndLearn-NW'],['SG-FamilyPALZone-SE','PlayAndLearn-SE'],['PersonalTrainers-NW','SG-Personal-Trainers'],['PersonalTrainers-SE','SG-Personal-Trainers'],'',['SG-WelcomeCenter','WelcomeCenter-NW'],['SG-WelcomeCenter','WelcomeCenter-SE'],['SG-Wellness-Coaches','WellnessCoaches-NW'],['SG-Wellness-Coaches','WellnessCoaches-SE'],['Yoga'],['IT-Test']]
Create.Presets.Descriptions=["Accounting","Lifeguard","Swim Instructor","Camp","Cardiac Rehab","Development","Group Ex","Maintenance","Play and Learn NW","Play and Learn SE","Personal Trainer NW","Personal Trainer SE","Sports","Welcome Center NW","Welcome Center SE","Wellness Coach NW","Wellness Coach SE","Yoga","IT Test Account"]



#endregion Class Variables
#endregion Classes

#region AutocompleteEntry

#Manually install the AutocompleteEntry module so I can change some parameters for it. 
#The AutocompleteEntry is a fork of the tkinter Entrybox which enables text to be automatically inserted more easily.
class AutocompleteEntry(tk.Entry):
    def set_completion_list(self, completion_list):
        self._completion_list = completion_list
        self._hits = []
        self._hit_index = 0
        self.position = 0
        self.bind('<KeyRelease>', self.handle_keyrelease)

    def autocomplete(self):
        # set position to end so selection starts where textentry ended
        self.position = len(self.get())

        # collect hits
        _hits = []
        for element in self._completion_list:
            if element.lower().startswith(self.get().lower()):
                _hits.append(element)

        # if we have a new hit list, keep this in mind 
        # if _hits != self._hits:
        self._hit_index = 0
        self._hits =_hits

        # # only allow cycling if we are in a known hit list
        # if _hits == self._hits and self._hits:
            # self._hit_index = (self._hit_index) % len(self._hits)

        # perform the auto completion
        if self._hits and autoCompToggleVar.get() == True:
            self.delete(0,tk.END)
            self.insert(0,self._hits[self._hit_index])
            self.select_range(self.position,tk.END)

    def handle_keyrelease(self, event):
        if len(event.keysym) == 1:
            self.autocomplete()
#endregion AutocompleteEntry

def Pull_Users():
    global User_List
    User_List=[]
    enabledList=[]
    disabledList=[]
    try: open(r'c:\ACP\userexportEnabled.txt', 'x')
    except FileExistsError: pass
    try: open(r'c:\ACP\userexportDisabled.txt', 'x')
    except FileExistsError: pass
    try: open(r'c:\ACP\groupExportSecurity.txt', 'x')
    except FileExistsError: pass
    try: open(r'c:\ACP\groupExportDistribution.txt', 'x')
    except FileExistsError: pass
    
    with open(r'c:\ACP\userexportEnabled.txt') as f:
        csvList = f.read().splitlines()
        for value in csvList:
            value = value[1:]
            value = value[:-1]
            enabledList.append(value)
            enabledList.sort()
    f.close()
    with open(r'c:\ACP\userexportDisabled.txt') as g:
        csvList = g.read().splitlines()
        for value in csvList:
            value = value[1:]
            value = value[:-1]
            disabledList.append(value)
            disabledList.sort()
    g.close()
    User_List=enabledList+disabledList

def Pull_Groups():
    #Pull the Security groups and assign them to the Groups.SG_Groups list.
    with open(r'c:\ACP\groupExportSecurity.txt') as h:
        csvList = h.read().splitlines()
        for value in csvList:
            value = value[1:]
            value = value[:-1]
            SG_Groups.append(value)
            SG_Groups.sort()


    #Pull the Distribution groups and assign them to the Groups.DG_Groups list.
    with open(r'c:\ACP\groupExportDistribution.txt') as i:
        csvList = i.read().splitlines()
        for value in csvList:
            value = value[1:]
            value = value[:-1]
            DG_Groups.append(value)
            DG_Groups.sort()
   

    #Assign the presets. These will also be apart of the inactive master when they are in there, but will behave as normal when active. On returing to inactive, make sure they go back into both lists.
    Create.Presets.Master=['Accounting (Preset)','Aquatics Lifeguards (Preset)','Aquatics Swim Instructior (Preset)','Camps (Preset)','Cardiac Rehab (Preset)','Development (Preset)','GroupEx (Preset)','Maintenance (Preset)','Play and Learn NW (Preset)','Play and Learn SE (Preset)','Personal Trainer NW (Preset)','Personal Trainer SE (Preset)','Sports (Preset)','Welcome Center NW (Preset)','Welcome Center SE (Preset)','Wellness Coach NW (Preset)','Wellness Coach SE (Preset)','Yoga (Preset)','IT Test Account (Preset)']
    Modify.Presets.Master=['Accounting (Preset)','Aquatics Lifeguards (Preset)','Aquatics Swim Instructior (Preset)','Camps (Preset)','Cardiac Rehab (Preset)','Development (Preset)','GroupEx (Preset)','Maintenance (Preset)','Play and Learn NW (Preset)','Play and Learn SE (Preset)','Personal Trainer NW (Preset)','Personal Trainer SE (Preset)','Sports (Preset)','Welcome Center NW (Preset)','Welcome Center SE (Preset)','Wellness Coach NW (Preset)','Wellness Coach SE (Preset)','Yoga (Preset)','IT Test Account (Preset)']

    #Assign list to Create.Master. Here I could filter out weird groups I don't want showing up.
    Create.Master=(DG_Groups+SG_Groups)
    Create.Inactive_Master=Create.Master
    for item in Create.Presets.Master:
        Create.Inactive_Master.append(item)

    #Assign list to Modify.Master. Here I could filter out weird groups I don't want showing up.
    Modify.Master=(DG_Groups+SG_Groups)
    Modify.Inactive_Master=Modify.Master

    Create.Default_Groups=['All-YMCA-Staff','SG-All-YMCA-Staff']
    Modify.Default_Groups=['All-YMCA-Staff','SG-All-YMCA-Staff']

    for Temp_Group in Create.Default_Groups:
        if not Temp_Group in Create.Active_Master:
            Create.Active_Master.append(Temp_Group)
        # else: 
        #     p rint(Temp_Group, 'is already in Create.Active_Master, and has not been added.')
        if Temp_Group in Create.Inactive_Master:
            Create.Inactive_Master.remove(Temp_Group)
#endregion Overhead

#region Functions

def Startup():  #Called after 'app=tkinterApp()' allowing other functions to work properly while still being positioned well. 
                #These things are the last things to actually happen when you run the program, until you hit buttons.
    try: open(r'c:\ACP\Query1.txt', 'x')
    except FileExistsError: pass
    try: open(r'c:\ACP\Query2.txt', 'x')
    except FileExistsError: pass
    try: open(r'c:\ACP\Query3.txt', 'x')
    except FileExistsError: pass
    try: open(r'c:\ACP\createLog.txt', 'x')
    except FileExistsError: pass
    try: open(r'c:\ACP\DirsyncUpdateLog.txt', 'x')
    except FileExistsError: pass
    try: open(r'c:\ACP\queryLog.txt', 'x')
    except FileExistsError: pass
    try: open(r'c:\ACP\modifyLog.txt', 'x')
    except FileExistsError: pass
    global loginName
    
    if not os.path.exists('c:\\ACP\\uName.txt'):
            with open('c:\\ACP\\uName.txt', 'w'): pass
    loginFile = open("c:\\ACP\\uName.txt","r+")
    loginName = loginFile.read()
    loginFile.close()
    #tk.messagebox.showinfo(title='Initialization Complete', message='Initialization Complete')
    Modify.Current_User_Status=''
    Modify.Current_User_Groups=[]
    Modify.Active_Group_Cache=[]  
    Modify.Current_Description=''
    Pull_Users()
    Pull_Groups()
    Modify.Name_Entry_Box.set_completion_list(User_List)
    Create.Inactive_Update()
    Modify.Inactive_Update()
    Create.Active_Update()
    Modify.Active_Update()
    if not os.path.exists('c:\\ACP\cred.xml'):
        with open('c:\\ACP\\cred.xml', 'w'): pass
        if not os.path.exists('c:\\ACP\credLog.txt'):
            with open('c:\\ACP\\credLog.txt', 'w'): pass
        global cred_time
        cred_time=pathlib.Path(r'c:\\ACP\\credLog.txt').stat().st_mtime
        loginNameBox = simpledialog.askstring(title="Input Username",prompt="Enter your username: (e.g: YMCA\\user-ad)",)
        if loginNameBox is not None:
            if not os.path.exists('c:\\ACP\\uName.txt'):
                with open('c:\\ACP\\uName.txt', 'w'): pass
            loginFile = open("c:\\ACP\\uName.txt","w+")
            loginFile.write(loginNameBox)
            loginFile.close()
            loginFile = open("c:\\ACP\\uName.txt","r+")
            loginName = loginFile.read()
            loginFile.close()
            #Run Powershell to prompt for password
            cred_time=pathlib.Path(r'c:\\ACP\\credLog.txt').stat().st_mtime
            
            if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName+'',r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit $credential = Get-Credential; $credential | Export-CliXml -Path "C:\ACP\cred.xml";Add-Content -Path c:\\ACP\\credLog.txt -Value "skjsdjkdsjksdfj";exit'])
            else: subprocess.call(['runas', '/savecred','/user:'+loginName+'',r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit $credential = Get-Credential; $credential | Export-CliXml -Path "C:\ACP\cred.xml";Add-Content -Path c:\\ACP\\credLog.txt -Value "skjsdjkdsjksdfj";'])
            def CredLoop(first_modified):
                if pathlib.Path(r'c:\\ACP\\credLog.txt').stat().st_mtime == first_modified:
                    app.after(1000,lambda : CredLoop(first_modified))
                else:
                    try: last_updated=open(r'c:\ACP\last_updated.txt', 'x')
                    except FileExistsError: pass
                    current_time=time.time()
                    last_updated=open(r'c:\ACP\last_updated.txt', 'r')
                    last_updated_time=last_updated.read()
                    if last_updated_time == '':last_updated_time=0
                    if float(current_time) - float(last_updated_time) >= 604800:
                        last_updated=open(r'c:\ACP\last_updated.txt', 'w')
                        last_updated.write(str(current_time))
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Staff,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportEnabled.txt;exit'])
                        else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Staff,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportEnabled.txt;'])
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Disabled-Users,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportDisabled.txt;exit'])
                        else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Disabled-Users,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportDisabled.txt;'])
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Security-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportSecurity.txt;exit'])
                        else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Security-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportSecurity.txt;'])
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Distribution-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportDistribution.txt;exit'])
                        else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Distribution-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportDistribution.txt;'])
                        path1=pathlib.Path(r'c:\ACP\userexportEnabled.txt')
                        path2=pathlib.Path(r'c:\ACP\userexportDisabled.txt')
                        path3=pathlib.Path(r'c:\ACP\groupExportSecurity.txt')
                        path4=pathlib.Path(r'c:\ACP\groupExportDistribution.txt')
                        def exportLoop(path1, path2, path3, path4, first_modified1, first_modified2, first_modified3, first_modified4, miliseconds=1000):
                            if path1.stat().st_mtime == first_modified1 or path2.stat().st_mtime == first_modified2 or path3.stat().st_mtime == first_modified3 or path4.stat().st_mtime == first_modified4:
                                app.after(1000,lambda : exportLoop(path1, path2, path3, path4, first_modified1, first_modified2, first_modified3, first_modified4, miliseconds=1000))
                            else:
                                Pull_Users()
                                Pull_Groups()
                                Modify.Name_Entry_Box.set_completion_list(User_List)
                        exportLoop(path1, path2, path3, path4, path1.stat().st_mtime, path2.stat().st_mtime, path3.stat().st_mtime, path4.stat().st_mtime,)
            CredLoop(cred_time)
    else:   
        try: last_updated=open(r'c:\ACP\last_updated.txt', 'x')
        except FileExistsError: pass
        current_time=time.time()
        last_updated=open(r'c:\ACP\last_updated.txt', 'r')
        last_updated_time=last_updated.read()
        if last_updated_time == '':last_updated_time=0
        if float(current_time) - float(last_updated_time) >= 604800:
            last_updated=open(r'c:\ACP\last_updated.txt', 'w')
            last_updated.write(str(current_time))
            if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Staff,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportEnabled.txt;exit'])
            else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Staff,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportEnabled.txt;'])
            if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Disabled-Users,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportDisabled.txt;exit'])
            else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Disabled-Users,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportDisabled.txt;'])
            if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Security-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportSecurity.txt;exit'])
            else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Security-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportSecurity.txt;'])
            if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Distribution-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportDistribution.txt;exit'])
            else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Distribution-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportDistribution.txt;'])
            path1=pathlib.Path(r'c:\ACP\userexportEnabled.txt')
            path2=pathlib.Path(r'c:\ACP\userexportDisabled.txt')
            path3=pathlib.Path(r'c:\ACP\groupExportSecurity.txt')
            path4=pathlib.Path(r'c:\ACP\groupExportDistribution.txt')
            def exportLoop(path1, path2, path3, path4, first_modified1, first_modified2, first_modified3, first_modified4, miliseconds=1000):
                if path1.stat().st_mtime == first_modified1 or path2.stat().st_mtime == first_modified2 or path3.stat().st_mtime == first_modified3 or path4.stat().st_mtime == first_modified4:
                    app.after(1000,lambda : exportLoop(path1, path2, path3, path4, first_modified1, first_modified2, first_modified3, first_modified4, miliseconds=1000))
                else:
                    Pull_Users()
                    Pull_Groups()
                    Modify.Name_Entry_Box.set_completion_list(User_List)
            exportLoop(path1, path2, path3, path4, path1.stat().st_mtime, path2.stat().st_mtime, path3.stat().st_mtime, path4.stat().st_mtime,)

def CredLoop(first_modified):
    if pathlib.Path(r'c:\\ACP\\credLog.txt').stat().st_mtime == first_modified:
            app.after(1000,lambda : CredLoop(first_modified))

def CredSetup(): #Prompt the user for credentials and save them
    global loginName
    loginNameBox = simpledialog.askstring(title="Input Username",prompt="Enter your username: (e.g: YMCA\\user-ad)",)
    if loginNameBox is not None:
        if not os.path.exists('c:\\ACP\\uName.txt'):
            with open('c:\\ACP\\uName.txt', 'w'): pass
        loginFile = open("c:\\ACP\\uName.txt","w+")
        loginFile.write(loginNameBox)
        loginFile.close()
        loginFile = open("c:\\ACP\\uName.txt","r+")
        loginName = loginFile.read()
        loginFile.close()
        #Run Powershell to prompt for password
        cred_time=pathlib.Path(r'c:\\ACP\\credLog.txt').stat().st_mtime
        CredLoop(cred_time)
        if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName+'',r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit $credential = Get-Credential; $credential | Export-CliXml -Path "C:\ACP\cred.xml";Add-Content -Path c:\\ACP\\credLog.txt -Value "skjsdjkdsjksdfj";exit'])
        else: subprocess.call(['runas', '/savecred','/user:'+loginName+'',r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit $credential = Get-Credential; $credential | Export-CliXml -Path "C:\ACP\cred.xml";Add-Content -Path c:\\ACP\\credLog.txt -Value "skjsdjkdsjksdfj";'])

def queryAccount(event=''): #Pull info about a user account
    #Grab the text box data and process it into a name.
    raw = Modify.Name_Entry_Box.get() #Raw is a local variable for the data pulled from the text entry box.
    

    if raw == "Enter the user's Full Name or their Username..." or raw == '': #Make sure the entry is not any of our "blank" states.
        tk.messagebox.showinfo(title='Not Enough Information', message='Enter something into the box...')
    elif raw == 'Skip' or raw == 'skip':
        Modify.Status_Identifier_Label.config(text='Enabled: ')
        Modify.Name_Identifier_Label.config(text='Name: ')
        Modify.Username_Identifier_Label.config(text='Username: ')
        Modify.Email_Identifier_Label.config(text='Email: ')
        Modify.Groups_Identifier_Label.config(text='Available Groups:                                 Current Groups: ')
        Modify.Name_Display_Label.config(text='Account Query Skipped')

        Modify.Inactive_Search.grid(row=10, column=0, columnspan=100, sticky='W', padx = 5)
        Modify.Inactive_Listbox.grid(row=11, column=0, columnspan=100, sticky='W', padx = 5)
        Modify.Active_Search.grid(row=10, column=1, columnspan=100, sticky='E', padx = 10)
        Modify.Active_Listbox.grid(row=11, column=1, columnspan=100, sticky='E', padx = 10)

        Modify.AdvancedLabel.grid(row=9, column=2, sticky='E',padx=1)
        Modify.AdvancedToggle.grid(row=9, column=3, sticky='W', padx=1)

        Modify.Auto_Label.grid(row=7, column=1, sticky="E")
        Modify.Auto_Toggle.grid(row=8, column=1, sticky='E')
        
    else:
        try:
            if '.' in raw:
                raw = raw.replace('.', '')
                raw_List=raw.split()
                print(raw)
            if ',' in raw:
                raw = raw.replace(',', '')
                raw_List=raw.split()
                if len(raw_List) == 1:
                    showwarning(title='Unable to Process', message='The system cannot process this account. Please try another name format.')
                if len(raw_List) == 2:  
                    Modify.Username=(raw_List[1][0:1]+raw_List[0]).lower()[:-1]
                    Modify.GivenName='\\"'+str(raw_List[1])+'\\"'
                    Modify.Surname='\\"'+str(raw_List[0])+'\\"'
                if len(raw_List) > 2:
                    Modify.Username=(raw_List[1][0:1]+raw_List[2]+raw_List[0]).lower()
                    Modify.GivenName='\\"'+str(raw_List[1])+'\\"'
                    Modify.Surname='\\"'+str(raw_List[0])+'\\"'
            else:
                raw_List=raw.split()
                if len(raw_List) == 1:
                    Modify.Username=raw_List[0].lower()
                if len(raw_List) == 2:
                    Modify.Username=(raw_List[0][0:1]+raw_List[1]).lower()
                    Modify.GivenName='\\"'+str(raw_List[0])+'\\"'
                    Modify.Surname='\\"'+str(raw_List[1])+'\\"'
                if len(raw_List) > 2:
                    Modify.Username=(raw_List[0][0:1]+raw_List[1]+raw_List[2]).lower()
                    Modify.GivenName='\\"'+str(raw_List[0])+'\\"'
                    Modify.Surname='\\"'+str(raw_List[2])+'\\"'


        except:
            tk.messagebox.showerror(title='Error', message='This entry cannot be processed.')
            return

        
        Modify.Query1 = open(r'c:\ACP\Query1.txt', "w"); Modify.Query1.close(); Modify.Query2 = open(r'c:\ACP\Query2.txt', "w"); Modify.Query2.close(); Modify.Query3 = open(r'c:\ACP\Query3.txt', "w"); Modify.Query3.close()
    
        
        
        #Trigger lockout to stop calling multiple instances: buttonName is passed from whatever originally triggered the search, and it tells the program what button should be locked out.
        Create.Create_User_Button.config(command=(lambda : tk.messagebox.showwarning(title='Unable to Proceed.',message='Another instance is already running.'))) #Lock the user out of creating new accounts.
        Modify.Search_User_Button.config(command=(lambda : tk.messagebox.showwarning(title='Unable to Proceed.',message='Another instance is already running.'))) #Also lock the user out of querying more account modifications, just to be safe.
        Modify.Loading_Indicator.grid(row=2, column=1)

        modify_path=pathlib.Path(r'c:\ACP\Query1.txt')
        #Run database search.
        if autoclose.get() == True: 
            try: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; Get-ADUser -Filter \'GivenName -eq '+Modify.GivenName+' -and Surname -eq '+Modify.Surname+'\' | Out-File -FilePath c:\\ACP\\Query1.txt -Encoding ASCII;exit'])
            except: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; Get-ADUser –Identity \''+Modify.Username+'\' | Out-File -FilePath c:\\ACP\\Query1.txt -Encoding ASCII;exit'])
        else: 
            try: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; Get-ADUser -Filter \'GivenName -eq '+Modify.GivenName+' -and Surname -eq '+Modify.Surname+'\' | Out-File -FilePath c:\\ACP\\Query1.txt -Encoding ASCII;'])
            except: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; Get-ADUser –Identity \''+Modify.Username+'\' | Out-File -FilePath c:\\ACP\\Query1.txt -Encoding ASCII;'])

        def queryLoop(path, where, first_modified, miliseconds=1000): #This function serves to handle all parts of the modification process that require waiting for PS functions to be completed.
        #Arguments:
        #path: The file path to check for updates in. When looping, it will loop until this file is changed.
        #where: This is a keyword that can be any string, as long as it matches one of the options below. This tells the function what actions to do after the file is updated.
        #first_modified: This data is fed to the function as a timestamp against which to cross-reference the current latest time the file was updated. When that time changes, actions will continue.
        #miliseconds: How long to wait before querying the file. By default, 1 second. Most queries will use the default, and this parameter can be omitted.

            if path.stat().st_mtime == first_modified:
                app.after(1000,lambda : queryLoop(path, where, first_modified, miliseconds=1000))
                Modify.Loading_Indicator.config(text='Loading user info...')
                app.after(333,lambda:Modify.Loading_Indicator.config(text='Loading user info.'))
                app.after(666,lambda:Modify.Loading_Indicator.config(text='Loading user info..'))
            else:
                if where == 'account_information_pulled':
                    if autoclose.get() == True: 
                        try: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; (Get-ADUser -Filter \'GivenName -eq \"'+Modify.GivenName+'\" -and Surname -eq \"'+Modify.Surname+'"\' –Properties MemberOf).MemberOf| Out-File -FilePath c:\\ACP\\Query2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\queryLog.txt -Value "skjsdjkdsjksdfj";exit'])
                        except: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; (Get-ADUser –Identity \''+Modify.Username+'\' –Properties MemberOf).MemberOf| Out-File -FilePath c:\\ACP\\Query2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\queryLog.txt -Value "skjsdjkdsjksdfj";exit'])
                    else:
                        try: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; (Get-ADUser -Filter \'GivenName -eq \"'+Modify.GivenName+'\" -and Surname -eq \"'+Modify.Surname+'"\' –Properties MemberOf).MemberOf| Out-File -FilePath c:\\ACP\\Query2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\queryLog.txt -Value "skjsdjkdsjksdfj";'])
                        except: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; (Get-ADUser –Identity \''+Modify.Username+'\' –Properties MemberOf).MemberOf| Out-File -FilePath c:\\ACP\\Query2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\queryLog.txt -Value "skjsdjkdsjksdfj";'])
                    queryLoop(pathlib.Path('c:\\ACP\\queryLog.txt'), 'groups pulled', pathlib.Path('c:\\ACP\\Query2.txt').stat().st_mtime)
                if where == 'groups pulled':
                    if autoclose.get() == True: 
                        try: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; Get-ADUser -Filter \'GivenName -eq \"'+Modify.GivenName+'\" -and Surname -eq \"'+Modify.Surname+'"\' -Properties Description | Select-Object -ExpandProperty Description | Out-File -FilePath c:\\ACP\\Query3.txt -Encoding ASCII;exit'])
                        except: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; Get-ADUser –Identity \''+Modify.Username+'\' -Properties Description | Select-Object -ExpandProperty Description | Out-File -FilePath c:\\ACP\\Query3.txt -Encoding ASCII;exit'])
                    else: 
                        try: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; Get-ADUser -Filter \'GivenName -eq \"'+Modify.GivenName+'\" -and Surname -eq \"'+Modify.Surname+'"\' -Properties Description | Select-Object -ExpandProperty Description | Out-File -FilePath c:\\ACP\\Query3.txt -Encoding ASCII;'])
                        except: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; Get-ADUser –Identity \''+Modify.Username+'\' -Properties Description | Select-Object -ExpandProperty Description | Out-File -FilePath c:\\ACP\\Query3.txt -Encoding ASCII;'])
                    queryLoop(pathlib.Path('c:\\ACP\\Query3.txt'), 'Description Pulled', pathlib.Path('c:\\ACP\\Query3.txt').stat().st_mtime)
                if where == 'Description Pulled':

                    Modify.Query1 = open(r'c:\ACP\Query1.txt', "r")
                    Modify.Query1_Splitlines = Modify.Query1.read().splitlines()

                    Modify.Query2 = open(r'c:\ACP\Query2.txt', "r")
                    Modify.Query2_Splitlines = Modify.Query2.read().splitlines()

                    Modify.Query3 = open(r'c:\ACP\Query3.txt', "r")
                    Modify.Query3_Splitlines = Modify.Query3.read().splitlines()
                    
                    
                    try: #If this condition does not encounter an error, then the account exists.
                        indexCount = 0
                        #Iterate through the lines of the file to isolate the groups of the user. 
                        for i in Modify.Query2_Splitlines:
                            Modify.Query2_Splitlines[indexCount]=Modify.Query2_Splitlines[indexCount].split(',')[0].split('=')[1]
                            indexCount += 1
                        Modify.Active_Master=Modify.Query2_Splitlines
                        if Modify.Active_Master==[]:
                            Modify.Query2 = open(r'c:\ACP\Query2.txt', "r")
                            Modify.Query2_Splitlines = Modify.Query2.read().splitlines()
                            for i in Modify.Query2_Splitlines:
                                Modify.Query2_Splitlines[indexCount]=Modify.Query2_Splitlines[indexCount].split(',')[0].split('=')[1]
                                indexCount += 1
                            Modify.Active_Master=Modify.Query2_Splitlines

                        Modify.Description = Modify.Query3_Splitlines[0]
                        Modify.Description_Entry.delete(0,END)
                        Modify.Description_Entry.insert(0,Modify.Description)
                        Modify.Active_Update()
                        
                        Modify.Inactive_Master=[]
                        for item in Modify.Master:
                            Modify.Inactive_Master.append(item)
                        for item in Modify.Presets.Master:
                            Modify.Inactive_Master.append(item)
                        
                        for item in Modify.Active_Master:
                            if item in Modify.Inactive_Master:
                                Modify.Inactive_Master.remove(item)
                                Modify.Inactive_Update()




                    
                        if Modify.Query1_Splitlines[3].split()[2] == 'True': #Account is enabled
                            statVar.set(1) #Set value to True to show account is enabled
                        elif Modify.Query1_Splitlines[3].split()[2] == 'False': #Account is disabled
                            statVar.set(0) #Set value to False to show account is disabled
                        else:
                            pass #p rint('There was an error in the enabled/disabled system')
                        
                        
                        try: 
                            Modify.Name_Display_Label.config(text=Modify.Query1_Splitlines[5].split()[2]+' '+Modify.Query1_Splitlines[5].split()[3]+' '+Modify.Query1_Splitlines[5].split()[4])
                        except IndexError:
                            Modify.Name_Display_Label.config(text=Modify.Query1_Splitlines[5].split()[2]+' '+Modify.Query1_Splitlines[5].split()[3])
                        Modify.Username_Display_Label.config(text=Modify.Query1_Splitlines[8].split()[2])
                        Modify.Email_Display_Label.config(text=Modify.Query1_Splitlines[11].split()[2])

                        #Modify.Inactive_Search.grid(row=8, column=1)
                        Modify.Inactive_Listbox.grid(row=9, column=1)
                        #Modify.Active_Search.grid.grid(row=8, column=2)
                        Modify.Active_Listbox.grid(row=9, column=2)

                        

                    except IndexError: #If there is an IndexError, then the file is (usually) empty, signifying the account does not exist. 
                        statVar.set(0) #Set the value to false, but this is not displayed anywhere currently.
                        successStatus=False
                        tk.messagebox.showerror(title='Account Not Found',message='This account does not appear to exist. This may be an error. Try querying the account again, or create the account instead. If the issue persists, contact IT.')

                        
                    #Make visible the text that hosts the information about the user.
                    Modify.Status_Button.grid(row=3, column = 1, sticky='W') #Enable the status button.
                    Modify.Status_Identifier_Label.config(text='Enabled: ') #Is the account enabled or not?
                    Modify.Name_Identifier_Label.config(text='Name: ')
                    Modify.Username_Identifier_Label.config(text='Username: ')
                    Modify.Email_Identifier_Label.config(text='Email: ')
                    Modify.Groups_Identifier_Label.config(text='Available Groups:                                 Current Groups: ') #This might be removed later, depending on the new groups system.


                    Modify.Loading_Indicator.grid_forget() #Also hide the loading label now.
                    Modify.Search_User_Button.config(command = queryAccount) #At the end of the process, reenable the button
                    Create.Loading_Indicator.config(text='')
                    Create.Create_User_Button.config(command = lambda : createAccount())


                    Modify.Inactive_Search.grid(row=10, column=0, columnspan=100, sticky='W', padx = 5)
                    Modify.Inactive_Listbox.grid(row=11, column=0, columnspan=100, sticky='W', padx = 5)
                    Modify.Active_Search.grid(row=10, column=1, columnspan=100, sticky='E', padx = 10)
                    Modify.Active_Listbox.grid(row=11, column=1, columnspan=100, sticky='E', padx = 10)

                    Modify.AdvancedLabel.grid(row=9, column=2, sticky='E',padx=1)
                    Modify.AdvancedToggle.grid(row=9, column=3, sticky='W', padx=1)

                    Modify.Auto_Label.grid(row=6, column=1, sticky="SE")
                    Modify.Auto_Toggle.grid(row=8, column=1, sticky='E')
                    Modify_Desc_Lock()
                    #Save the info that the pushChanges func will check against.
                    Modify.Current_User_Status=statVar.get()
                    
                    

                    Modify.Push_Changes_Button.grid(row=12, column=0, padx=5, pady=5, sticky="W", columnspan=100)

                    Lock_Groups()
                    
                    Modify.Description_Label.grid(row=8, column=0, sticky="W")
                    
                    Modify.Description_Entry.grid(row=8, column=1, sticky='W')
                    Modify.Groups_to_remove=[]
                    Modify.Presets.Active=[]
                    for groups in Modify.Presets.Corresponding:
                        for item in groups:
                            Modify.Success=True
                            if not item in Modify.Active_Master:
                                Modify.Success=False
                                break
                        if Modify.Success == True:
                            Modify.Presets.Active.append(Modify.Presets.Master[Modify.Presets.Corresponding.index(groups)])
                            for individual_group in groups:
                                Modify.Groups_to_remove.append(individual_group)
                    for preset in Modify.Presets.Active:
                        Modify.Inactive_to_Active(event='', selected_group=preset)
                    for group in Modify.Groups_to_remove:
                        try: Modify.Active_to_Inactive(event='', selected_group=group)
                        except: pass
                    
                    Modify.Current_User_Groups=Modify.Current_User_Groups+Modify.Active_Master
                    
                    Modify.Current_Description=''
                    if Modify.Auto_Mode.get()==True:
                        Modify.Description_Entry.config(state="normal")
                        Modify.Current_Description=Modify.Current_Description+Modify.Description_Entry.get()
                        Modify.Description_Entry.config(state="disabled")
                    else: Modify.Current_Description=Modify.Current_Description+Modify.Description_Entry.get()

        queryLoop(modify_path, 'account_information_pulled', modify_path.stat().st_mtime)

def pushChanges(event=''):
    Modify.Description=Modify.Description_Entry.get()
    Modify.Groups_Queue=[]
    Modify.Presets_Queue=[]

    groupsAdded=[]
    groupsRemoved=[]
    groupsAddedBool=False
    groupsRemovedBool=False
    statChangedBool=False
    descChangedBool=False


    message='Confirm you would like to make these changes to the user account:\n\n'
    #Get status of user and user groups from end of query account.
    if statVar.get() != Modify.Current_User_Status or Modify.Active_Master != Modify.Current_User_Groups or Modify.Description != Modify.Current_Description:
        if statVar.get() != Modify.Current_User_Status:
            if str(Modify.Current_User_Status) == '0':
                original_status = 'Disabled'
            else:
                original_status = 'Enabled'
            if str(statVar.get()) == '0':
                new_status = 'Disabled'
            else:
                new_status = 'Enabled'
            message=message+'User status: '+original_status+' -> '+new_status+'\n\n'
            statChangedBool=True
        
        
        if Modify.Active_Master != Modify.Current_User_Groups:
            for item in Modify.Current_User_Groups:
                if item not in Modify.Active_Master:
                    if '(Preset)' in item:
                        #Find the group in the list of presets, take the index number of where that group locates, and then check the same index number in the list of groups for each preset.
                        for group in Modify.Presets.Corresponding[Modify.Presets.Master.index(item)]:
                            #Exclude blank presets to prevent unnecessary errors
                            if group != '':
                                groupsRemoved.append(group)
                    else: groupsRemoved.append(item)
            for item in Modify.Active_Master:
                if item not in Modify.Current_User_Groups:
                    groupsAdded.append(item)
            if groupsAdded != []:
                message=message+'Groups Added: '+str(groupsAdded)+'\n\n'
                groupsAddedBool=True
            if groupsRemoved != []:
                message=message+'Groups Removed: '+str(groupsRemoved)+'\n\n'
                groupsRemovedBool=True
        
        if Modify.Description != Modify.Current_Description:
            message=message+'Description: "'+Modify.Current_Description+'"  -> "'+Modify.Description+'"'
            descChangedBool=True
        pushYN = askokcancel(title='Confirm Changes', message=message)

        if pushYN==True:
            #Update the user's settings
            Modify.Current_User_Status=''
            Modify.Current_User_Groups=[]
            Modify.Current_User_Status=statVar.get()
            Modify.Current_User_Groups=Modify.Current_User_Groups+Modify.Active_Master
            try:Modify.Remove_Groups ='$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;Get-ADUser -Filter \'GivenName -eq \"'+Modify.GivenName+'\" -and Surname -eq \"'+Modify.Surname+'"\' -Properties MemberOf | ForEach-Object {  $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false}'
            except:Modify.Remove_Groups ='$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;Get-ADUser -Identity '+Modify.Username+' -Properties MemberOf | ForEach-Object {  $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false}'
            #except:Modify.Remove_Groups ='$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;Get-ADUser -Identity '+Modify.Username+' -Properties MemberOf | ForEach-Object {  $_.MemberOf | Remove-ADGroupMember -Members $_.DistinguishedName -Confirm:$false}'
            Modify.Add_Groups=''
            try:Modify.Move_Disabled="$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \""+Modify.GivenName+"\" -and Surname -eq \""+Modify.Surname+"\"\'; Move-ADObject -Identity $user.DistinguishedName -TargetPath \'OU=Disabled-Users,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\';Disable-ADAccount -Identity $user.SamAccountName"
            except:Modify.Move_Disabled="$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Identity "+Modify.Username+"; Move-ADObject -Identity $user.DistinguishedName -TargetPath \'OU=Disabled-Users,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\';Disable-ADAccount -Identity $user.SamAccountName"
            try:Modify.Move_Enabled='$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \"'+Modify.GivenName+'\" -and Surname -eq \"'+Modify.Surname+'"\' ;Move-ADObject -Identity $user.DistinguishedName -TargetPath \'OU=Staff,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\';Enable-ADAccount -Identity $user.DistinguishedName'
            except:Modify.Move_Enabled='$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Identity '+Modify.Username+' ;Move-ADObject -Identity $user.DistinguishedName -TargetPath \'OU=Staff,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\';Enable-ADAccount -Identity $user.DistinguishedName'
            try:Modify.Strip_License="$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \""+Modify.GivenName+"\" -and Surname -eq \""+Modify.Surname+"\"\';Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses \"monroecountyymca:STANDARDWOFFPACK\";"
            except:Modify.Strip_License="$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Identity "+Modify.Username+";Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -RemoveLicenses \"monroecountyymca:STANDARDWOFFPACK\";"
            try:Modify.Add_License="$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \""+Modify.GivenName+"\" -and Surname -eq \""+Modify.Surname+"\"\';Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses \"monroecountyymca:STANDARDWOFFPACK\";"
            except:Modify.Add_License="$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Identity "+Modify.Username+";Set-MsolUserLicense -UserPrincipalName $user.UserPrincipalName -AddLicenses \"monroecountyymca:STANDARDWOFFPACK\";"
            try:Modify.Change_Password="$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \""+Modify.GivenName+"\" -and Surname -eq \""+Modify.Surname+"\"\';Set-ADAccountPassword -Identity $user.SamAccountName -NewPassword (ConvertTo-SecureString -String \'"+(''.join(random.choice(string.ascii_letters + string.digits) for i in range(12)) + "!")+"\' -AsPlainText -Force);"
            except:Modify.Change_Password="$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Identity "+Modify.Username+";Set-ADAccountPassword -Identity $user.SamAccountName -NewPassword (ConvertTo-SecureString -String \'"+(''.join(random.choice(string.ascii_letters + string.digits) for i in range(12)) + "!")+"\' -AsPlainText -Force);"
            try:Modify.Change_Description='$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \"'+Modify.GivenName+'\" -and Surname -eq \"'+Modify.Surname+'\"\';Set-ADUser -Identity $user.DistinguishedName -Description \''+Modify.Description+'\''
            except:Modify.Change_Description='$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Identity '+Modify.Username+';Set-ADUser -Identity $user.DistinguishedName -Description \''+Modify.Description+'\''
            
            if statChangedBool==True and statVar.get()==False: #If the status has changed, and is now Disabled:
                #Remove the groups, remove the license, change the password, and move to disabled
                if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Remove_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Removed_Groups";exit'])
                else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Remove_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Removed_Groups";'])
                def Modify_Loop_1(path, where, first_modified, miliseconds=1000): #The following must be done in a function, since it waits for files to be checked on. 
                    #If the path given is equal to the time given, check the path again, after the time provided, default is 1 second.
                    if path.stat().st_mtime == first_modified:
                        app.after(miliseconds,lambda : Modify_Loop_1(path, where, first_modified))
                    else:
                        if where == "Groups Removal Completed":
                            if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Strip_License+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Removed_License";exit'])
                            else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Strip_License+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Removed_License";'])
                            Modify_Loop_1(pathlib.Path(r'c:\ACP\modifyLog.txt'), "License Removed", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)
                        if where == "License Removed":
                            if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Change_Password+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Changed_Password";exit'])                            
                            else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Change_Password+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Changed_Password";'])                            
                            Modify_Loop_1(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Password Changed", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)
                        if where == "Password Changed":
                            if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Move_Disabled+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Moved_to_Disabled_Users";exit'])                            
                            else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Move_Disabled+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Moved_to_Disabled_Users";'])                            
                            Modify_Loop_1(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Moved to Disabled", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)
                        if where == "Moved to Disabled":
                            pass #p rint('done')
                
                Modify_Loop_1(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Groups Removal Completed", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)

            elif statChangedBool==True and statVar.get()==True: #If the status has changed, and is now Enabled:
                #Enable the account and move it
                if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Move_Enabled+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Moved_Enabled";exit'])
                else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Move_Enabled+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Moved_Enabled";'])
                #Prepare the groups, and add in the presets.
                Modify.Groups_Queue=Modify.Groups_Queue+Modify.Active_Master
                for group in Modify.Active_Master:
                    #if the group is a preset
                    if '(Preset)' in group:
                        #Find the group in the list of presets, take the index number of where that group locates, and then check the same index number in the list of groups for each preset.
                        for item in Modify.Presets.Corresponding[Modify.Presets.Master.index(group)]:
                            #Exclude blank presets to prevent unnecessary errors
                            if item != '':
                                #Add the groups to the master list
                                Modify.Groups_Queue.append(item)
                        #remove the group that contains '(Preset)'
                        Modify.Groups_Queue.remove(group)      
                def Modify_Loop_2(path, where, first_modified, miliseconds=1000): #The following must be done in a function, since it waits for files to be checked on. 
                    #If the path given is equal to the time given, check the path again, after the time provided, default is 1 second.
                    if path.stat().st_mtime == first_modified:
                        app.after(miliseconds,lambda : Modify_Loop_2(path, where, first_modified))
                    else:
                        if where == "Moved to Enabled" and len('"'+'\',\''.join(Modify.Groups_Queue)+'"|Add-ADGroupMember -Members "'+Modify.Username+'"') > 600:
                            #Add the groups    
                            
                            Modify.All_Groups='\',\''.join(Modify.Groups_Queue)
                            groupsPart1='\',\''.join(Modify.Groups_Queue[:len(Modify.Groups_Queue)//2])
                            groupsPart2='\',\''.join(Modify.Groups_Queue[len(Modify.Groups_Queue)//2:])
                            Modify.Add_Groups = "$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \""+Modify.GivenName+"\" -and Surname -eq \""+Modify.Surname+"\"\';\'"+groupsPart1+"\'|Add-ADGroupMember -Members $user.SamAccountName"
                            Modify.Add_Groups_Part_2 = "$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \""+Modify.GivenName+"\" -and Surname -eq \""+Modify.Surname+"\"\';\'"+groupsPart2+"\'|Add-ADGroupMember -Members $user.SamAccountName"
                            if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Added_Groups_Part_1";exit'])
                            else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Added_Groups_Part_1";'])
                            Modify_Loop_2(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Added_Groups_Part_1", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime) 
                        if where == "Added_Groups_Part_1":
                            #add the second part of groups
                            if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups_Part_2+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Added_Groups_Part_2";exit'])
                            else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups_Part_2+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Added_Groups_Part_2";'])
                            Modify_Loop_2(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Added_Groups_Final", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime) 
                        if where == "Moved to Enabled" and not len('"'+'\',\''.join(Modify.Groups_Queue)+'"|Add-ADGroupMember -Members "'+Modify.Username+'"') > 600:
                            #Add groups if less than 600
                            Modify.All_Groups='\',\''.join(Modify.Groups_Queue)
                            Modify.Add_Groups ="$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \""+Modify.GivenName+"\" -and Surname -eq \""+Modify.Surname+"\"\';\'"+Modify.All_Groups+"\'|Add-ADGroupMember -Members $user.SamAccountName "
                            if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Added";exit'])                            
                            else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Added";'])                            
                            Modify_Loop_2(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Added_Groups_Final", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)
                        if where == "Added_Groups_Final":
                            #Start the dirsync to the cloud. Returns when we get confirmation the sync has started. 
                            if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit Invoke-Command -ComputerName dirsync -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test";exit'])
                            else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit Invoke-Command -ComputerName dirsync -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test";'])
                            Modify_Loop_2(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'dirsync_started', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime)
                        
                        if where == 'dirsync_started':
                            #Pull information as to the last time the sync occurred, which is equivalent to the first_modified time on the python side.
                            if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test2";exit'])
                            else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test2";'])
                            Modify_Loop_2(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'original_sync_time_query', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime)
                        
                        if where == 'original_sync_time_query':
                            #Once we get the time to hold as the original, we can begin checking the time again every ten seconds, to check for a change.
                            ResultsFile = open(r'c:\ACP\DirsyncUpdateLog2.txt', "r")
                            Lines = ResultsFile.read().splitlines()
                            Create.OriginalSyncTime=Lines[20]
                            if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";exit'])
                            else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";'])
                            Modify_Loop_2(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'new_sync_time_query', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime,10000)
                        
                        if where == 'new_sync_time_query':
                            #Check if the time is different or not. If it is not, the program runs again, returning to this block. If it is different, it proceeds.
                            ResultsFile = open(r'c:\ACP\DirsyncUpdateLog2.txt', "r")
                            Lines = ResultsFile.read().splitlines()
                            if Lines[20]==Create.OriginalSyncTime:
                                if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";exit'])
                                else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";'])
                                Modify_Loop_2(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'new_sync_time_query', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime,10000)
                            else:
                                #Add the license
                                if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_License+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Added_License";exit'])
                                else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_License+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Added_License";'])
                                Modify_Loop_2(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Added License", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime) 
                        if where == "Added License":
                            pass
                            #p rint('Done')                
                Modify_Loop_2(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Moved to Enabled", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)        
                return

            def Modify_Remove_Groups():
                Modify.All_Groups_Removal='\',\''.join(groupsRemoved)
                if len('"'+Modify.All_Groups_Removal+'"|Remove-ADGroupMember -Members "'+Modify.Username+'" -Confirm:$false') > 600:
                    part1=groupsRemoved[:len(groupsRemoved)//2]
                    part2=groupsRemoved[len(groupsRemoved)//2:]
                    groupsRemovedPart1='\',\''.join(part1)
                    groupsRemovedPart2='\',\''.join(part2)
                    Modify.Remove_Groups = '$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \"'+Modify.GivenName+'\" -and Surname -eq \"'+Modify.Surname+'"\' ; \''+groupsRemovedPart1+'\'|Remove-ADGroupMember -Members $user.DistinguishedName -Confirm:$false'
                    Modify.Remove_Groups_Part_2 = '$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \"'+Modify.GivenName+'\" -and Surname -eq \"'+Modify.Surname+'"\' ; \''+groupsRemovedPart2+'\'|Remove-ADGroupMember -Members $user.DistinguishedName -Confirm:$false'
                    if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Modify.Remove_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Removed";exit'])                            
                    else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Modify.Remove_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Removed";'])                            

                    def Modify_Loop_5(path, where, first_modified, miliseconds=1000): #The following must be done in a function, since it waits for files to be checked on. 
                        #If the path given is equal to the time given, check the path again, after the time provided, default is 1 second.
                        if path.stat().st_mtime == first_modified:
                            app.after(miliseconds,lambda : Modify_Loop_5(path, where, first_modified))
                        else:
                            if where == "Groups Removed":
                                if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Remove_Groups_Part_2+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Removed_Part_2";exit'])                            
                                else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Remove_Groups_Part_2+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Removed_Part_2";'])                            
                                Modify_Loop_5(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Groups Part 2 Removed", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)
                            if where == "Groups Part 2 Removed":
                                pass
                    Modify_Loop_5(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Groups Removed", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)
                    
                else:
                    Modify.Remove_Groups ='$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \"'+Modify.GivenName+'\" -and Surname -eq \"'+Modify.Surname+'"\' ; \''+Modify.All_Groups_Removal+'\'|Remove-ADGroupMember -Members $user.DistinguishedName -Confirm:$false'
                    if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Remove_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Removed";exit'])                            
                    else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Remove_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Removed";'])                            

                    def Modify_Loop_6(path, where, first_modified, miliseconds=1000): #The following must be done in a function, since it waits for files to be checked on. 
                        #If the path given is equal to the time given, check the path again, after the time provided, default is 1 second.
                        if path.stat().st_mtime == first_modified:
                            app.after(miliseconds,lambda : Modify_Loop_6(path, where, first_modified))
                        else:
                            if where == "Groups Removed":
                                if statChangedBool==False and statVar.get()==True and groupsAddedBool==True:
                                    Modify_Add_Groups()
                    Modify_Loop_6(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Groups Removed", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)

            def Modify_Add_Groups():
                Modify.Groups_Queue=Modify.Groups_Queue+Modify.Active_Master
                for group in Modify.Active_Master:
                    #if the group is a preset
                    if '(Preset)' in group:
                        #Find the group in the list of presets, take the index number of where that group locates, and then check the same index number in the list of groups for each preset.
                        for item in Modify.Presets.Corresponding[Modify.Presets.Master.index(group)]:
                            #Exclude blank presets to prevent unnecessary errors
                            if item != '':
                                #Add the groups to the master list
                                Modify.Groups_Queue.append(item)
                        #remove the group that contains '(Preset)'
                        Modify.Groups_Queue.remove(group)
                Modify.All_Groups='\',\''.join(Modify.Groups_Queue)
                if len('"'+Modify.All_Groups+'"|Add-ADGroupMember -Members "'+Modify.Username+'"') > 600:
                    part1=Modify.Groups_Queue[:len(Modify.Groups_Queue)//2]
                    part2=Modify.Groups_Queue[len(Modify.Groups_Queue)//2:]
                    groupsPart1='\',\''.join(part1)
                    groupsPart2='\',\''.join(part2)
                    Modify.Add_Groups = "$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \""+Modify.GivenName+"\" -and Surname -eq \""+Modify.Surname+"\"\';\'"+groupsPart1+"\'|Add-ADGroupMember -Members $user.SamAccountName"
                    Modify.Add_Groups_Part_2 = "$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \""+Modify.GivenName+"\" -and Surname -eq \""+Modify.Surname+"\"\';\'"+groupsPart2+"\'|Add-ADGroupMember -Members $user.SamAccountName"
                    if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Added";exit'])                            
                    else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Added";'])                            

                    def Modify_Loop_3(path, where, first_modified, miliseconds=1000): #The following must be done in a function, since it waits for files to be checked on. 
                        #If the path given is equal to the time given, check the path again, after the time provided, default is 1 second.
                        if path.stat().st_mtime == first_modified:
                            app.after(miliseconds,lambda : Modify_Loop_3(path, where, first_modified))
                        else:
                            if where == "Groups Added":
                                if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups_Part_2+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Added_Part_2";exit'])                            
                                else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups_Part_2+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Added_Part_2";'])                            
                                Modify_Loop_3(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Groups Part 2 Added", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)
                            if where == "Groups Part 2 Added":
                                pass
                                #p rint("The Groups Have Been Added")
                    Modify_Loop_3(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Groups Added", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)
                    
                else:
                    Modify.Add_Groups = "$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;$user = Get-ADUser -Filter \'GivenName -eq \""+Modify.GivenName+"\" -and Surname -eq \""+Modify.Surname+"\"\';\'"+Modify.All_Groups+"\'|Add-ADGroupMember -Members $user.SamAccountName"
                    if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Added";exit'])                            
                    else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Add_Groups+'; Add-Content -Path c:\\ACP\\modifyLog.txt -Value "Groups_Added";'])                            

                    def Modify_Loop_4(path, where, first_modified, miliseconds=1000): #The following must be done in a function, since it waits for files to be checked on. 
                        #If the path given is equal to the time given, check the path again, after the time provided, default is 1 second.
                        if path.stat().st_mtime == first_modified:
                            app.after(miliseconds,lambda : Modify_Loop_4(path, where, first_modified))
                        else:
                            if where == "Groups Added":
                                pass
                                #p rint('Done')
                    
                    Modify_Loop_4(pathlib.Path(r'c:\ACP\modifyLog.txt'), "Groups Added", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)
            if statChangedBool==False and statVar.get()==True and groupsRemovedBool==True: #If the account's status is enabled, is remaining enabled, and groups were removed:
                #If we removed groups, then run the whole group removal system. This algorithm will also run the group add commands afterwards if the criteria is met.
                Modify_Remove_Groups()
            if statChangedBool==False and statVar.get()==True and groupsRemovedBool==False and groupsAddedBool==True:
                #If no groups were removed, but some were added, then add groups.
                Modify_Add_Groups()         

            if descChangedBool==True:
                if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Change_Description+'; Add-Content -Path c:\ACP\modifyLog.txt -Value "Added_Description";exit'])
                else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Modify.Change_Description+'; Add-Content -Path c:\ACP\modifyLog.txt -Value "Added_Description";'])
                def Modify_Loop_7(path, where, first_modified, miliseconds=1000): #The following must be done in a function, since it waits for files to be checked on. 
                    #If the path given is equal to the time given, check the path again, after the time provided, default is 1 second.
                    if path.stat().st_mtime == first_modified:
                        app.after(miliseconds,lambda : Modify_Loop_7(path, where, first_modified))
                    else:
                        Modify.Current_Description=''
                        Modify.Current_Description=Modify.Current_Description+Modify.Description
                Modify_Loop_7(pathlib.Path(r'c:\ACP\modifyLog.txt'), "", pathlib.Path(r'c:\ACP\modifyLog.txt').stat().st_mtime)

def Lock_Groups():
    if statVar.get()==False:
        Modify.Active_Group_Cache=[]
        Modify.Active_Group_Cache=Modify.Active_Group_Cache+Modify.Active_Master
        #Here we need to add all the old groups to the inactive side.
        for item in Modify.Active_Group_Cache:
            if item not in Modify.Default_Groups:
                Modify.Active_to_Inactive(event='',selected_group=item)

        Modify.Active_Master=[]
        Modify.Active_Master=Modify.Active_Master+Modify.Default_Groups
        Modify.Active_Update()
        Modify.Inactive_Update()
        Modify.Active_Listbox.config(state="disabled")
        Modify.Inactive_Listbox.config(state="disabled")
        Modify.Active_Listbox.bind('<Double-Button-1>', lambda event: "break")
        Modify.Inactive_Listbox.bind('<Double-Button-1>', lambda event: "break")
    if statVar.get()==True:
        for item in Modify.Active_Group_Cache:
            if item not in Modify.Default_Groups:
                Modify.Inactive_to_Active(event='',selected_group=item)
        Modify.Active_Group_Cache=[]
        Modify.Active_Listbox.config(state="normal")
        Modify.Inactive_Listbox.config(state="normal")
        Modify.Active_Listbox.bind('<Double-Button-1>', Modify.Active_to_Inactive)
        Modify.Inactive_Listbox.bind('<Double-Button-1>', Modify.Inactive_to_Active)
        Modify.Active_Update()
        Modify.Inactive_Update()==False

def createAccount(): #Creates the entire account for a new user
    Create.Groups_Queue=[]
    Create.Presets_Queue=[]
    Create.RepCount=0
    
    if Create.Supervisor.get() == "Other":
        result = simpledialog.askstring(title="Input Supervisor",prompt="Enter the user's Supervisor:",)
        Create.Supervisor=StringVar(value = result) 
    answer = askokcancel(title='Run Command?', message='Confirm you would like to create the user.') #Confirmation is key. Makes a pop-up message asking the user to confirm that they wish to make the account.
    if answer == True:
        #Check for duplicates
        if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; Get-ADUser -Filter \'Name -eq '+('\\"'+Create.Fullname+'\\"')+'\' | Out-File -FilePath c:\\ACP\\Query1.txt -Encoding ASCII;exit'])
        else: subprocess.call(['runas', '/savecred','/user:'+loginName,'C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \"; Get-ADUser -Filter \'Name -eq '+('\\"'+Create.Fullname+'\\"')+'\' | Out-File -FilePath c:\\ACP\\Query1.txt -Encoding ASCII;'])
        Create.Create_User_Button.config(command='') #Lock the user out of creating new accounts.
        Modify.Search_User_Button.config(command='') #Also lock the user out of querying more account modifications, just to be safe.
        #Define path to the search results file.
        create_path=pathlib.Path(r'c:\ACP\Query1.txt')
        #Everything else hinges on the timer function, which is below.
        def Create_After(path, where, first_modified, miliseconds=1000): #The following must be done in a function, since it waits for files to be checked on. 
            #If the path given is equal to the time given, check the path again, after the time provided, default is 1 second.
            if path.stat().st_mtime == first_modified:
                if Create.RepCount < 300:
                    app.after(miliseconds,lambda : Create_After(path, where, first_modified))
                    Create.Loading_Indicator.config(text='Checking for duplicates and Creating User...')
                    app.after(333,lambda:Create.Loading_Indicator.config(text='Checking for duplicates and Creating User.'))
                    app.after(666,lambda:Create.Loading_Indicator.config(text='Checking for duplicates and Creating User..'))
                    Create.RepCount=Create.RepCount+1
                    print(Create.RepCount)
                else:
                    if send_mail.get() == True:
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName+'',r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit Send-MailMessage -From usermgmt-at@monroecountyymca.org -To ithelpdesk@MonroeCountyYMCA.org -Subject '"+Create.Fullname+" Account Creation Failed or was Cancelled' -Body 'An account was attempted to be made, but failed. Account name: "+Create.Fullname+"' -SmtpServer 10.10.100.1;exit"])
                        else: subprocess.call(['runas', '/savecred','/user:'+loginName+'',r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit Send-MailMessage -From usermgmt-at@monroecountyymca.org -To ithelpdesk@MonroeCountyYMCA.org -Subject '"+Create.Fullname+" Account Creation Failed or was Cancelled' -Body 'An account was attempted to be made, but failed. Account name: "+Create.Fullname+"' -SmtpServer 10.10.100.1;"])
                    showerror(title='Account Creation Failed', message='The account creation has failed or was cancelled. ')
                    Create.RepCount = 0
                    Create.Loading_Indicator.config(text='')
                    return 1
                    
                    
            #This code is structured in a strange format. Every time a file is checked, it is passed a unique code as the 'where' variable. When the file is changed, this code is used to detect what to do next. 
            else:
                if where == 'does_it_exist':
                    #After the duplicate check finishes
                    Query1 = open(r'c:\ACP\Query1.txt', "r")
                    Query1_Splitlines = Query1.read().splitlines()
                    try:
                        #Try to see if the account is there, and if it is enabled or disabled respectively.
                        if Query1_Splitlines[3].split()[2] == 'True':
                            Modify.Status_Button.grid(row=3, column = 1, sticky='W')
                            Create.Create_User_Button.config(command = lambda : createAccount())
                            Create.Duplicate_Warning=showwarning(title='Duplicate User', message="This user is a duplicate. Consider changing the user's name, or switching to the Modify/Disable page.")
                            Create.Finished==True
                            
                        if Query1_Splitlines[3].split()[2] == 'False':
                            Modify.Status_Button.grid(row=3, column = 1, sticky='W')
                            Create.Create_User_Button.config(command = lambda : createAccount())
                            Create.Duplicate_Warning=showwarning(title='Duplicate User', message="This user is a duplicate. Consider changing the user's name, or switching to the Modify/Disable page.")
                            Create.Finished==True
                        Create.Loading_Indicator.config(text='')
                    except IndexError:
                        #If this exception is raised, then the account does not exist. This is the result we want usually.
                        statVar.set(0)
                        Create.Create_User_Button.config(command = lambda : createAccount())

                        #region Once the account is determined to be new, generate the commands.
                        Create.Password = ''.join(random.choice(string.ascii_letters + string.digits) for i in range(12)) + "!"
                        
                        if Create.MidName != '':
                            Create.Command_1 ="New-RemoteMailbox -Name \'"+Create.Fullname+"\' -FirstName "+Create.FirstName+" -LastName "+Create.LastName+" -Initials "+Create.MidName+" -UserPrincipalName "+Create.Email+" -Password (ConvertTo-SecureString -String \'"+Create.Password +"\' -AsPlainText -Force) -ResetPasswordOnNextLogon $true -OnPremisesOrganizationalUnit \"ymca.local/MC-YMCA/Users/Staff\""
                        else:
                            Create.Command_1 ="New-RemoteMailbox -Name \'"+Create.Fullname+"\' -FirstName "+Create.FirstName+" -LastName "+Create.LastName+" -UserPrincipalName "+Create.Email+" -Password (ConvertTo-SecureString -String \'"+Create.Password +"\' -AsPlainText -Force) -ResetPasswordOnNextLogon $true -OnPremisesOrganizationalUnit \"ymca.local/MC-YMCA/Users/Staff\""
                        #Create.Description = "No Description"
                        Create.Groups_Queue=Create.Groups_Queue+Create.Active_Master
                        for group in Create.Active_Master:
                            #if the group is a preset
                            if '(Preset)' in group:
                                #Find the group in the list of presets, take the index number of where that group locates, and then check the same index number in the list of groups for each preset.
                                for item in Create.Presets.Corresponding[Create.Presets.Master.index(group)]:
                                    #Exclude blank presets to prevent unnecessary errors
                                    if item != '':
                                        #Add the groups to the master list
                                        Create.Groups_Queue.append(item)
                                Create.Groups_Queue.remove(group)     
                        
                        Create.All_Groups='\',\''.join(Create.Groups_Queue)
                        if len('"'+Create.All_Groups+'"|Add-ADGroupMember -Members "'+Create.Username+'"') > 600:
                            part1=Create.Groups_Queue[:len(Create.Groups_Queue)//2]
                            part2=Create.Groups_Queue[len(Create.Groups_Queue)//2:]
                            groupsPart1='\',\''.join(part1)
                            groupsPart2='\',\''.join(part2)
                            Create.Command_2 ='\''+groupsPart1+'\' |Add-ADGroupMember -Members "'+Create.Username+'"'
                            Create.Command_2b='\''+groupsPart2+'\' |Add-ADGroupMember -Members "'+Create.Username+'"'
                        else:
                            Create.Command_2 ='\''+Create.All_Groups+'\'|Add-ADGroupMember -Members "'+Create.Username+'"'

                        Create.Command_3='Set-ADUser -Identity "'+Create.Username+' " -Description \''+Create.Description+'\''

                        
                        Create.Command_4 = "$credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential;Connect-AzureAD -Credential $credential;Set-MsolUser -UserPrincipalName \""+Create.Email+"\" -UsageLocation US;Set-MsolUserLicense -UserPrincipalName \""+Create.Email+"\" -AddLicenses \"monroecountyymca:STANDARDWOFFPACK\""

                        #endregion create the commands

                        #Create the account in AD. Give name, email, password.
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_1+'| Out-File -FilePath c:\\ACP\\createLog.txt -Encoding ASCII;exit'])
                        else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_1+'| Out-File -FilePath c:\\ACP\\createLog.txt -Encoding ASCII;'])
                        Create_After(pathlib.Path(r'c:\ACP\createLog.txt'), 'create_account', pathlib.Path(r'c:\ACP\createLog.txt').stat().st_mtime)
                if where == 'create_account':
                    #The initial creation trips the update twice, so this catches that issue.
                    Create_After(pathlib.Path(r'c:\ACP\createLog.txt'), 'dummy', pathlib.Path(r'c:\ACP\createLog.txt').stat().st_mtime)

                if where == 'dummy':
                    #Assign groups to the account. If the total groups command is too long, then it needs to be split up. Currently, all the split does is assign groups in 1 command, and description in another. 
                    #This if statement checks how long the command will be. If it is too long, the command is split into two parts. 
                    #If the command is less than 600, we run command 2.
                    #If command is over 600, we run command 2, then command 2_b. If we only run 1 command, "where='groups_part1_added" is skipped, and we go straight to "'groups_final_added'".
                    
                    #These two parts look the same, but one runs straight to the description change, while the other stops at a second "group add" command first.
                    # The assigning of command_2 and command_b is done when the create page is previewed, and is run on the same condition of over 600 length, so as to be consistant.
                     
                    if len(str(Create.Groups_Queue)+' |Add-ADGroupMember -Members "'+Create.Username+'"') > 600:
                        #p rint('The groups command is too long, splitting into two...')
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_2+'; Add-Content -Path c:\\ACP\\createLog.txt -Value \'Added Groups1\';exit'])
                        else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_2+'; Add-Content -Path c:\\ACP\\createLog.txt -Value \'Added Groups1\';'])
                        Create_After(pathlib.Path(r'c:\ACP\createLog.txt'), 'groups_part1_added', pathlib.Path(r'c:\ACP\createLog.txt').stat().st_mtime)

                    else:
                        #p rint('Groups command is short, running single.')                        
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_2+'; Add-Content -Path c:\\ACP\\createLog.txt -Value \'Added Groups\';exit'])
                        else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_2+'; Add-Content -Path c:\\ACP\\createLog.txt -Value \'Added Groups\';'])
                        Create_After(pathlib.Path(r'c:\ACP\createLog.txt'), 'groups_final_added', pathlib.Path(r'c:\ACP\createLog.txt').stat().st_mtime)


                if where == 'groups_part1_added':
                    #Add the Description
                    if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_2b+'; Add-Content -Path c:\\ACP\\createLog.txt -Value \'Added Groups2\';exit']) 
                    else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_2b+'; Add-Content -Path c:\\ACP\\createLog.txt -Value \'Added Groups2\';']) 
                    Create_After(pathlib.Path(r'c:\ACP\createLog.txt'), 'groups_final_added', pathlib.Path(r'c:\ACP\createLog.txt').stat().st_mtime)

                if where == 'groups_final_added':
                    if autoclose.get() == True: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_3+'; Add-Content -Path c:\\ACP\\createLog.txt -Value \'Added Description\';exit'])
                    else: subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_3+'; Add-Content -Path c:\\ACP\\createLog.txt -Value \'Added Description\';'])
                    Create_After(pathlib.Path(r'c:\ACP\createLog.txt'), 'description_added', pathlib.Path(r'c:\ACP\createLog.txt').stat().st_mtime)

                if where == 'description_added':
                    #Start the dirsync to the cloud. Returns when we get confirmation the sync has started. 
                    if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit Invoke-Command -ComputerName dirsync -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test";exit'])
                    else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit Invoke-Command -ComputerName dirsync -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test";'])
                    Create_After(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'dirsync_started', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime)
                
                if where == 'dirsync_started':
                    #Pull information as to the last time the sync occurred, which is equivalent to the first_modified time on the python side.
                    if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test2";exit'])
                    else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test2";'])
                    Create_After(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'original_sync_time_query', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime)
                
                if where == 'original_sync_time_query':
                    #Once we get the time to hold as the original, we can begin checking the time again every ten seconds, to check for a change.
                    ResultsFile = open(r'c:\ACP\DirsyncUpdateLog2.txt', "r")
                    Lines = ResultsFile.read().splitlines()
                    Create.OriginalSyncTime=Lines[20]
                    if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";exit'])
                    else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";'])
                    Create_After(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'new_sync_time_query', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime,10000)
                
                if where == 'new_sync_time_query':
                    #Check if the time is different or not. If it is not, the program runs again, returning to this block. If it is different, it proceeds.
                    ResultsFile = open(r'c:\ACP\DirsyncUpdateLog2.txt', "r")
                    Lines = ResultsFile.read().splitlines()
                    if Lines[20]==Create.OriginalSyncTime:
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";exit'])
                        else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";'])
                        Create_After(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'new_sync_time_query', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime,10000)
                    else:
                        #If time has changed, wait ten seconds for dirsync to finalize, and then assign license.
                        Create_After(pathlib.Path(r'c:\ACP\createLog.txt'), 'license_user', pathlib.Path(r'c:\ACP\createLog.txt').stat().st_mtime)
                        if Create.License_User_Var.get()==True:
                            if autoclose.get() == True: app.after(35000,lambda : subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_4+'; Add-Content -Path c:\\ACP\\createLog.txt -Value "test5";exit']))
                            else: app.after(35000,lambda : subprocess.call(['runas', '/savecred', '/user:'+loginName+'','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit  \". \'C:\\Program Files\\Microsoft\\Exchange Server\\V15\\bin\\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";'+Create.Command_4+'; Add-Content -Path c:\\ACP\\createLog.txt -Value "test5";']))
                        else: 
                            temp=open('c:\\ACP\\createLog.txt', 'a')
                            temp.write('no license')

                if where == 'license_user':
                    Create_After(pathlib.Path(r'c:\ACP\createLog.txt'), 'MFA Added', pathlib.Path(r'c:\ACP\createLog.txt').stat().st_mtime)
                    #if Create.Enable_MFA_Var.get()==True:
                    #    subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; $mf= New-Object -TypeName Microsoft.Online.Administration.StrongAuthenticationRequirement; $mf.RelyingParty = \"*\"; $mfa = @($mf); Set-MsolUser -UserPrincipalName \"'+Create.Email+'\" -StrongAuthenticationRequirements $mfa; Add-Content -Path c:\\ACP\\createLog.txt -Value "Enabled_MFA";'])
                    #else:
                    temp=open('c:\\ACP\\createLog.txt', 'a')
                    temp.write('no MFA')        

                if where == 'MFA Added':
                    #p rint('The user is created, grouped, synced, and licensed.')
                    Create.Loading_Indicator.config(text='')
                    try: create_final=open('c:\ACP\create_final.txt','x') 
                    except: pass
                    Create_final=open('c:\ACP\create_final.txt','w') 
                    #p rint(Create.Description)
                    if Create.Description_Lock_Var.get()==True:
                        Create.Description_Entry.config(state="normal")
                        Create_final.writelines(['Name: ',Create.Fullname, '\n','Username: ', Create.Username, '\n',"Email: ", Create.Email, '\n','Password: ', Create.Password, '\n','Description: ', Create.Description_Entry.get(), "\nSupervisor: ", Create.Supervisor.get()])
                        Create.Description_Entry.config(state="disabled")
                    else: Create_final.writelines(['Name: ',Create.Fullname, '\n','Username: ', Create.Username, '\n',"Email: ", Create.Email, '\n','Password: ', Create.Password, '\n','Description: ', Create.Description_Entry.get(), "\nSupervisor: ", Create.Supervisor.get()])
                                        
                    if send_mail.get()==True:
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName+'',r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit Send-MailMessage -From usermgmt-at@monroecountyymca.org -To ithelpdesk@MonroeCountyYMCA.org -Subject '"+Create.Fullname+" Account Created' -Body (Get-Content -Path C:\ACP\create_final.txt | Out-String) -SmtpServer 10.10.100.1;exit"])
                        else: subprocess.call(['runas', '/savecred','/user:'+loginName+'',r"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit Send-MailMessage -From usermgmt-at@monroecountyymca.org -To ithelpdesk@MonroeCountyYMCA.org -Subject '"+Create.Fullname+" Account Created' -Body (Get-Content -Path C:\ACP\create_final.txt | Out-String) -SmtpServer 10.10.100.1;"])
                    
                    webbrowser.open(r"c:\ACP\create_final.txt")
        Create_After(create_path, 'does_it_exist', create_path.stat().st_mtime)  

def Modify_Description_Update():
    if Modify.Auto_Mode.get()==1:
        Modify.Description="No Description"
        for group in Modify.Active_Master:
            #if the group is a preset
            if '(Preset)' in group:
                if Modify.Description == "No Description":
                    Modify.Description = ''
                Modify.Description=Modify.Description+Modify.Presets.Descriptions[Modify.Presets.Master.index(group)]+", "
        Modify.Description_Entry.config(state="normal")
        Modify.Description_Entry.delete(0,END)
        Modify.Description_Entry.insert(0,Modify.Description)
        Modify.Description_Entry.config(state="disabled")
    else:
        Modify.Description=Modify.Description_Entry.get()

def Modify_Desc_Lock():
    if Modify.Auto_Mode.get()==1:
        Modify_Description_Update()
        Modify.Description_Entry.config(state="disabled")
    if Modify.Auto_Mode.get()==0:
        Modify.Description_Entry.config(state="normal")

def Create_Description_Update():
    if Create.Description_Lock_Var.get()==1:
        Create.Description="No Description"
        for group in Create.Active_Master:
            #if the group is a preset
            if '(Preset)' in group:
                if Create.Description == "No Description":
                    Create.Description = ''
                Create.Description=Create.Description+Create.Presets.Descriptions[Create.Presets.Master.index(group)]+", "
                #p rint(Create.Description)
        Create.Description_Entry.config(state="normal")
        Create.Description_Entry.delete(0,END)
        Create.Description_Entry.insert(0,Create.Description)
        Create.Description_Entry.config(state="disabled")
    else:
        Create.Description=Create.Description_Entry.get()

def Create_Desc_Lock():
    if Create.Description_Lock_Var.get()==1:
        Create_Description_Update()
        Create.Description_Entry.config(state="disabled")
    if Create.Description_Lock_Var.get()==0:
        Create.Description_Entry.config(state="normal")

#endregion Functions

LARGEFONT =("Verdana", 35)  
class tkinterApp(tk.Tk): #Driver Code

    # __init__ function for class tkinterApp
    def __init__(self, *args, **kwargs):
    

        # __init__ function for class Tk
        tk.Tk.__init__(self, *args, **kwargs)
        
        # creating a container
        container = tk.Frame(self) 
        container.pack(side = "top", fill = "both", expand = True)
  
        container.grid_rowconfigure(0, weight = 1)
        container.grid_columnconfigure(0, weight = 1)

        # initializing frames to an empty array
        self.frames = {} 

        self.title('YMCA Account Creation Program')
        self.geometry('900x600')
        
        # iterating through a tuple consisting
        # of the different page layouts
        for F in (StartPage, Creation, Modification, Settings_Page):
  
            frame = F(container, self)
  
            # initializing frame of that object from startpage, page1, page2 respectively with for loop
            self.frames[F] = frame
            
            frame.grid(row = 0, column = 0, sticky ="nsew")
  
        self.show_frame(StartPage)
  
    # to display the current frame passed as
    # parameter
    def show_frame(self, cont):
        frame = self.frames[cont]
        frame.tkraise()

class StartPage(tk.Frame): #Initial Page Frame
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = '#375280'
        global gotoCreate, gotoModify, gotoSettings_Page, buttonCreds

        label1=tk.Label(self,  bg='#375280',text='Thank you for using the YMCA Account Creation Program. \n For usage info, click here. To begin, select an option below.',fg='ghost white',font=("Verdana", 12),wraplength=0)
        label1.grid(row = 0, column=2, columnspan = 100, padx = 5, pady = 5)

        gotoCreate = tk.Button(self, text ="Create\nan Account",command = lambda : controller.show_frame(Creation), width=10, height=5, font=("Verdana", 9))
        gotoCreate.grid(row = 1, column = 0, padx = 5, pady = 5, sticky="NSEW")
        gotoModify = tk.Button(self, text ="Modify/Disable\nan Account",command = lambda : controller.show_frame(Modification), width=10, height=5, font=("Verdana", 9))
        gotoModify.grid(row = 2, column = 0, padx = 5, pady = 5, sticky="NSEW")
        gotoSettings_Page = tk.Button(self, text ="Settings", command = lambda : controller.show_frame(Settings_Page), width=10, height=5, font=("Verdana", 9))
        gotoSettings_Page.grid(row = 3, column = 0, padx = 5, pady = 5, sticky="NSEW")
        buttonCreds = tk.Button(self, text ="Update your\nCredentials",command = lambda : CredSetup(), width=10, height=5, font=("Verdana", 9))
        buttonCreds.grid(row = 4, column = 0, padx = 5, pady = 5, sticky="NSEW")

class Creation(tk.Frame): #Page frame for creating new user accounts.
    def __init__(self, parent, controller):    
        tk.Frame.__init__(self, parent)
        self['bg'] = '#375280'
        
        #region Preview Logic
        dataLabel1a=Label(self)
        dataLabel2a=Label(self)
        dataLabel3a=Label(self)
        dataLabel1b=Label(self)
        dataLabel2b=Label(self)
        dataLabel3b=Label(self)
        Create.Create_User_Button=tk.Button(self, text='Create User', command=lambda : createAccount())
        
        Create.Advanced_Mode=IntVar()
        Create.Advanced_Mode.set(False)

        Create.Inactive_Listbox=Listbox(self,width=50,height=10)
        Create.Inactive_Listbox.bind('<Double-Button-1>', Create.Inactive_to_Active)

        Create.Inactive_Search = Entry(self,width=50)
        Create.Inactive_Search.bind('<KeyRelease>', Create.Inactive_Update)

        Create.Active_Listbox=Listbox(self,width=50,height=10)
        Create.Active_Listbox.bind('<Double-Button-1>', Create.Active_to_Inactive)

        Create.Active_Search = Entry(self,width=50)
        Create.Active_Search.bind('<KeyRelease>', Create.Active_Update)

        Create.Description_Entry = tk.Entry(self, width=38, font=("Verdana", 10),fg = 'black')

        Create.Supervisor = StringVar()
        entries = [
            "No Supervisor Selected",
            "No Supervisor Selected",
            "Chelcey Bostic",
            "Chris Stone",
            "David Stowe",
            "Emily Abbott",
            "Grace Haskett",
            "Grace Thomas",
            "Haley Cook",
            "Herb Susenbach",
            "Jake Steinmetz",
            "Jason Winkle",
            "Jeff Albertson",
            "Jodi Baker",
            "Kevin Thompson",
            "Kevin Vail",
            "Kristin Good",
            "Margie Kobow",
            "Matt Osgood",
            "Megan Irwin",
            "Scott Warrick",
            "Shannon Kane",
            "Other",
        ]

        Create.Supervisor_Dropdown = OptionMenu(self, Create.Supervisor, *entries)
        
        def Preview(event=''):
            Create.Fullname=''
            Create.Username=''
            if  Create.Name_Entry_Box.get()=="Enter the user's Full Name..." or Create.Name_Entry_Box.get()==" " or Create.Name_Entry_Box.get()=='':
                #p rint('Enter something into the box.')
                dataLabel1a.grid_forget()
                dataLabel2a.grid_forget()
                dataLabel3a.grid_forget()
                Create.Create_User_Button.grid_forget()
                return
            else:
                Create.Fullname=Create.Name_Entry_Box.get()
                Create.Entry_Split_to_List=Create.Name_Entry_Box.get().split()

                Create.FirstName=''
                Create.LastName = ''
                Create.MidName = ''
                if len(Create.Entry_Split_to_List) == 2:
                    Create.FirstName=Create.Entry_Split_to_List[0]
                    Create.LastName = Create.Entry_Split_to_List[1]
                    Create.MidName = ''
                    Create.Username=Create.Entry_Split_to_List[0].lower()[0:1]+Create.Entry_Split_to_List[1].lower()
                elif len(Create.Entry_Split_to_List) > 2:
                    Create.FirstName=Create.Entry_Split_to_List[0]
                    Create.LastName = Create.Entry_Split_to_List[2]
                    Create.MidName = Create.Entry_Split_to_List[1]
                    Create.Username=Create.Entry_Split_to_List[0].lower()[0:1]+Create.Entry_Split_to_List[1].lower()+Create.Entry_Split_to_List[2].lower()
                elif len(Create.Entry_Split_to_List) == 1:
                    Create.FirstName=Create.Entry_Split_to_List[0]     
                    dataLabel1a.grid_forget()
                    dataLabel2a.grid_forget()
                    dataLabel3a.grid_forget()
                    Create.Create_User_Button.grid_forget()
                    return
            

            
            Create.Description='No Description'
            Create.Description_Entry.grid(row=8, column=1, sticky='W')
            Create.Description_Entry.config(text=Create.Description)

            Create.Description_Label=tk.Label(self, text='Description: ', background='#375280',foreground='ghost white',font=("Verdana", 10))
            Create.Description_Label.grid(row=8, column=0, sticky="W")
            
            Create.Description_Lock.grid(row=8, column=1, sticky='E')
            Create.Desc_Lock_Label=tk.Label(self, text='Auto:', background='#375280',foreground='ghost white',font=("Verdana", 10))
            Create.Desc_Lock_Label.grid(row=7, column=1, sticky="E")

            Create.License_User_Button.grid(row=100, column=0)
            Create.License_User_Label.grid(row=99, column=0)
            Create.Groups_Label=tk.Label(self, text='Available Groups:                                 Current Groups: ', background='#375280',foreground='ghost white',font=("Verdana", 10))
            Create.Groups_Label.grid(row=9, column=0, sticky="W", columnspan=100)

            dataLabel1a.config(text=Create.Fullname, background='#375280',foreground='ghost white',font=("Verdana", 10))
            dataLabel1a.grid(row=5,column=1, sticky='W',pady=5,)

            dataLabel2a.config(text=Create.Username, background='#375280',foreground='ghost white',font=("Verdana", 10))
            dataLabel2a.grid(row=6,column=1, sticky='W',pady=5,)
            
            Create.Supervisor_Dropdown.grid(row=12, column = 2, sticky='W', columnspan=100)


            if Create.Username != '':
                Create.Email=Create.Username+'@monroecountyymca.org'
            else: Create.Email=''
            dataLabel3a.config(text=Create.Email, background='#375280',foreground='ghost white',font=("Verdana", 10))
            dataLabel3a.grid(row=7,column=1, sticky='W',pady=5,)

            dataLabel1b.config(text='Name:', background='#375280',foreground='ghost white',font=("Verdana", 10))
            dataLabel1b.grid(row=5,column=0, sticky='W',pady=5,)

            dataLabel2b.config(text='Username: ', background='#375280',foreground='ghost white',font=("Verdana", 10))
            dataLabel2b.grid(row=6,column=0, sticky='W',pady=5,)

            dataLabel3b.config(text='Email: ', background='#375280',foreground='ghost white',font=("Verdana", 10))
            dataLabel3b.grid(row=7,column=0, sticky='W',pady=5,)
            
            Create.Create_User_Button.grid(row=12, column=1, padx = 0, pady = 10, sticky="W")

            Create.Inactive_Search.grid(row=10, column=0, columnspan=100, sticky='W', padx = 5)
            Create.Inactive_Listbox.grid(row=11, column=0, columnspan=100, sticky='W', padx = 5)
            Create.Active_Search.grid(row=10, column=1, columnspan=100, sticky='E', padx = 10)
            Create.Active_Listbox.grid(row=11, column=1, columnspan=100, sticky='E', padx = 10)

            Create.AdvancedLabel.grid(row=9, column=2, sticky='E',padx=1)
            Create.AdvancedToggle.grid(row=9, column=3, sticky='W', padx=1)


            

        #endregion Preview Logic
        Create.Spacer_Label = tk.Label(self, text="", bg='#375280')
        Create.Spacer_Label.grid(row=4, column=0)
        global backButton2
        backButton2 = tk.Button(self, text ="< Back to Home",command = lambda : controller.show_frame(StartPage))
        backButton2.grid(row = 0, column = 0, padx = 5, pady = 5, sticky="NW")

        # Label at top with directions
        Create.Directions_Label = tk.Label(self, text='Account Creation: \n Enter the user\'s name, then press "Preview".',fg='ghost white',font=("Verdana", 12),wraplength=0, bg='#375280')
        Create.Directions_Label.grid(row = 0, column=1, columnspan = 100, padx = 5, pady = 5)

        Create.Loading_Indicator = tk.Label(self,fg='ghost white',font=("Courier", 10),wraplength=500, bg='#375280')
        Create.Loading_Indicator.grid(row=2, column=1, columnspan=100)

        #region Text Boxes w/ Functions
        def Create_Entry_On_Click(event):
            """function that gets called whenever entry is clicked"""
            if Create.Name_Entry_Box.get() == "Enter the user's Full Name...":
                Create.Name_Entry_Box.delete(0, "end") # delete all the text in the entry
                Create.Name_Entry_Box.insert(0, '') #Insert blank for user input
                Create.Name_Entry_Box.config(fg = 'black')
        def Create_Entry_On_FocusOut(event):
            if Create.Name_Entry_Box.get() == '':
                Create.Name_Entry_Box.insert(0, "Enter the user's Full Name...")
                Create.Name_Entry_Box.config(fg = 'grey')
                
        Create.Name_Entry_Box=tk.Entry(self, width=40, font=("Verdana", 10),fg = 'grey')#Create menu's entry box.
        Create.Name_Entry_Box.grid(row = 1, column=1, padx = 5, pady = 2)#Grid it
        
        #Ghost text
        Create.Name_Entry_Box.insert(0, "Enter the user's Full Name...")
        Create.Name_Entry_Box.bind('<FocusIn>', Create_Entry_On_Click)
        Create.Name_Entry_Box.bind('<FocusOut>', Create_Entry_On_FocusOut)
        Create.Name_Entry_Box.bind('<Return>', Preview)


        #endregion Text Boxes w/ Functions
        global preview
        preview=tk.Button(self, text='Preview', command=lambda : Preview(event=''))
        preview.grid(row=1, column=2, padx = 0, pady = 0, sticky="W")
        
        Create.AdvancedToggle=tk.Checkbutton(self,variable=Create.Advanced_Mode, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280', command=lambda : Create.Inactive_Update(event=None,value=Create.Inactive_Search.get()))
        Create.AdvancedLabel=tk.Label(self, text='Advanced Settings',background='#375280',foreground='ghost white',font=("Verdana", 8))
        Create.Description_Lock_Var=IntVar(value=True)
        Create.Description_Lock=tk.Checkbutton(self,variable=Create.Description_Lock_Var, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280', command= lambda : Create_Desc_Lock())
        Create_Desc_Lock()
        Create.License_User_Var=IntVar(value=1)
        Create.License_User_Button=tk.Checkbutton(self,variable=Create.License_User_Var, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280')
        Create.License_User_Label=tk.Label(self, text='Assign License:',background='#375280',foreground='ghost white',font=("Verdana", 8))
        Create.Enable_MFA_Var=IntVar(value=1)
        Create.Enable_MFA_Button=tk.Checkbutton(self,variable=Create.Enable_MFA_Var, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280')
        Create.Enable_MFA_Label=tk.Label(self, text='Enable MFA:',background='#375280',foreground='ghost white',font=("Verdana", 8))


        Create.Inactive_Update()
        Create.Active_Update()

class Modification(tk.Frame): #Page Frame for the Modify Menu
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = '#375280'
        #Button to return to homepage.
        global backButton3
        backButton3 = tk.Button(self, text ="< Back to Home",command = lambda : controller.show_frame(StartPage))
        backButton3.grid(row = 0, column = 0, padx = 5, pady = 5, sticky="NW")

        # Label at top with directions
        Directions = tk.Label(self, bg='#375280', text="Modify Accounts: \nTo find the user you wish to modify, enter either the username or exact full name, including an initial if the account was created with one.",fg='ghost white',font=("Verdana", 12),wraplength=500)
        Directions.grid(row = 0, column=1, columnspan = 100, padx = 5, pady = 5)
        Modify.Auto_Mode=IntVar()
        Modify.Auto_Mode.set(False)
        global statVar
        Modify.Description_Label=tk.Label(self, text='Description: ', background='#375280',foreground='ghost white',font=("Verdana", 10),padx = 5,)
        Modify.Description_Entry = tk.Entry(self, width=38, font=("Verdana", 10),fg = 'black')

        Modify.Advanced_Mode=IntVar()
        Modify.Advanced_Mode.set(False)

        Modify.Inactive_Listbox=Listbox(self,width=50,height=10)
        Modify.Inactive_Listbox.bind('<Double-Button-1>', Modify.Inactive_to_Active)

        Modify.Inactive_Search = Entry(self,width=50)
        Modify.Inactive_Search.bind('<KeyRelease>', Modify.Inactive_Update)

        Modify.Active_Listbox=Listbox(self,width=50,height=10)
        Modify.Active_Listbox.bind('<Double-Button-1>', Modify.Active_to_Inactive)

        Modify.Active_Search = Entry(self,width=50)
        Modify.Active_Search.bind('<KeyRelease>', Modify.Active_Update)
        
        Modify.AdvancedToggle=tk.Checkbutton(self,variable=Modify.Advanced_Mode, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280', command=lambda : Modify.Inactive_Update(event=None,value=Modify.Inactive_Search.get()))
        Modify.AdvancedLabel=tk.Label(self, text='Advanced Settings',background='#375280',foreground='ghost white',font=("Verdana", 8))

        Modify.Inactive_Update()
        Modify.Active_Update()

        Modify.Auto_Toggle=tk.Checkbutton(self,variable=Modify.Auto_Mode, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280', command=lambda : Modify_Desc_Lock())
        Modify.Auto_Label=tk.Label(self, text='Auto:',background='#375280',foreground='ghost white',font=("Verdana", 8))
        
        #Push Changes Button
        Modify.Push_Changes_Button=tk.Button(self, text='Review/Save Changes', command=lambda:pushChanges())

        #Buttons to trigger searches. Should eventually be combined into one searchbox that detects if a username or a full-name is entered.
        Modify.Search_User_Button = tk.Button(self, text ="Search for this User",command = lambda : queryAccount(),width=16)
        Modify.Search_User_Button.grid(row = 1, column = 2, padx = 5, pady = 1, sticky="NW")
        #initialize the loading label for activation later
        Modify.Loading_Indicator=tk.Label(self ,fg='ghost white',font=("Courier", 10),wraplength=500, bg='#375280')
        
        #region Creates text boxes and manages ghost text.
        def Modify_Entry_On_Click(event=''):
            """function that gets called whenever entry is clicked"""
            if Modify.Name_Entry_Box.get() == "Enter the user's Full Name or their Username...":
                Modify.Name_Entry_Box.delete(0, "end") # delete all the text in the entry
                Modify.Name_Entry_Box.insert(0, '') #Insert blank for user input
                Modify.Name_Entry_Box.config(fg = 'black')
        def Modify_Entry_On_FocusOut(event=''):
            if Modify.Name_Entry_Box.get() == '':
                Modify.Name_Entry_Box.insert(0, "Enter the user's Full Name or their Username...")
                Modify.Name_Entry_Box.config(fg = 'grey')
        Modify.Name_Entry_Box=AutocompleteEntry(self, width=40, font=("Verdana", 10))

        Modify.Name_Entry_Box.grid(row = 1, column=1, padx = 5, pady = 2)
        Modify.Name_Entry_Box.config(fg = 'grey')
        Modify.Name_Entry_Box.insert(0, "Enter the user's Full Name or their Username...")
        Modify.Name_Entry_Box.bind('<FocusIn>', Modify_Entry_On_Click)
        Modify.Name_Entry_Box.bind('<FocusOut>', Modify_Entry_On_FocusOut)
        Modify.Name_Entry_Box.bind('<Return>', queryAccount)
        #endregion Text Boxes w/ Functions
        statVar=IntVar(value=0)
        #Labels to display the data pulled from PowerShell.
        #Blank label to offset the rest
        Modify.Filler_Label=tk.Label(self ,fg='ghost white',font=("Verdana", 10), bg='#375280')
        Modify.Filler_Label.grid(row = 2, column=0, padx = 5, pady = 2,columnspan=100,sticky='W')
        
        #Tell the user that the following data is the user status.
        Modify.Status_Identifier_Label=tk.Label(self ,fg='ghost white',font=("Verdana", 10), bg='#375280')
        Modify.Status_Identifier_Label.grid(row = 3, column=0, padx = 5, pady = 2,columnspan=100,sticky='W')

        #Displays a check if the user is enabled, empty if disabled or nonexistent. 
        Modify.Status_Button=tk.Checkbutton(self,variable=statVar, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280', command=lambda : Lock_Groups())

        #Displays the queried user's name.
        Modify.Name_Display_Label=tk.Label(self ,fg='ghost white',font=("Verdana", 10), bg='#375280')
        Modify.Name_Display_Label.grid(row = 4, column=1, padx = 5, pady = 2,columnspan=100,sticky='W')
        
        #Tell the user that the following data is the user's name.
        Modify.Name_Identifier_Label=tk.Label(self ,fg='ghost white',font=("Verdana", 10), bg='#375280')
        Modify.Name_Identifier_Label.grid(row = 4, column=0, padx = 5, pady = 2,columnspan=100,sticky='W')

        #Displays the queried user's username.
        Modify.Username_Display_Label=tk.Label(self ,fg='ghost white',font=("Verdana", 10), bg='#375280')
        Modify.Username_Display_Label.grid(row = 5, column=1, padx = 5, pady = 2,columnspan=100,sticky='W')
        
        #Tell the user that the following data is the user's username.
        Modify.Username_Identifier_Label=tk.Label(self ,fg='ghost white',font=("Verdana", 10), bg='#375280')
        Modify.Username_Identifier_Label.grid(row = 5, column=0, padx = 5, pady = 2,columnspan=100,sticky='W')

        #Displays the queried user's email.
        Modify.Email_Display_Label=tk.Label(self ,fg='ghost white',font=("Verdana", 10), bg='#375280')
        Modify.Email_Display_Label.grid(row = 6, column=1, padx = 5, pady = 2,columnspan=100,sticky='W')
        
        #Tell the user that the following data is the user's email.
        Modify.Email_Identifier_Label=tk.Label(self ,fg='ghost white',font=("Verdana", 10), bg='#375280')
        Modify.Email_Identifier_Label.grid(row = 6, column=0, padx = 5, pady = 2,columnspan=100,sticky='W')
       
        #Tell the user that the following data is the user's groups.
        Modify.Groups_Identifier_Label=tk.Label(self ,fg='ghost white',font=("Verdana", 10), bg='#375280')
        Modify.Groups_Identifier_Label.grid(row = 9, column=0, padx = 5, pady = 2,columnspan=100,sticky='W')

class Settings_Page(tk.Frame): #Page Frame for Settings
    def __init__(self, parent, controller):
        tk.Frame.__init__(self, parent)
        self['bg'] = '#375280'
        if not os.path.exists('c:\\ACP'):
            os.makedirs('c:\\ACP')
        #Button to return to homepage.
        global backButton4
        backButton4 = tk.Button(self, text ="< Back to Home",command = lambda : controller.show_frame(StartPage))
        backButton4.grid(row = 0, column = 0, padx = 5, pady = 5, sticky="NW")

        # Label at top with directions
        Directions = tk.Label(self, bg='#375280', text="Settings: \nSettings changed here will apply over the entire program, and will save between sessions.",fg='ghost white',font=("Verdana", 12),wraplength=500)
        Directions.grid(row = 0, column=1, columnspan = 100, padx = 5, pady = 5,sticky="NW")             

        global autoCompToggleVar, show_hovertips, autoclose, send_mail
        Default_Settings=['[Account Creation Program Settings\n','autoCompleteEnabled: True\n','show_hovertips: True\n','autoclose: True\n','send_mail: True\n']
        
        
        
        autoCompToggleLabel=tk.Label(self, fg='ghost white',font=("Verdana", 10), bg='#375280', text='Enable Text-Box Autocomplete: ')
        autoCompToggleLabel.grid(row=1,column=0, sticky="S")

        #make sure the file exists.
        try: Settings_File=open(r'c:\ACP\Settings.txt','x'); Settings_File.close()
        except: pass

        #Read the settings to a list
        Settings_File=open(r'c:\ACP\Settings.txt','r+')
        if Settings_File.readlines()==[]: Settings_File.writelines(Default_Settings)
        Settings_File=open(r'c:\ACP\Settings.txt','r+')
        Settings=Settings_File.readlines()
        #Assign the settings from the list
        autoCompToggleVar=BooleanVar(value=Settings[1].strip().split()[1])
        show_hovertips=BooleanVar(value=Settings[2].strip().split()[1])
        autoclose=BooleanVar(value=Settings[3].strip().split()[1])
        send_mail=BooleanVar(value=Settings[4].strip().split()[1])

        def settings_save():
            Settings[1]=('autoCompleteEnabled: '+str(autoCompToggleVar.get())+'\n')
            Settings[2]=('show_hovertips: '+str(show_hovertips.get())+'\n')
            Settings[3]=('autoclose: '+str(autoclose.get())+'\n')
            Settings[4]=('send_mail: '+str(send_mail.get())+'\n')
            Settings_File=open(r'c:\ACP\Settings.txt','w')
            Settings_File.writelines(Settings)
            hovertip_update()
            
        autoCompToggle=tk.Checkbutton(self, variable=autoCompToggleVar, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280', command=lambda : settings_save())
        autoCompToggle.grid(row=2,column=0, padx = 5, pady = 0, sticky="N")
        def pullUAG():
            try: last_updated=open(r'c:\ACP\last_updated.txt', 'x')
            except FileExistsError: pass
            current_time=time.time()
            last_updated=open(r'c:\ACP\last_updated.txt', 'w')
            last_updated.write(str(current_time))
            if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Staff,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportEnabled.txt;exit'])
            else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Staff,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportEnabled.txt;'])
            if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Disabled-Users,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportDisabled.txt;exit'])
            else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADUser -Filter * -SearchBase \"OU=Disabled-Users,OU=Users,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\userexportDisabled.txt;'])
            if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Security-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportSecurity.txt;exit'])
            else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Security-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportSecurity.txt;'])
            if autoclose.get() == True: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Distribution-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportDistribution.txt;exit'])
            else: subprocess.call(['runas', '/savecred','/user:'+loginName,r'C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -noexit  \". \'C:\Program Files\Microsoft\Exchange Server\V15\bin\RemoteExchange.ps1\'; Connect-ExchangeServer -auto -ClientApplication:ManagementShell \";Get-ADGroup -Filter * -SearchBase \"OU=Distribution-Groups,OU=Groups,OU=MC-YMCA,DC=ymca,DC=local\" -Properties * | Select-Object name | export-csv -path c:\ACP\groupExportDistribution.txt;'])
            path1=pathlib.Path(r'c:\ACP\userexportEnabled.txt')
            path2=pathlib.Path(r'c:\ACP\userexportDisabled.txt')
            path3=pathlib.Path(r'c:\ACP\groupExportSecurity.txt')
            path4=pathlib.Path(r'c:\ACP\groupExportDistribution.txt')
            def exportLoop(path1, path2, path3, path4, first_modified1, first_modified2, first_modified3, first_modified4, miliseconds=1000):
                if path1.stat().st_mtime == first_modified1 or path2.stat().st_mtime == first_modified2 or path3.stat().st_mtime == first_modified3 or path4.stat().st_mtime == first_modified4:
                    app.after(1000,lambda : exportLoop(path1, path2, path3, path4, first_modified1, first_modified2, first_modified3, first_modified4, miliseconds=1000))
                else:
                    Pull_Groups()
                    Pull_Users()
                    Modify.Name_Entry_Box.set_completion_list(User_List)
            exportLoop(path1, path2, path3, path4, path1.stat().st_mtime, path2.stat().st_mtime, path3.stat().st_mtime, path4.stat().st_mtime,)
            
        pullUAGlabel=tk.Label(self, fg='ghost white',font=("Verdana", 10), bg='#375280', text='Update User and Group Lists: ')
        pullUAGlabel.grid(row=1,column=1, sticky="S")
        
        pullUAGbutton=tk.Button(self, command=lambda : pullUAG(), text="UPDATE")
        pullUAGbutton.grid(row=2,column=1, padx = 5, pady = 0, sticky="N")

        def forceDirsync():
            if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit Invoke-Command -ComputerName dirsync -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test";exit'])
            else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit Invoke-Command -ComputerName dirsync -ScriptBlock {Start-ADSyncSyncCycle -PolicyType Delta}; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test";'])
            #Everything else hinges on the timer function, which is below.
            def forceDirsyncLoop(path, where, first_modified, miliseconds=1000): #The following must be done in a function, since it waits for files to be checked on. 
                #If the path given is equal to the time given, check the path again, after the time provided, default is 1 second.
                if path.stat().st_mtime == first_modified:
                    app.after(miliseconds,lambda : forceDirsyncLoop(path, where, first_modified))
                #This code is structured in a strange format. Every time a file is checked, it is passed a unique code as the 'where' variable. When the file is changed, this code is used to detect what to do next. 
                else:
                    if where == 'dirsync_started':
                        #Pull information as to the last time the sync occurred, which is equivalent to the first_modified time on the python side.
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test2";exit'])
                        else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test2";'])
                        forceDirsyncLoop(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'original_sync_time_query', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime)
                    
                    if where == 'original_sync_time_query':
                        #Once we get the time to hold as the original, we can begin checking the time again every ten seconds, to check for a change.
                        ResultsFile = open(r'c:\ACP\DirsyncUpdateLog2.txt', "r")
                        Lines = ResultsFile.read().splitlines()
                        Create.OriginalSyncTime=Lines[20]
                        if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";exit'])
                        else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";'])
                        forceDirsyncLoop(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'new_sync_time_query', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime,10000)
                    
                    if where == 'new_sync_time_query':
                        #Check if the time is different or not. If it is not, the program runs again, returning to this block. If it is different, it proceeds.
                        ResultsFile = open(r'c:\ACP\DirsyncUpdateLog2.txt', "r")
                        Lines = ResultsFile.read().splitlines()
                        if Lines[20]==Create.OriginalSyncTime:
                            if autoclose.get() == True: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";exit'])
                            else: subprocess.call(['runas', '/savecred', r'/user:YMCA\UserMgmt-AT','C:\\Windows\\System32\\WindowsPowerShell\\v1.0\\powershell.exe -noexit $credential = Import-CliXml -Path \'c:\\ACP\\cred.xml\';Connect-MsolService -Credential $credential; Get-MsolCompanyInformation | Out-File -FilePath c:\\ACP\\DirsyncUpdateLog2.txt -Encoding ASCII; Add-Content -Path c:\\ACP\\DirsyncUpdateLog.txt -Value "test3";'])
                            forceDirsyncLoop(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'new_sync_time_query', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime,10000)
                        else:
                            tk.messagebox.showinfo(title='Sync Complete', message='The manual account sync has been completed.')

            forceDirsyncLoop(pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt'), 'dirsync_started', pathlib.Path(r'c:\ACP\DirsyncUpdateLog.txt').stat().st_mtime)

        forceDirSyncLabel=tk.Label(self, fg='ghost white',font=("Verdana", 10), bg='#375280', text='Sync Accounts to Cloud: ')
        forceDirSyncLabel.grid(row=1,column=2, sticky="S")
        
        forceDirSyncButton=tk.Button(self, command=lambda : forceDirsync(), text="SYNC")
        forceDirSyncButton.grid(row=2,column=2, padx = 5, pady = 0, sticky="N")

        hovertip_label=tk.Label(self, fg='ghost white',font=("Verdana", 10), bg='#375280', text='Hovertips Enabled: ')
        hovertip_label.grid(row=1,column=3, sticky="S")
        
        hovertip_toggle=tk.Checkbutton(self, variable=show_hovertips, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280', command=lambda : settings_save())
        hovertip_toggle.grid(row=2,column=3, padx = 5, pady = 0, sticky="N")
        #Add hovertips to wherever they are relevant.
        def hovertip_update(first=False):
            if show_hovertips.get() == True:
                #name.tip=Hovertip(name, "text", hover_delay=600)


                Create.Name_Entry_Box.tip=Hovertip(Create.Name_Entry_Box, "Enter the user's Full Name. (i.e., John Doe or John M Doe)", hover_delay=600)
                backButton2.tip=Hovertip(backButton2, "Return to the Main Menu", hover_delay=600)
                backButton3.tip=Hovertip(backButton3, "Return to the Main Menu", hover_delay=600)
                backButton4.tip=Hovertip(backButton4, "Return to the Main Menu", hover_delay=600)
                gotoCreate.tip=Hovertip(gotoCreate, "Enter the menu for creating new accounts.", hover_delay=600)
                gotoModify.tip=Hovertip(gotoModify, "Enter the menu for changing existing accounts, \nincluding adding/removing groups, and disabling accounts.", hover_delay=600)
                #gotoSettings_Page.tip=Hovertip(gotoSettings_Page, "Change settings about the program.", hover_delay=600)
                buttonCreds.tip=Hovertip(buttonCreds, "Enter the username and password used \nby the program to run its commands.", hover_delay=600)
                preview.tip=Hovertip(preview, "Generate the details of the account, \nallowing you to ensure all details are correct.", hover_delay=600)
                Create.Description_Entry.tip=Hovertip(Create.Description_Entry, "Displays a short summary of the user's position.\nUsually their job title.", hover_delay=600)
                Create.Description_Lock.tip=Hovertip(Create.Description_Lock, "Choose between manually entering the description\nor the program creating it automatically.", hover_delay=600)
                Create.Inactive_Listbox.tip=Hovertip(Create.Inactive_Listbox, "Groups not enabled for the user, and can be added.\n(Double-Click)", hover_delay=600)
                Create.Active_Listbox.tip=Hovertip(Create.Active_Listbox, "Groups that are enabled for the user, and can be removed. (Double-Click)\n*Note: \"All-YMCA-Staff\" and \"SG-All-YMCA-Staff\" cannot be removed.", hover_delay=600)
                Create.Inactive_Search.tip=Hovertip(Create.Inactive_Search, "Filter the available groups.", hover_delay=600)
                Create.Active_Search.tip=Hovertip(Create.Active_Search, "Filter the added groups.", hover_delay=600)
                Create.AdvancedLabel.tip=Hovertip(Create.AdvancedLabel, "Switch the display of groups to show non-preset groups.\nAlso allows for disabling the license.", hover_delay=600)
                Create.AdvancedToggle.tip=Hovertip(Create.AdvancedToggle, "Switch the display of groups to show non-preset groups.\nAlso allows for disabling the license.", hover_delay=600)
                Create.License_User_Button.tip=Hovertip(Create.License_User_Button, "If the user is not given a license, they will not \nbe able to access Microsoft Office products.", hover_delay=600)
                Create.Create_User_Button.tip=Hovertip(Create.Create_User_Button, "Create the new accounts with the info listed above.", hover_delay=600)
                Create.Supervisor_Dropdown.tip=Hovertip(Create.Supervisor_Dropdown, "Select the supervisor for the account.\nNothing is done automatically, this is only documentation.", hover_delay=600)
                
                Modify.Name_Entry_Box.tip=Hovertip(Modify.Name_Entry_Box, "Enter the full name or the username \nof the account you wish to modify.", hover_delay=600)
                Modify.Search_User_Button.tip=Hovertip(Modify.Search_User_Button, "Pull up the info for the name listed in the box.", hover_delay=600)
                Modify.Description_Entry.tip=Hovertip(Modify.Description_Entry, "Displays a short summary of the user's position.\nUsually their job title.", hover_delay=600)
                Modify.Auto_Toggle.tip=Hovertip(Modify.Auto_Toggle, "Choose between manually entering the description\nor the program creating it automatically.", hover_delay=600)
                Modify.Inactive_Listbox.tip=Hovertip(Modify.Inactive_Listbox, "Groups not enabled for the user, and can be added.\n(Double-Click)", hover_delay=600)
                Modify.Active_Listbox.tip=Hovertip(Modify.Active_Listbox, "Groups that are enabled for the user, and can be removed. (Double-Click)\n*Note: \"All-YMCA-Staff\" and \"SG-All-YMCA-Staff\" cannot be removed.", hover_delay=600)
                Modify.Inactive_Search.tip=Hovertip(Modify.Inactive_Search, "Filter the available groups.", hover_delay=600)
                Modify.Active_Search.tip=Hovertip(Modify.Active_Search, "Filter the added groups.", hover_delay=600)
                Modify.Push_Changes_Button.tip=Hovertip(Modify.Push_Changes_Button, "Opens a review screen to see the changes made.", hover_delay=600)
                Modify.Status_Button.tip=Hovertip(Modify.Status_Button, "Choose whether the account is enabled or disabled.\nDisabled accounts cannot be used, and are meant for employee terminations.", hover_delay=600)
                Modify.AdvancedLabel.tip=Hovertip(Modify.AdvancedLabel, "Switch the display of groups to show non-preset groups.", hover_delay=600)
                Modify.AdvancedToggle.tip=Hovertip(Modify.AdvancedToggle, "Switch the display of groups to show non-preset groups.", hover_delay=600)
                autoCompToggleLabel.tip=Hovertip(autoCompToggleLabel, "The text-box in the Modify/Disable page has auto-complete.\nChoose to have it on or off.", hover_delay=600)
                autoCompToggle.tip=Hovertip(autoCompToggle, "The text-box in the Modify/Disable page has auto-complete.\nChoose to have it on or off.", hover_delay=600)
                pullUAGlabel.tip=Hovertip(pullUAGlabel, "Update the list of groups and users. \n(This is normally done automatically on a schedule.)", hover_delay=600)
                pullUAGbutton.tip=Hovertip(pullUAGbutton, "Update the list of groups and users. \n(This is normally done automatically on a schedule.)", hover_delay=600)
                forceDirSyncLabel.tip=Hovertip(forceDirSyncLabel, "Forces a sync of local accounts to the cloud servers. \nThis can fix some account issues.", hover_delay=600)
                forceDirSyncButton.tip=Hovertip(forceDirSyncButton, "Forces a sync of local accounts to the cloud servers. \nThis can fix some account issues.", hover_delay=600)



            else:
                if first != True: 
                    exityn=tk.messagebox.askokcancel(title='Restart Required', message='A reboot is required to disable hovertips.\nPlease press "OK", and then restart the program.')
                    if exityn == True:
                        exit()

        hovertip_update(True)
        hovertip_label.tip=Hovertip(hovertip_label, "Toggles the popup messages that give extra information. \n(Like this one!)", hover_delay=600)
        hovertip_toggle.tip=Hovertip(hovertip_toggle, "Toggles the popup messages that give extra information. \n(Like this one!)", hover_delay=600)


        autoclose_label=tk.Label(self, fg='ghost white',font=("Verdana", 10), bg='#375280', text='Auto-Close PS Windows: ')
        autoclose_label.grid(row=3,column=0, sticky="S", pady=(15,0))
        
        autoclose_toggle=tk.Checkbutton(self, variable=autoclose, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280', command=lambda : settings_save())
        autoclose_toggle.grid(row=4,column=0, padx = 5, pady = 0, sticky="N")

        sendmail_label=tk.Label(self, fg='ghost white',font=("Verdana", 10), bg='#375280', text='Send Email to Lansweeper: ')
        sendmail_label.grid(row=3,column=1, sticky="S", pady=(15,0))

        sendmail_toggle=tk.Checkbutton(self, variable=send_mail, bg='#375280', padx = 5, fg='#375280', bd=4,activebackground='#375280', command=lambda : settings_save())
        sendmail_toggle.grid(row=4,column=1, padx = 5, pady = 0, sticky="N")

# Driver Code
app = tkinterApp()
Startup()
app.mainloop()