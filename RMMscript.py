
#main imports and calls
import time
import datetime
import sys
# import keyboard
from selenium import webdriver
from selenium.webdriver.common.by import By
from selenium.webdriver.common.keys import Keys
from webdriver_manager.chrome import ChromeDriverManager
from email.mime.text import MIMEText
from email.mime.multipart import MIMEMultipart
import tkinter as tk
import tkinter.simpledialog as simpledialog
import tkinter.messagebox as msgbox
# from smtplib import SMTP_SSL as SMTP 
import smtplib, ssl
import ctypes


def user_input():

    #get user inputs (exe file will not run without variables this here for even as global variables)
    global userEmail 
    global userPassword 
    global numCommands
    global commandList

    #set up tkinter window 
    root = tk.Tk()
    root.withdraw()
    #root.iconbitmap(default='ITG.ico')
    boxTitle = "RMM Powershell Automation" 
    
    #informational message with instructions
    if msgbox.askyesno(boxTitle, "Welcome to my hacky but also pretty well thought out Powershell automation program. Please enter the devices serial numbers you would like to run Powershell commands on, one device per line, in the devices.txt file. Once complete, hit continue", parent=root) == False:
        sys.exit()

    #ask for email address and save it to userEmail variable
    userEmail = simpledialog.askstring(boxTitle, "      Please enter your Datto email address:      ", parent=root)
    if userEmail == None:
        sys.exit()

    #ask for password and save it to userPassword variable
    userPassword = simpledialog.askstring(boxTitle, "      Please enter your Datto password:      ", parent=root, show="*")
    if userPassword == None:
        sys.exit()

    #get number of commands to be entered
    numCommands = simpledialog.askinteger(boxTitle, "     How many commands would you like to run per device? (1-5)     ", parent=root, minvalue=1, maxvalue=5)
    if numCommands == None:
        sys.exit()

    #Create array to hold commands
    commandList = []

    #get commands from user.Loops through number of commands based off numCommands variable
    for i in range(numCommands):
        commandList.append(simpledialog.askstring(boxTitle, "                                     Please enter PowerShell Command " + str(i+1) + ":                                    ", parent=root))
        if commandList[i] == None:
            sys.exit()

        #check for empty string, if empty show error message and ask for input again
        while commandList[i] == "":
            msgbox.showerror(boxTitle, "No Command detected")
            commandList[i] = simpledialog.askstring("Autotask RMM Powershell Script", "No command detected. Enter PowerShell Command " + str(i+1) + ":", parent=root)
            if commandList[i] == None:
                sys.exit()

    #show user inputs and ask if they are correct
    if msgbox.askyesno(boxTitle, "You entered the following commands: " + str(commandList) + ". Would you like to continue?") == False:
        sys.exit()

    #show popup message that the RMM serialnumber column needs to e displayed
    if msgbox.askyesno(boxTitle, "Please make sure your RMM is defualt to open in new UI and the serialnumber column is displayed. If not the program will not work.") == False:
        sys.exit()

    #show lessage that all log files will be cleared and ask if they want to continue
    if msgbox.askyesno(boxTitle, "All log files will be cleared. Would you like to continue?") == False:
        sys.exit()

    #returns this variale. Again shouldnt need with global variables but exe didnt work without it
    return commandList

        
#this section loads the browser and logs in
def load_browser():

    url = "https://auth.datto.com/login" #url to login page
    
    #try and load browser url. If it fails, show error message and wait 5 seconds before trying again
    #if it fails 2 times, exit program and update errorsLogs file
    try:
        driver.get(url)
        #print(str(datetime.datetime.now()) + " - Loaded " + url)
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Loaded " + url + "\n")
            f.close()
    except:
        #print(str(datetime.datetime.now()) +" - Failed to load " + url + " ....waiting to try again")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Failed to load " + url + " ....waiting to try again" + "\n")
            f.close()
        time.sleep(5)
        try:
            driver.get(url)
            #print(str(datetime.datetime.now()) + " - Second attempt loading " + url)
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Second attempt loading " + url + "\n")
                f.close()
        except:
            #print(str(datetime.datetime.now()) +" - Failed second attempt to load " + url + "....quiting")
            with open("errorsLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Browser failed to load " + url + " for second time....quiting" + "\n")
                f.close()
            driver.quit()
            sys.exit()

    #wait for page to load
    time.sleep(2)
    
    #locate username field, enter username and hit enter
    try:
        drivers(By.ID, "form_username").send_keys(userEmail, Keys.ENTER)
        #print(str(datetime.datetime.now()) + " - Entering username ")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Entering username " + "\n")
            f.close()
    except:
        #print(str(datetime.datetime.now()) + " - Failed to enter username....waiting to try again")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Failed to enter username....waiting to try again" + "\n")
            f.close()
        time.sleep(5)
        try:
            drivers(By.ID, "form_username").send_keys(userEmail, Keys.ENTER)
            #print(str(datetime.datetime.now()) + " - Second attempt entering username ")
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Second attempt entering username " + "\n")
                f.close()
        except:
            #print(str(datetime.datetime.now()) + " - Failed second attempt to enter username....quiting")
            with open("errorsLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) +  " - Failed to enter username....quiting" + "\n" )
                f.close()
            driver.quit()
            sys.exit()

    #wait for page to load
    time.sleep(2)

    #locate password field,enter password hit enter
    try:
        drivers(By.ID, "form_password").send_keys(userPassword, Keys.ENTER)
        #print(str(datetime.datetime.now()) + " - Entering password ")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Entering password " + "\n")
            f.close()
    except:
        #print(str(datetime.datetime.now()) + " - Failed to enter password....waiting to try again")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Failed to enter password....waiting to try again" + "\n")
            f.close()
        time.sleep(5)
        try:
            drivers(By.ID, "form_password").send_keys(userPassword, Keys.ENTER)
            #print(str(datetime.datetime.now()) + " - Second attempt entering password ")
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Second attempt entering password " + "\n")
                f.close()
        except:
            #print(str(datetime.datetime.now()) + " - Failed second attempt to enter Password....quiting")
            with open("errorsLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Failed to enter password....quiting" + "\n")
                f.close()
            driver.quit()
            sys.exit()

    #*****wait for you you enter MFA code and hit enter*****, if user has mfa off, no wait time is needed
    time.sleep(20)

    #try and see is MFA feild is shown. If it is, user is not logged in and script will exit
    try:
        if drivers(By.ID, "authy_token").is_displayed() == True:
            #print(str(datetime.datetime.now()) + " - MFA code not entered....quiting")
            with open("errorsLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - MFA code not entered....quiting" + "\n")
                f.close()
            driver.quit()
            sys.exit()
    except:
        #print(str(datetime.datetime.now()) + " - MFA code entered")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - MFA code entered" + "\n")
            f.close()
        pass

    
#this section searches for device, records its status and returns status value to main
def search_RMM():

    time.sleep(2)
    driver.switch_to.window(driver.window_handles[1]) #switch to new tab
    #print(str(datetime.datetime.now()) + " - Switched to new tab")
    with open("scriptLog.txt", "a") as f:
        f.write(str(datetime.datetime.now()) + " - Switched to new tab" + "\n")
        f.close()

    #try to search page via searchbox. If it fails, show error message and wait 5 seconds before trying again
    try:
        #print(str(datetime.datetime.now()) + " - Searching for " + serialNumber) 
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Searching for " + serialNumber + "\n")
            f.close()
        drivers(By.XPATH, "//input[@id='headerSearch']").click()
        drivers(By.XPATH, "//input[@id='headerSearch']").send_keys(serialNumber)
        time.sleep(2)  
        drivers(By.XPATH, "//input[@id='headerSearch']").send_keys(Keys.ARROW_DOWN)
        drivers(By.XPATH, "//input[@id='headerSearch']").send_keys(Keys.ENTER)
    except:
        #print(str(datetime.datetime.now()) + " - Failed to search for " + serialNumber + " ....waiting to try again")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Failed to search for " + serialNumber + " ....waiting to try again" + "\n")
            f.close()
        time.sleep(5)
        try:
            #print(str(datetime.datetime.now()) + " - Second attempt searching for " + serialNumber)
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Second attempt searching for " + serialNumber + "\n")
                f.close()
            drivers(By.XPATH, "//input[@id='headerSearch']").click()
            drivers(By.XPATH, "//input[@id='headerSearch']").send_keys(serialNumber)
            time.sleep(2)            
            drivers(By.XPATH, "//input[@id='headerSearch']").send_keys(Keys.ARROW_DOWN)
            drivers(By.XPATH, "//input[@id='headerSearch']").send_keys(Keys.ENTER)
        except:
            #print(str(datetime.datetime.now()) + " - Failed second attemot to search for " + serialNumber + "....quiting")
            with open("errorsLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Failed to search for " + serialNumber + " ....quiting" + "\n")
                f.close()
            # driver.quit()
            # sys.exit()

    time.sleep(2)

    #look for serialnumber column. If it fails, show error message and wait 5 seconds before trying again
    try:         
        drivers(By.XPATH, "//td[@title='" + serialNumber + "']")
        #print(str(datetime.datetime.now()) + " - Searching for device serial number displayed in device row")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Searching for device serial number displayed in device row" + "\n")
            f.close()
        isDevice = True
    except:
        #print(str(datetime.datetime.now()) + " - No results found for " + serialNumber + " ....waiting to try again")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - No results found for " + serialNumber + " ....waiting to try again" + "\n")
            f.close()
        time.sleep(1)
        try:
            drivers(By.XPATH, "//td[@title='" + serialNumber + "']")
            #print(str(datetime.datetime.now()) + " - Second attempt searching for device serial number displayed in device row")
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Second attempt searching for device serial number displayed in device row" + "\n")
                f.close()
            isDevice = True
        except:
            #print(str(datetime.datetime.now()) + " - Failed second attempt to find results for " + serialNumber)
            with open("errorsLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - No results found for "  + serialNumber + "\n")
                f.close()
            isDevice = False

    time.sleep(2)

    #if device is found, get device status
    if isDevice == True:  
        try:
            drivers(By.XPATH, "//button[contains(text(),'Web Remote')]").click()
            #print(str(datetime.datetime.now()) + " - Web Remote button found for " + serialNumber)
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Web Remote button found for " + serialNumber + "\n")
                f.close()
            status = "online"
        except:
            #print(str(datetime.datetime.now()) + " - Web Remote button not found for " + serialNumber + " ....waiting to try again")
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Web Remote button not found for " + serialNumber + " ....waiting to try again" + "\n")
                f.close()
            time.sleep(1)
            try:
                drivers(By.XPATH, "//button[contains(text(),'Web Remote')]").click()
                #print(str(datetime.datetime.now()) + " - Second Attempt Web Remote button found for " + serialNumber)
                with open("scriptLog.txt", "a") as f:
                    f.write(str(datetime.datetime.now()) + " - Second Attempt Web Remote button found for " + serialNumber + "\n")
                    f.close()
                status = "online"
            except:
                #print(str(datetime.datetime.now()) + " - Failed second attempt to find Web Remote button for " + serialNumber +", device offline ")
                #error log
                with open("scriptLog.txt", "a") as f:
                    f.write(str(datetime.datetime.now()) + " - Failed second attempt to find Web Remote button for " + serialNumber +", device offline " + "\n")
                    f.close()
                status = "offline"
    else:
        status = "does not exist"
        #print(str(datetime.datetime.now()) + " - No results found in RMM for " + serialNumber)
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - No results found in RMM for " + serialNumber + "\n")
            f.close()
    return status

#this section opens PowerShell, if the device is online it will run a command. Then is closes the window and clears the search box
def device_ps():

    time.sleep(1)
    driver.switch_to.window(driver.window_handles[2]) #switch to new tab that is opened
    with open("scriptLog.txt", "a") as f:
        f.write(str(datetime.datetime.now()) + " - Switching to new tab" + "\n")
        f.close()
    time.sleep(1)

    #click on powershell link
    try:
        drivers(By.LINK_TEXT, "PowerShell").click()
        #print(str(datetime.datetime.now()) + " - Opening PowerShell for " + serialNumber)
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Opening PowerShell for " + serialNumber + "\n")
            f.close()
    except:
        #print(str(datetime.datetime.now()) + " - Failed to open PowerShell for " + serialNumber + " ....waiting to try again")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Failed to open PowerShell for " + serialNumber + " ....waiting to try again" + "\n")
            f.close()
        time.sleep(5)
        try:
            drivers(By.LINK_TEXT, "PowerShell").click()
            #print(str(datetime.datetime.now()) + " - Second Attempt opening PowerShell for " + serialNumber)
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Second Attempt opening PowerShell for " + serialNumber + "\n")
                f.close()
        except:
            #print(str(datetime.datetime.now()) + " - Failed second attempt to open PowerShell for " + serialNumber + "\n")
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Failed second attempt to open PowerShell for " + serialNumber + "\n")
                f.close()
            #error log
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Failed to open PowerShell for " + serialNumber +  "\n")
                f.close()
                # driver.quit()
                # sys.exit()

    time.sleep(3)

    #click into powershell window
    try:
        drivers(By.ID, "app").click()
        #print(str(datetime.datetime.now()) + " - Clicking on PowerShell window for " + serialNumber)
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Clicking on PowerShell window for " + serialNumber + "\n")
            f.close()
    except:
        #print(str(datetime.datetime.now()) + " - Failed to click on PowerShell window for " + serialNumber + " ....waiting to try again")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Failed to click on PowerShell window for " + serialNumber + " ....waiting to try again" + "\n")
            f.close()
        time.sleep(5)
        try:
            drivers(By.ID, "app").click()
            #print(str(datetime.datetime.now()) + " - Second attempt clicking on PowerShell window for " + serialNumber)
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Second attempt clicking on PowerShell window for " + serialNumber + "\n")
                f.close()
        except:
            #print(str(datetime.datetime.now()) + " - Failed second attempt to click on PowerShell window for " + serialNumber)
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Failed second attempt to click on PowerShell window for " + serialNumber + "\n")
                f.close()
            #error log
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Failed to click on PowerShell window for " + serialNumber + "\n")
                f.close()
                # driver.quit()
                # sys.exit()

    time.sleep(3)
    
    #loop through and run commands in powershell console
    for i in range(numCommands):
        try:
            drivers(By.XPATH, "//body/div[@id='app']/div[1]/div[1]/div[1]/div[2]/div[1]/textarea[1]").send_keys(commandList[i])
            time.sleep(1)
            drivers(By.XPATH, "//body/div[@id='app']/div[1]/div[1]/div[1]/div[2]/div[1]/textarea[1]").send_keys(Keys.ENTER)
            time.sleep(1)
            #print(str(datetime.datetime.now()) + " - Running command " + commandList[i] + " on " + serialNumber)
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Running command " + commandList[i] + " on " + serialNumber + "\n")
                f.close()
        except:
            #print(str(datetime.datetime.now()) + " - Failed to run command on " + serialNumber + " ....waiting to try again")
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Failed to run command on " + serialNumber + " ....waiting to try again" + "\n")
                f.close()
            time.sleep(5)
            try:
                drivers(By.XPATH, "//body/div[@id='app']/div[1]/div[1]/div[1]/div[2]/div[1]/textarea[1]").send_keys(commandList[i])
                time.sleep(1)
                drivers(By.XPATH, "//body/div[@id='app']/div[1]/div[1]/div[1]/div[2]/div[1]/textarea[1]").send_keys(Keys.ENTER)
                time.sleep(1)
                #print(str(datetime.datetime.now()) + " - Second attempt running command " + commandList[i] + " on " + serialNumber)
                with open("scriptLog.txt", "a") as f:
                    f.write(str(datetime.datetime.now()) + " - Second attempt running command " + commandList[i] + " on " + serialNumber + "\n")
                    f.close()
            except:
                #print(str(datetime.datetime.now()) + " - Failed second attempt to run command on " + serialNumber)
                with open("scriptLog.txt", "a") as f:
                    f.write(str(datetime.datetime.now()) + " - Failed second attempt to run command on " + serialNumber + "\n")
                    f.close()
    time.sleep(3)
    
    driver.close() #close current tab and move focus back to RMM tab
    driver.switch_to.window(driver.window_handles[1]) #switch back to tab with RMM
    with open("scriptLog.txt", "a") as f:
        f.write(str(datetime.datetime.now()) + " - Closing PowerShell window and tab for " + serialNumber + "\n")
        f.close()

    time.sleep(3)
                  
    try:
        if drivers(By.XPATH, "//input[@id='headerSearch']").is_displayed() == True:
            drivers(By.XPATH, "//input[@id='headerSearch']").clear()
            #print(str(datetime.datetime.now()) + " - Clearing search box for device " + serialNumber)
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Clearing search box for device " + serialNumber + "\n")
                f.close()
        else:
            #print(str(datetime.datetime.now()) + " - Failed to clear search box for device " + serialNumber + " ....waiting to try again")
            with open("scriptLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) + " - Failed to clear search box for device " + serialNumber + " ....waiting to try again" + "\n")
                f.close()
            time.sleep(5)
    except:
        #print(str(datetime.datetime.now()) + " - Failed to clear search box for device " + serialNumber + " ....waiting to try again")
        with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Failed to clear search box for device " + serialNumber + " ....waiting to try again" + "\n")
            f.close()
        time.sleep(5)
        try:
            if drivers(By.XPATH, "//input[@id='headerSearch']").is_displayed() == True:
                drivers(By.XPATH, "//input[@id='headerSearch']").clear()
                #print(str(datetime.datetime.now()) + " - Second attempt clearing search box for device " + serialNumber)
                with open("scriptLog.txt", "a") as f:
                    f.write(str(datetime.datetime.now()) + " - Second attempt clearing search box for device " + serialNumber + "\n")
                    f.close()
        except:
            #print(str(datetime.datetime.now()) + " - Failed second attempt to clear search box for device " + serialNumber)
            with open("errorsLog.txt", "a") as f:
                f.write(str(datetime.datetime.now()) +  " - Failed to clear search box for " + serialNumber + "\n")
                f.close()
                driver.quit()
                sys.exit()


def email_error_log(): #send email with error log
    sender_email = "****"
    receiver_email = userEmail
    password = "*******"
    message = MIMEMultipart("alternative")
    message["Subject"] = "RMM Powershell Automation Crash Report"
    message["From"] = sender_email
    message["To"] = receiver_email

    with open("scriptLog.txt", "a") as f:
        f.write(str(datetime.datetime.now()) + " - Sending email to " + receiver_email + "\n")
        f.close()

    html = """\
        <html>
        <body>
            This is an automatted message from the RMM Powershell Script.<br>
            <br>
            Script started at """ + str(startTime) + """<br>
            Script ran for """ + str(totalTime) + """<br>
            <br>
            The following errors were encountered:<br>
            """
    with open("errorsLog.txt", "r") as f:
        for line in f:
            html += line + "<br>"
        f.close()



def email_log():

    #set email information 
    sender_email = "email"
    receiver_email = userEmail
    password = "pass"
    message = MIMEMultipart("alternative")
    message["Subject"] = "RMM Powershell Automation Results" 
    message["From"] = sender_email
    message["To"] = receiver_email

    with open("scriptLog.txt", "a") as f:
        f.write(str(datetime.datetime.now()) + " - Sending email to " + receiver_email + "\n")
        f.close()
    

    #create an html version of the message
    html = """\
    <html>
    <body>
        This is an automatted message from the RMM Powershell Script.<br>
        <br>
        Script started at """ + str(startTime) + """<br>
        Script ran for """ + str(totalTime) + """<br>
        <br>
        Number of devices script tested: """ + str(deviceTotal) + """<br>
        Number of devices that were online: """ + str(isOnline) + """<br>
        <br>
        Number of devices that were offline: """ + str(isOffline) + """<br>
        """
    with open("isOffline.txt", "r") as f:
        for line in f:
            html += line + "<br>"
        f.close()
  

    with open("scriptLog.txt", "a") as f:
            f.write(str(datetime.datetime.now()) + " - Adding second part of html to email" + "\n")
            f.close()
            
    #add second part of html
    html += """\
        <br>
        number of devices that do not exist in the RMM: """ + str(doesNotExist) + """<br>
        """
    with open("doesNotExist.txt", "r") as f:
        for line in f:
            html += line + "<br>"
        f.close()
    #add third part of html
    html += """\
        <br>
    </body>
    </html>
    """
  
    part1 = MIMEText(html, "html") #turn these into plain/html MIMEText objects

    
    message.attach(part1) #add HTML parts to MIMEMultipart message

    with open("scriptLog.txt", "a") as f:
        f.write(str(datetime.datetime.now()) + " - Creating secure connection with server and sending email" + "\n")
        f.close()

    #create secure connection with server and send email
    context = ssl.create_default_context()
    with smtplib.SMTP("smtp.office365.com", 587) as server:
        server.ehlo()
        server.starttls(context=context)
        server.ehlo()
        server.login(sender_email, password)
        server.sendmail(sender_email, receiver_email, message.as_string())

    with open("scriptLog.txt", "a") as f:
        f.write(str(datetime.datetime.now()) + " - Email sent to " + receiver_email + "\n")
        f.close()

#Program starts here
if __name__ == "__main__":  

    #kill cmd window 
    ctypes.windll.user32.ShowWindow(ctypes.windll.kernel32.GetConsoleWindow(), 0)

    startTime1 = datetime.datetime.now() #start timer
    startTime = startTime1.strftime("%H:%M:%S") #format start time
    #print(str(datetime.datetime.now()) + " - Welcome to the Autotask RMM Powershell Script") #opens and #prints to terminal windows
    #creates variables as global so they can be accessed by all functions   
    global url
    global serialNumberList
    global isOnline
    global isOffline
    global doesNotExist
    global deviceTotal
    global status
    global driver
    global drivers
    global serialNumber    
    global totalTime

    user_input() #call function to get all user inpits needed

    #clear all files content to start fresh
    open('errorsLog.txt', 'w').close()
    open('isOffline.txt', 'w').close()
    open('doesNotExist.txt', 'w').close()
    open('isOnline.txt', 'w').close()
    open('scriptLog.txt', 'w').close()

    #defines variables
    status = "On"
    isOffline = 0
    isOnline = 0
    doesNotExist = 0
    
    
    #turns off notifications
    options = webdriver.ChromeOptions()
    # options.add_argument("--disable-notifications")
    options.add_experimental_option('excludeSwitches', ['enable-logging'])

    driver = webdriver.Chrome(ChromeDriverManager().install(), options=options ) #installs chrome driver
    drivers = driver.find_element #sets find_element command to its own variable and shortens the code
    

    load_browser() #call load browser function to open browser and log in
    

    #set windows to not be seen while running
    driver.set_window_position(0, -2000)
    #driver.set_window_size(0, 0)
    #set command prompt windo to not be seen while running
    # hwnd = win32gui.GetForegroundWindow()
    # win32gui.ShowWindow(hwnd, 0)



    with open("devices.txt") as f: #reads serialNumber in from a text file
        num_lines = sum(1 for line in open("devices.txt")) #count the number of lines in the file
        deviceTotal = num_lines #set deviceTotal to the original/max number of lines in the file aka devices in file
        
        #while not 0, loop through each line in the file
        while num_lines > 0:
            #loops through each line of the text file
            for serialNumber in f:
                serialNumber = serialNumber.strip() #removes any extra white space from the serialNumber
                #print(str(datetime.datetime.now()) + " - " + str(num_lines) + " devices left to run")
            
                num_lines = num_lines - 1 #starts counitin down the number of devices left to run
                
                status = search_RMM() #call search RMM function to search for device and return status, dont need function set equal to return value with global variables

                #if device is online
                if status == "online":
                    #print(str(datetime.datetime.now()) + " - Device " + serialNumber + " is online")

                    #add serialNumber to isOnline file
                    with open("isOnline.txt", "a") as f:
                        f.write(str(serialNumber+ "\n"))
                        f.close()

                    device_ps() #call device_ps function to open powershell and run commands
                    isOnline = isOnline + 1 #counter for number of devices that are online

                #if device is offline
                elif status == "offline":
                    #print(str(datetime.datetime.now()) + " - " + serialNumber + " is offline")
                    isOffline = isOffline + 1 #counter for number of devices that are offline
                    
                    #add serialNumber to isOffline file
                    with open("isOffline.txt", "a") as f:
                        f.write(serialNumber + "\n")
                        f.close()
                        continue
                # if not online or offline then it does not exist in the RMM
                else:
                    #print(str(datetime.datetime.now()) + " - " + serialNumber + " does not exist")
                    doesNotExist = doesNotExist + 1 #counter for number of devices that do not exist in the RMM

                    #add serialNumber to doesNotExist file
                    with open("doesNotExist.txt", "a") as f:
                        f.write(str(serialNumber +"\n"))
                        f.close()

            #when loop finishes, aka all devices have been tested, come here
            #print(str(datetime.datetime.now()) + " - There are no devices left in the files")
            endTime1 = datetime.datetime.now() #get current time
            endTime = endTime1.strftime("%H:%M:%S") #format time
            totalTime = endTime1 - startTime1 #get total time it took to run script
            #print(str(datetime.datetime.now()) + " - Total time to run script: " + str(totalTime)) ##print current time in hh:mm:ss format

        
        email_log()#jump to email log function to send email with results
        
        #close browser and end program properly
        driver.quit()
        sys.exit()
