
import tkinter, smtplib, openpyxl, os
#import stat, binascii
from tkinter import messagebox
from tkinter.filedialog import *
from email.mime.multipart import MIMEMultipart
from email.mime.base import MIMEBase
from email.mime.text import MIMEText
from email.utils import formatdate
from email import encoders
from smtplib import *
import oauth2 #oauth2.py
import webbrowser, os, sys, subprocess
from datetime import datetime

class App(tkinter.Tk):
    def __init__(self, parent):
        tkinter.Tk.__init__(self,parent)
        self.parent = parent
        self.initialize()
    def initialize(self):
        # Window title
        self.wm_title("Mass Emailing App")
        #self.config(background = "#FFFFFF")

        # Draw frames
        leftFrame = tkinter.Frame(self, width=200, height = 200)
        leftFrame.grid(row=0, column=0, padx=10, pady=10)
        rightFrame = tkinter.Frame(self, width=200, height = 200)
        rightFrame.grid(row=0, column=1, padx=10, pady=10)

        # Create drop-down menus
        self.sheet_num = tkinter.StringVar()
        self.sheet_num.set('1')
        self.lbl_sheet_num = tkinter.Label(rightFrame, text='Sheet number: ')
        self.lbl_sheet_num.grid(row=1, column=0, padx=10, pady=2, sticky='E')
        self.sheet_numOptions = tkinter.OptionMenu(rightFrame, self.sheet_num,'1', '2', '3', '4', '5', '6', '7', '8', '9', '10')
        self.sheet_numOptions.grid(row=1, column=1, sticky='E')
        
        self.rowNum = tkinter.StringVar()
        self.rowNum.set('1')
        self.lbl_rowstart = tkinter.Label(rightFrame, text='Start row: ')
        self.lbl_rowstart.grid(row=2, column=0, padx=10, pady=2, sticky='E')
        self.rowOptions = tkinter.OptionMenu(rightFrame,self.rowNum,'1', '2', '3', '4', '5', '6', '7', '8', '9', '10')
        self.rowOptions.grid(row=2, column=1, sticky='E')

        self.names_col = tkinter.StringVar()
        self.names_col.set('A')
        self.lbl_names_col = tkinter.Label(rightFrame, text='Names Column: ')
        self.lbl_names_col.grid(row=3, column=0, padx=10, pady=2, sticky='E')
        self.names_colOptions = tkinter.OptionMenu(rightFrame,self.names_col,'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J')
        self.names_colOptions.grid(row=3, column=1, sticky='E')

        self.emails_col = tkinter.StringVar()
        self.emails_col.set('B')
        self.lbl_emails_col = tkinter.Label(rightFrame, text='Emails Column: ')
        self.lbl_emails_col.grid(row=4, column=0, padx=10, pady=2, sticky='E')
        self.emails_colOptions = tkinter.OptionMenu(rightFrame,self.emails_col,'A', 'B', 'C', 'D', 'E', 'F', 'G', 'H', 'I', 'J')
        self.emails_colOptions.grid(row=4, column=1, sticky='E')

        self.greeting = tkinter.StringVar()
        self.greeting.set('Hello')
        self.lbl_greeting = tkinter.Label(rightFrame, text='Greeting: ')
        self.lbl_greeting.grid(row=5, column=0, padx=10, pady=2, sticky='E')
        self.greetingOptions = tkinter.OptionMenu(rightFrame,self.greeting,'Hello','Dear','Hi','Greetings','To whom it may concern','Dear Sir/Madam')
        self.greetingOptions.grid(row=5, column=1, sticky='E')

        self.lbl_files = tkinter.Label(rightFrame, text="Files attached: ")
        self.lbl_files.grid(row=6, column=0, padx=10, pady=2, sticky='E')
        self.lbl_files2 = tkinter.Label(rightFrame, text="")
        self.lbl_files2.grid(row=6, column=1, padx=10, pady=2, sticky='E', rowspan=10)

        self.btn = tkinter.Button(leftFrame, text="Send Emails from Spreadsheet", command=self.read_spreadsheet, width=25)
        self.btn.grid(row=1, column=0, padx=10, pady=2, sticky='W')

        self.btn_1 = tkinter.Button(leftFrame, text="Do Not Email List", command=self.edit_unwanted, width=25)
        self.btn_1.grid(row=2, column=0, padx=10, pady=2, sticky='W')

        self.btn_2 = tkinter.Button(leftFrame, text="Edit Email Subject", command=self.email_subject, width=25)
        self.btn_2.grid(row=3, column=0, padx=10, pady=2, sticky='W')

        self.btn_3 = tkinter.Button(leftFrame, text="Edit Email Body", command=self.email_body, width=25)
        self.btn_3.grid(row=4, column=0, padx=10, pady=2, sticky='W')

        self.btn_4 = tkinter.Button(leftFrame, text="Attach files", command=self.attach_files, width=25)
        self.btn_4.grid(row=5, column=0, padx=10, pady=2, sticky='W')

        self.btn_5= tkinter.Button(leftFrame, text="Remove all attached files", command=self.delete_files, width=25)
        self.btn_5.grid(row=7, column=0, padx=10, pady=2, sticky='W')

        self.loading = tkinter.Label(leftFrame, text='SENDING EMAILS...', font=('TkDefaultFont', 30), justify='center')
        self.loading.grid(row=1, column=0)
        self.loading.grid_remove()

        self.info = []
        self.files = []
        self.email = ''
        self.auth_code = ''

        self.GOOGLE_CLIENT_ID = 'your_client_id'
        self.GOOGLE_CLIENT_SECRET = 'your_client_secret'

    def hide_widgets(self):
        self.lbl_sheet_num.grid_remove()
        self.sheet_numOptions.grid_remove()
        self.lbl_rowstart.grid_remove()
        self.rowOptions.grid_remove()
        self.lbl_names_col.grid_remove()
        self.names_colOptions.grid_remove()
        self.lbl_emails_col.grid_remove()
        self.emails_colOptions.grid_remove()
        self.lbl_greeting.grid_remove()
        self.greetingOptions.grid_remove()
        self.lbl_files.grid_remove()
        self.lbl_files2.grid_remove()
        self.btn.grid_remove()
        self.btn_1.grid_remove()
        self.btn_2.grid_remove()
        self.btn_3.grid_remove()
        self.btn_4.grid_remove()
        self.btn_5.grid_remove()
        
        self.loading.grid()

        self.update()

    def show_widgets(self):
        self.lbl_sheet_num.grid()
        self.sheet_numOptions.grid()
        self.lbl_rowstart.grid()
        self.rowOptions.grid()
        self.lbl_names_col.grid()
        self.names_colOptions.grid()
        self.lbl_emails_col.grid()
        self.emails_colOptions.grid()
        self.lbl_greeting.grid()
        self.greetingOptions.grid()
        self.lbl_files.grid()
        self.lbl_files2.grid()
        self.btn.grid()
        self.btn_1.grid()
        self.btn_2.grid()
        self.btn_3.grid()
        self.btn_4.grid()
        self.btn_5.grid()

        self.loading.grid_remove()

        self.update()

    def edit_unwanted(self):
        os.startfile('Do_not_email.xlsx')

    def email_subject(self):
        os.startfile('Email_subject.txt')

    def email_body(self):
        os.startfile('Email_body.txt')

    def attach_files(self):
        # Maximum number of files is 5
        if len(self.files) >= 5:
            messagebox.showerror(' ', 'You can only attach up to 5 files')
            return
        filenamex = askopenfilename()
        if filenamex == '':
            # Return if user presses cancel button
            return
        elif filenamex in self.files:
            messagebox.showerror(' ', 'You have already attached this file.')
            return
        elif os.path.getsize(filenamex)/1048576 > 20:
            messagebox.showerror(' ', 'File must be smaller than 20MB. \nPlease link large files as a Google Drive link instead.')
            return
        else:
            # Append file path to self.files
            self.files.append(filenamex)

        # Set label text to show attached files
        filelabeltext = ''
        for file in self.files:
            filename = os.path.basename(file)
            filelabeltext = filelabeltext + filename + '\n'

        self.lbl_files2['text'] = filelabeltext

    def delete_files(self):
        self.files = []
        self.lbl_files2['text'] = ''

    def read_spreadsheet(self):
        self.info = []
        
        names_col = self.names_col.get()
        emails_col = self.emails_col.get()
        # Get user to choose Excel file with info
        filenamex = askopenfilename(filetypes=([("Excel files", "*.xlsx")]))
        if filenamex == '':
            # Return if user presses cancel button
            return

        # Load spreadsheet
        wb = openpyxl.load_workbook(filenamex)
        sheets = wb.sheetnames
        sheet_num = int(self.sheet_num.get())

        # Show error message if sheet number user selected
        # does not exist
        if len(wb.sheetnames) < sheet_num:
            messagebox.showerror(' ', 'The sheet number you have selected does not exist. \nPlease select a valid sheet number.')
            return

        # Find the correct sheet in the file
        sheet = wb[sheets[sheet_num - 1]]

        # Read from spreadsheet
        for row in range(int(self.rowNum.get()), sheet.max_row + 1):
            
            # Skip if no email is listed
            if sheet[emails_col + str(row)].value is None:
                continue

            # If no name is listed, put empty string
            elif sheet[names_col + str(row)].value is None:
                name = ''
            else:
                name = sheet[names_col + str(row)].value
            email = sheet[emails_col + str(row)].value

            # Each person's info is stored in a dictionary
            data = {'Name':name, 'Email':email}
            self.info.append(data)

        # Show error message if length of info list is zero
        if len(self.info) == 0:
            messagebox.showerror(' ', 'No emails found.')
            return

        elif len(self.info) > 75:
            messagebox.showerror(' ', 'Please select a spreadsheet with less than 75 rows.')
            return

        # Check for unicode errors
        self.check_unicode()

    def remove_unwanted(self):
        # Retrieve list of people who did not want to be emailed
        # and remove from info list
        emailsFile = openpyxl.load_workbook('Do_not_email.xlsx')
        sheets = emailsFile.sheetnames

        # Find first sheet
        sheet = emailsFile[sheets[0]]
        unwanted_emails = []

        # Read from spreadsheet
        for row in range(1, sheet.max_row + 1):
            
            # Skip if no email is listed
            if sheet['A' + str(row)].value is None:
                continue
            else:
                unwanted_emails.append(sheet['A' + str(row)].value.strip())

        emailsFile.close()

        for item in unwanted_emails:
            for i in range(len(self.info)):
                for person in self.info:
                    if item.strip() == person['Email']:
                        self.info.remove(person)

    def login(self):
        # Email login in new window
        login_window = Toplevel(height=140, width=300)
        login_window.geometry("+500+250")
        login_window.title("Email Login")

        msg = Label(login_window, text='Please use a gmail account.')
        msg.place(x=20, y=10)

        emailLabel = Label(login_window, text='Email: ')
        emailEntry = Entry(login_window, width=30)
        emailLabel.place(x=20, y=40)
        emailEntry.place(x=80, y=40)

        def get_email():
            self.email = emailEntry.get()

        button2 = Button(login_window, text='Ok', width=10, command=lambda: [get_email(), login_window.destroy(), self.authenticate()])
        button2.place(x=20,y=90)

        button = Button(login_window, text="Cancel", width=10, command=login_window.destroy)
        button.place(x=150,y=90)
        
        login_window.grab_set()

    def check_unicode(self):
        subjectFile = open('Email_subject.txt')
        bodyFile = open('Email_body.txt')
        subject = subjectFile.read()
        body = bodyFile.read()
        subjectFile.close()
        bodyFile.close()

        # Check for unicode error
        try:
            splitsubject = subject.split()
            for n in splitsubject:
                unicode_data = n.encode('ascii')
        except UnicodeEncodeError:
            messagebox.showerror(' ', 'Unicode Error in email subject: \nCharacter: %s' %(n))
            return

        try:
            splitbody = body.split()
            for n in splitbody:
                unicode_data = n.encode('ascii')
        except UnicodeEncodeError:
            messagebox.showerror(' ', 'Unicode Error in email body: \nCharacter: %s' %(n))
            return

        for n in range(len(self.info)):
            try:
                unicode_name = self.info[n]['Name'].encode('ascii')
            except UnicodeEncodeError:
                messagebox.showerror(' ', 'Unicode Error in Name: %s' %(self.info[n]['Name']))
                return
            except:
                messagebox.showerror(' ', 'Error reading name: %s' %(self.info[n]['Name']))
                return

        self.login()

    def auth_code_window(self):
        # Let user type in authentication code obtained from browser url
        auth_window = Toplevel(height=140, width=300)
        auth_window.geometry("+500+250")
        auth_window.title("Authentication Code")

        authLabel = Label(auth_window, text='Authentication Code: ')
        authEntry = Entry(auth_window, width=30)
        authLabel.place(x=20, y=40)
        authEntry.place(x=20, y=60)

        def get_code(code):
            self.auth_code = code

        button2 = Button(auth_window, text='Ok', width=10, command=lambda: [get_code(authEntry.get()), auth_window.destroy()])
        button2.place(x=20,y=90)

        button = Button(auth_window, text="Cancel", width=10, command=auth_window.destroy)
        button.place(x=150,y=90)
        
        auth_window.grab_set()
        auth_window.wait_window()

    def authenticate(self):
        email = self.email
        
        # Get auth string using oauth2
        refreshFile = open('Refresh_token.txt')
        refresh_token = refreshFile.read()
        refreshFile.close()
        if refresh_token == '':

            # If no refresh token has been obtained (new user) yet
            # Direct user to authentication url to get code
            url = oauth2.GeneratePermissionUrl(self.GOOGLE_CLIENT_ID)
            webbrowser.open_new_tab(url)
            self.auth_code_window()
            auth_code = self.auth_code
            if auth_code == '':
                # User pressed cancel
                return

            # Save refresh token
            try:
                response = oauth2.AuthorizeTokens(self.GOOGLE_CLIENT_ID, self.GOOGLE_CLIENT_SECRET, auth_code)
                access_token = response['access_token']
                refresh_token = response['refresh_token']

                refreshFile2 = open('Refresh_token.txt', 'w')
                refreshFile2.write(refresh_token)
                refreshFile2.close()

                auth_string = oauth2.GenerateOAuth2String(email, access_token)
            except Exception as e:

                # Error procedure
                messagebox.showerror(' ', 'Authentication error')
                
                errorFileRelativeLocation = 'Error logs//' + 'Error log ' + datetime.now().strftime('%Y-%m-%d %H;%M;%S') + '.txt'
                errorFileLocation = os.path.abspath(errorFileRelativeLocation)
                errorFile = open(errorFileLocation, 'w')
                errorFile.write(str(e))
                errorFile.close()

                if sys.platform == "win32":
                    os.startfile(errorFileLocation)
                else:
                    opener ="open" if sys.platform == "darwin" else "xdg-open"
                    subprocess.call([opener, errorFileLocation])
                return

        # Get auth code using access code
        else:
            try:
                response = oauth2.RefreshToken(self.GOOGLE_CLIENT_ID, self.GOOGLE_CLIENT_SECRET, refresh_token)

            # If refresh token is revoked
            except Exception as e:
                open('Refresh_token.txt', 'w').close()
                self.authenticate()
                return

            access_token = response['access_token']
            auth_string = oauth2.GenerateOAuth2String(email, access_token)

        self.send_emails(auth_string)

    def send_emails(self, auth_string):

        self.hide_widgets()

        email = self.email

        # Remove unwanted emails
        self.remove_unwanted()
        
        # Set up SMTP server and authenticate with auth code
        smtpObj = smtplib.SMTP('smtp.gmail.com', 587)
        smtpObj.ehlo()
        smtpObj.starttls()
        smtpObj.ehlo()
        smtpObj.docmd('AUTH', 'XOAUTH2 ' + auth_string)

        # Retrieve email subject and body
        subjectFile = open('Email_subject.txt')
        bodyFile = open('Email_body.txt')
        subject = subjectFile.read()
        body = bodyFile.read()
        subjectFile.close()
        bodyFile.close()

        # Greeting
        greetingchoice = self.greeting.get()
            
        # List of unsent emails
        unsent = []

        # Error log file
        errorFileRelativeLocation = 'Error logs//' + 'Error log ' + datetime.now().strftime('%Y-%m-%d %H;%M;%S') + '.txt'
        errorFileLocation = os.path.abspath(errorFileRelativeLocation)
        errorFile = open(errorFileLocation, 'w')        
        
        # Send email
        for n in range(len(self.info)):
            # Display greeting based on if there is a name or not
            if greetingchoice == 'Dear' and self.info[n]['Name'] =='':
                greeting = 'Dear Sir/Madam' + ','
            elif greetingchoice == 'To whom it may concern' or greetingchoice == 'Dear Sir/Madam' or self.info[n]['Name'] =='':
                greeting = greetingchoice + ','
            else:
                greeting = greetingchoice + ' ' + self.info[n]['Name'] + ','

            # Create email body with attachments using MIME
            msg = MIMEMultipart()
            msg['From'] = email
            msg['To'] = self.info[n]['Email']
            msg['Date'] = formatdate(localtime=True)
            msg['Subject'] = subject

            msg.attach(MIMEText(greeting +'\n\n' + body))

            for path in self.files:
                part = MIMEBase('application', "octet-stream")
                with open(path, 'rb') as file:
                    part.set_payload(file.read())
                encoders.encode_base64(part)
                part.add_header('Content-Disposition',
                                'attachment; filename="{}"'.format(os.path.basename(path)))
                msg.attach(part)

            try:
                smtpObj.sendmail(email, self.info[n]['Email'], msg.as_string())
            except SMTPServerDisconnected or SMTPConnectError as e:
                messagebox.showerror(' ', 'You have exceeded the email rate limit. Please try again in 2 minutes.')
                errorFile.write(str(e))
                errorFile.close()
                if sys.platform == "win32":
                    os.startfile(errorFileLocation)
                else:
                    opener ="open" if sys.platform == "darwin" else "xdg-open"
                    subprocess.call([opener, errorFileLocation])
                self.show_widgets()
                return
            except SMTPSenderRefused or SMTPAuthenticationError as e:
                messagebox.showerror(' ', 'Authentication error')
                errorFile.write(str(e))
                errorFile.close()
                if sys.platform == "win32":
                    os.startfile(errorFileLocation)
                else:
                    opener ="open" if sys.platform == "darwin" else "xdg-open"
                    subprocess.call([opener, errorFileLocation])
                self.show_widgets()
                return
            except SMTPRecipientsRefused as e:
                unsent.append(self.info[n]['Email'])
                errorFile.write(str(e))
            except Exception as e:
                unsent.append(self.info[n]['Email'])
                errorFile.write(str(e))

        self.show_widgets()

        if len(unsent) == 0:
            messagebox.showinfo(' ', 'Emails sent')
        else:
            # Show list of unsent emails
            emailstext = ''
            for item in unsent:
                emailstext = emailstext + item + '\n'
            messagebox.showerror(' ', 'Emails sent \nError sending emails to: \n%s' %(emailstext))
        smtpObj.quit()

        # If errors occurred, open error log file. Otherwise delete empty file
        errorFile.close()
        errorFile = open(errorFileLocation, 'r')
        errorText = errorFile.read()
        errorFile.close()
        if errorText == '':
            os.remove(errorFileLocation)
        else:
            if sys.platform == "win32":
                os.startfile(errorFileLocation)
            else:
                opener ="open" if sys.platform == "darwin" else "xdg-open"
                subprocess.call([opener, errorFileLocation])

if __name__ == '__main__':
    app = App(None)
    app.geometry("+400+250")
    app.mainloop() #this will run until it closes


