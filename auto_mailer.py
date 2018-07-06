# -*- coding: utf-8 -*-
"""
Created on Fri Apr  6 12:54:23 2018

@author: Josiah Outram Halstead
@author: Marc.Matterson

"""

# %% Imports

import os
import time
import smtplib
import base64
import getpass
from email.mime.multipart import MIMEMultipart
from email.mime.application import MIMEApplication
from email.mime.text import MIMEText
from html.parser import HTMLParser

# %% Useful functions 



class MLStripper(HTMLParser):
    def __init__(self):
        super().__init__()
        self.reset()
        self.strict = False
        self.convert_charrefs= True
        self.fed = []
    def handle_data(self, d):
        self.fed.append(d)
    def get_data(self):
        return ''.join(self.fed)
    def strip_tags(self, html):
        self.feed(html)
        return self.get_data()

def run_from_ipython():
    try:
        __IPYTHON__
        return True
    except NameError:
        return False

def preview_message(message_html):
    if run_from_ipython():
        from IPython.display import display_html
        display_html(message_html, raw=True)
        display_html('<hr>', raw=True)
    else:
        import tempfile
        from pathlib import Path
        import webbrowser
        with tempfile.NamedTemporaryFile(delete=False, suffix='.html') as out_file:
            file_path = Path(out_file.name)
            out_file.write(message_html.encode('utf8'))
            webbrowser.open(file_path.as_uri())

#%% Email sender class

class EmailSender():

    def __init__(self, username=None, password=None):
        
        self.username = username
        self.password = password
        #self.gmail_user = username if username else input("Enter Gmail Username:")
        self.gmail_user = None
        
    def get_creds(self):
        if not self.gmail_user:
            self.gmail_user = f'{getpass.getuser()}@deeset.co.uk'
            self.gmail_password = self.password if self.password else getpass.getpass(f"Enter Password for {getpass.getuser()}@deeset.co.uk: \n")
            self.login()
        
    def send_email(self,
               message_html='', 
               subject='',
               sent_from="Data and Analytics Team <report.requests@deeset.co.uk>",
               to=[], 
               cc=[], 
               bcc=[], 
               reciepts=[],
               attachments=[],
               extra_headers={'Return-Path':'<>', 'Auto-Submitted':'auto-generated'},
              ):
        
        if to == [] and cc == [] and bcc == []:
            print("No recipients")
            return

        # ============================================================================
        # Create a text-only version of the message
        # =============================================================================

        message_text = '\n\n'.join([x.strip() for x in MLStripper().strip_tags(message_html).split('\n') if x.strip()])

        # ==========================================================================
        # Create the e-mail
        # =============================================================================
    
        # Create message container - the correct MIME type is multipart/alternative.
        msg = MIMEMultipart('alternative')

        msg['Subject'] = subject
        msg.add_header('reply-to', sent_from)
        msg['From'] = sent_from 
        msg['To'] = ", ".join(to)
        msg['Cc'] = ", ".join(cc)
        msg['Bcc'] = ", ".join(bcc)
        if reciepts:
            msg['Disposition-Notification-To'] = ", ".join(reciepts)

        # Record the MIME types of both parts - text/plain and text/html.
        part1 = MIMEText(message_text, 'plain')
        part2 = MIMEText(message_html, 'html')

        # Attach parts into message container.
        # According to RFC 2046, the last part of a multipart message, in this case
        # the HTML message, is best and preferred.
        msg.attach(part1)
        msg.attach(part2)

        for f in attachments:
            with open(f, "rb") as fil:
                part = MIMEApplication(
                    fil.read(),
                    Name=os.path.basename(f)
                )
            # After the file is closed
            part['Content-Disposition'] = f'attachment; filename="{os.path.basename(f)}"'
            msg.attach(part)
            
        for header, value in extra_headers.items():
            msg.add_header(header, value)
    
        self._send(msg)
    
    def check_logged_in(self):
        try:
            response = self.server.noop()[0] == 250
        except:
            return False
        return response
    
    def _send(self, msg):
        self.get_creds()
        if not self.check_logged_in():
            time.sleep(1)
            self.login()
        try:
            self.server.send_message(msg)
        except Exception as e:
            print('Failed', e)
            print('Waiting 1 minute before trying again')
            time.sleep(60) 
            self.login()
            self._send(msg)
    
    def login(self):
        self.server = smtplib.SMTP_SSL('smtp.gmail.com', 465)
        self.server.ehlo_or_helo_if_needed()
        try:
            response = self.server.login(self.gmail_user, self.gmail_password)
        except smtplib.SMTPServerDisconnected:
            print('Server Disconnected, waiting 1 minute before trying again')
            time.sleep(60)            
            return self.login()
        else:
            if response[0] == 235:
                print('Logged in successfully')
            else:
                print('Login Failed - Check username and password')
        
    def logout(self):
        self.server.close()