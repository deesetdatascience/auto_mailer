# -*- coding: utf-8 -*-
"""
Created on Fri Apr  6 12:54:23 2018

@author: Josiah Outram Halstead
@author: Marc.Matterson
@author: Chris.Taaffe

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
    """Strips HTML tags from a message to create a plain-text version."""

    def __init__(self):
        super().__init__()
        self.reset()
        self.strict = False
        self.convert_charrefs = True
        self.fed = []

    def handle_data(self, d):
        self.fed.append(d)

    def get_data(self):
        return "".join(self.fed)

    def strip_tags(self, html):
        self.feed(html)
        return self.get_data()


def run_from_ipython():
    """Check if we are running inside IPython"""
    try:
        # The below variable is defined in ipython environments
        __IPYTHON__
        return True
    except NameError:
        return False


def preview_message(message_html):
    """Previews an HTML string.
    
    If we are running in IPython, use it display machinary 
    to show a preview of the HTML message. Otherwise, show
    the message in a web browser.
    """

    if run_from_ipython():
        from IPython.display import display_html

        display_html(message_html, raw=True)
        display_html("<hr>", raw=True)
    else:
        import tempfile
        from pathlib import Path
        import webbrowser

        with tempfile.NamedTemporaryFile(delete=False, suffix=".html") as out_file:
            file_path = Path(out_file.name)
            out_file.write(message_html.encode("utf8"))
            webbrowser.open(file_path.as_uri())


# %% Email sender class


class EmailSender:
    """Class that remembers credentials and sends emails"""

    def __init__(self, username=None, password=None):
        """Set up variables for later use"""
        
        if username:
            self.use_relay = False
            # Save passed username and password if passed as args
            self.username = username
            self.password = password
        else:
            self.use_relay = True

        # Set up variales for actual login credentials
        self.gmail_user = None
        self.gmail_password = None

    def get_creds(self):
        """Ask for credentials with getpass if not already provided"""
        # Get credentials if they do not yet exist
        if not self.gmail_user:
            # Get e-mail username from login name if one is not provided
            self.gmail_user = (
                self.username if self.username else f"{getpass.getuser()}@deeset.co.uk"
            )
            # Ask for password if one is not provided
            self.gmail_password = (
                self.password
                if self.password
                else getpass.getpass(
                    f"Enter Password for {getpass.getuser()}@deeset.co.uk: \n"
                )
            )

    def send_email(
        self,
        message_html="",
        subject="",
        sent_from="Data and Analytics Team <report.requests@deeset.co.uk>",
        to=[],
        cc=[],
        bcc=[],
        reciepts=[],
        attachments=[],
        extra_headers={"Return-Path": "<>", "Auto-Submitted": "auto-generated"},
    ):
        """Compile and send a html email message"""

        # We cannot send an email if there is nobody to send it to
        if to == [] and cc == [] and bcc == []:
            print("No recipients")
            return

        # Create a text-only version of the message
        message_text = "\n\n".join(
            [
                x.strip()
                for x in MLStripper().strip_tags(message_html).split("\n")
                if x.strip()
            ]
        )

        # Create message container - the correct MIME type is multipart/alternative.
        msg = MIMEMultipart("alternative")

        # Fill message header fields
        msg["Subject"] = subject
        msg.add_header("reply-to", sent_from)
        msg["From"] = sent_from
        msg["To"] = ", ".join(to)
        msg["Cc"] = ", ".join(cc)
        msg["Bcc"] = ", ".join(bcc)
        # Do we want read-reciepts?
        if reciepts:
            msg["Disposition-Notification-To"] = ", ".join(reciepts)

        # Read, encode and add any attachments
        for f in attachments:
            with open(f, "rb") as fil:
                part = MIMEApplication(fil.read(), Name=os.path.basename(f))
            # After the file is closed
            part[
                "Content-Disposition"
            ] = f'attachment; filename="{os.path.basename(f)}"'
            msg.attach(part)

        # Record the MIME types of both parts - text/plain and text/html.
        part1 = MIMEText(message_text, "plain")
        part2 = MIMEText(message_html, "html")

        # Attach parts into message container.
        # According to RFC 2046, the last part of a multipart message, in this case
        # the HTML message, is best and preferred.
        msg.attach(part1)
        msg.attach(part2)

        # Add any extra headers that we want to the message
        for header, value in extra_headers.items():
            msg.add_header(header, value)

        # Actually send the message
        self._send(msg)

    def check_logged_in(self):
        """Try and check if we are logged in to the the STMP server"""
        try:
            response = self.server.noop()[0] == 250
        except:
            return False
        return response

    def _send(self, msg):
        """Actually send the email message.
        
        Get credentials and log in if necessary. If sending fails, wait and try again.
        """
        if not self.use_relay:
            self.get_creds()
        
        if not self.check_logged_in():
            time.sleep(1)
            self.login()

        try:
            self.server.send_message(msg)
        except Exception as e:
            print("Failed", e)
            print("Waiting 1 minute before trying again")
            time.sleep(60)
            self.login()
            self._send(msg)

    def login(self):
        """Log in to the STMP server"""
        if self.use_relay:
            self.server = smtplib.SMTP_SSL("smtp-relay.gmail.com", 465)
                         #smtplib.SMTP('smtp-relay.gmail.com')
        else:
            self.get_creds()
            self.server = smtplib.SMTP_SSL("smtp.gmail.com", 465)
            self.server.ehlo_or_helo_if_needed()
            try:
                response = self.server.login(self.gmail_user, self.gmail_password)
            except smtplib.SMTPServerDisconnected:
                print("Server Disconnected, waiting 1 minute before trying again")
                time.sleep(60)
                return self.login()
            else:
                if response[0] == 235:
                    print("Logged in successfully")
                else:
                    print("Login Failed - Check username and password")

    def logout(self):
        """Close the connection to the STMP server"""
        self.server.close()
