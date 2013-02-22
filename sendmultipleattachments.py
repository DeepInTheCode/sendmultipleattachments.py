# Adapted from https://gist.github.com/4009671 and other sources by David Young
# added directory searching functionality to add all files in folder
# disabled username / password logon for use in our Exchange environment


######### Setup your stuff here #######################################

path='//whatever.com/FilePath' # location of files 
archiveFolderName = 'archive' # name of folder under path where files will be archived

host = 'webmail.whatever.com' # specify port, if required, using a colon and port number following the hostname

fromaddr = 'donotreply@whatever.com' # must be a vaild 'from' address in your environment
toaddr  = ['toaddress@whatever.com'] # list of email addresses
replyto = fromaddr # unless you want a different reply-to

# username = 'username' # not used in our Exchange environment
# password = 'password' # not used in our Exchange environment

msgsubject = 'Message Subjectt'

htmlmsgtext = """<h2>The documents are attached.</h2>""" # text with appropriate HTML tags

######### In normal use nothing changes below this line ###############

import smtplib, os, sys, shutil
from datetime import date
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email.MIMEText import MIMEText
from email.Utils import COMMASPACE
from email import Encoders
from HTMLParser import HTMLParser

archivePath = os.path.join(path, archiveFolderName) # full path where files will be archived

if not os.path.exists(archivePath): # create archive folder if it doesn't exist
    os.makedirs(archivePath)
    print 'Archive folder created at ' + archivePath + '.'

# A snippet - class to strip HTML tags for the text version of the email

class MLStripper(HTMLParser):
    def __init__(self):
        self.reset()
        self.fed = []
    def handle_data(self, d):
        self.fed.append(d)
    def get_data(self):
        return ''.join(self.fed)

def strip_tags(html):
    s = MLStripper()
    s.feed(html)
    return s.get_data()

########################################################################


try:
    # Make text version from HTML - First convert tags that produce a line break to carriage returns
    msgtext = htmlmsgtext.replace('</br>',"\r").replace('<br />',"\r").replace('</p>',"\r")
    # Then strip all the other tags out
    msgtext = strip_tags(msgtext)

    # necessary mimey stuff
    msg = MIMEMultipart()
    msg.preamble = 'This is a multi-part message in MIME format.\n'
    msg.epilogue = ''

    body = MIMEMultipart('alternative')
    body.attach(MIMEText(msgtext))
    body.attach(MIMEText(htmlmsgtext, 'html'))
    msg.attach(body)
    attachments = os.listdir(path)

    if 'attachments' in globals() and len('attachments') > 0: # are there attachments?
        for filename in attachments:
            if os.path.isfile(os.path.join(path, filename)):
                f = os.path.join(path, filename)
                part = MIMEBase('application', "octet-stream")
                part.set_payload( open(f,"rb").read() )
                Encoders.encode_base64(part)
                part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(f))
                msg.attach(part)    
    
    msg['From'] = fromaddr
    msg['To'] = COMMASPACE.join(toaddr)
    msg['Subject'] = msgsubject
    msg['Reply-To'] = replyto
    
    print 'To addresses follow:'
    print toaddr

    # The actual email sendy bits
    server = smtplib.SMTP(host)
    server.set_debuglevel(False) # set to True for verbose output
    
    # Comment this block and uncomment the below try/except block if TLS or user/pass is required.
    server.sendmail(fromaddr, toaddr, msg.as_string())
    print 'Email sent.'
    server.quit() # bye bye
	
    # try:
        # # If TLS is used
        # server.starttls()
        # server.login(username,password)
        # server.sendmail(msg['From'], [msg['To']], msg.as_string())
        # print 'Email sent.'
        # server.quit() # bye bye
    # except:
        # # if tls is set for non-tls servers you would have raised an exception, so....
        # server.login(username,password)
        # server.sendmail(msg['From'], [msg['To']], msg.as_string())
        # print 'Email sent.'
        # server.quit() # bye bye
        
    try:
        if 'attachments' in globals() and len('attachments') > 0: # are there attachments?
            for filename in attachments:
                if os.path.isfile(os.path.join(path, filename)):
                    f1 = os.path.join(path, filename)
                    x = filename.find('.')
                    filename2 = filename[:x] + '_' + str(date.today()) + filename[x:]
                    f2 = os.path.join(path, filename2)
                    os.rename(f1, f2)
                    print "File " + filename + " renamed to " + filename2 + "."
                    shutil.move(f2, archivePath)
                    print "File " + filename2 + " moved to " + archivePath + "."
                           
    except:
        print "Files not successfully renamed and/or archived."
        	
except:
    print "Email NOT sent to %s successfully. ERR: %s %s %s " % (str(toaddr), str(sys.exc_info()[0]), str(sys.exc_info()[1]), str(sys.exc_info()[2]) )
    #just in case   
