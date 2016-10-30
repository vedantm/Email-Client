import vsgui as vg	#import the gui file
import sys,os,socket
import poplib, email, string,imaplib,smtplib	#libraries for different protocols
import Tkinter
from Tkinter import *
from tkFileDialog import askopenfilename	
import datetime	#for date time libraries
import shutil	#for removing directory
import urllib2,urllib 	#to check internet connection
#different modules for email parsing 
from email import parser	
from email.MIMEMultipart import MIMEMultipart
from email.MIMEBase import MIMEBase
from email import Encoders
from email.mime.text import MIMEText

global allmails

global info		#would contain servernames and username,password
global inboxdir,sentdir,working,penddir		#directory address of inbox,sentmails and working=1 means connected to internet, else 0
inboxdir=""
sentdir=""
info=[]
penddir=""

def checknet():		#function which tries to connect to some site to check internet connectivity, also proxy settings are set
	global working
	proxy = urllib2.ProxyHandler({'http': 'http://singhp:neeraj@proxy.iitk.ac.in:3128'})
	auth = urllib2.HTTPBasicAuthHandler()
	opener = urllib2.build_opener(proxy, auth, urllib2.HTTPHandler)
	urllib2.install_opener(opener)
	try:
		urllib2.urlopen("http://www.yahoo.com")
	except urllib2.URLError, e:
		#print "Network currently down."
		working=0 
		vg.msgbox("Could not connect to internet!")		
	else:
		working=1
		vg.msgbox("Connected to internet!")
		#print "Up and running." 
	return working

#it downloads the mails in inbox from the server and saves them in ./INBOX/username as similar to sendandsave() and along with attachments.
def downloadinbox():
	global allmails
	#print "IN downloadinbox"
	global info,inboxdir,sentdir
	useimap=1
	if(info[0].find("imap")!=-1 or useimap==1):		#imap server being used
		mail = imaplib.IMAP4_SSL(info[0])
		mail.login(info[2], info[3])
		mail.list()
		mail.select("INBOX") # connect to inbox.

		typ, data = mail.search(None, 'ALL')
		ids = data[0]
		id_list = ids.split()
		#get the most recent email id
		latest_email_id = int( id_list[-1] )
		#iterate through 15 messages in decending order starting with latest_email_id
		#the '-1' makes reverse looping order
		for i in range( latest_email_id, latest_email_id-15, -1 ):
			typ, data = mail.fetch( i, '(RFC822)' )
			body=""
			for response_part in data:
			  if isinstance(response_part, tuple):
				  msg = email.message_from_string(response_part[1])
				  varSubject = msg['subject']
				  varFrom = msg['from']
				  date=msg['date']
				  varFrom = varFrom.replace('<', '')
				  varFrom = varFrom.replace('>', '')
				  SUBJECT="Subject: "+varSubject	
				  FROM="   From: " +varFrom
				  DATE="   Date: "+date
				  inboxdir = "./INBOX/"+info[2]
				  os.chdir(inboxdir)
				  emailfoldername=FROM+SUBJECT+DATE
				  if not os.path.exists(emailfoldername):
					os.makedirs(emailfoldername)
				  if msg.get_content_maintype() == 'multipart' or 1:
					for part in msg.walk():
						if part.get_content_type() == "text/plain":
							body = part.get_payload(decode=True)
							newfile = open(emailfoldername+"/body","w")
							newfile.write(body)
							newfile.close()
							#print "HEEEEEEEEEEEERE"+body
						else: 
							continue
				  os.chdir("../../")

		   #print '[' + varFrom.split()[-1] + '] ' + varSubject + "Date"+date
	else:	# pop server being used
		mailserver = poplib.POP3(info[0])
	#print "connected"
		mailserver.user(info[2]) 
		mailserver.pass_(info[3]) 
		#messages= [mailserver.retr(i) for i in range(1,len(mailserver.list()[1])+1)  ]
		count=21	
		if(len(mailserver.list()[1])<21):
			count=len(mailserver.list()[1])
		messages= [mailserver.retr(i) for i in range(1,count) ]		#displaying & 20 or less mails only
		messages=['\n'.join(m[1]) for m in messages]
		messages= [parser.Parser().parsestr(m) for m in messages] #a list of raw messages
		number = len(mailserver.list()[1])
		#print "You have %d messages in inbox" % (number)
		body=""
		i=0
		for i in range(len(messages)):
			if(messages[i]['Subject']==None):
				continue
			SUBJECT="Subject: "+messages[i]['Subject']	
			FROM="   From: " +messages[i]['From']
			DATE="   Date: "+messages[i]['Date']
			inboxdir = "./INBOX/"+info[2]
			os.chdir(inboxdir)
			emailfoldername=FROM+SUBJECT+DATE
			if not os.path.exists(emailfoldername):
				os.makedirs(emailfoldername)
			else:
				os.chdir("../../")
				continue	
			for part in messages[i].walk():
				if part.get_content_type()=='text/plain':
					body = part.get_payload()
				if part.get('Content-Disposition') is None:
					continue
				else:
					fileName = part.get_filename()
					if bool(fileName):
						filePath=emailfoldername+'/'+fileName
						if not os.path.isfile(filePath) :
							#print fileName
							os.chdir(emailfoldername+'/')
							f = open(fileName, 'wb')
							f.write(part.get_payload(decode=True))
							f.close()
							os.chdir("../")
			newfile = open(emailfoldername+"/body","w")
			newfile.write(body)
			newfile.close()
			os.chdir("../../")
			#for fila in os.listdir('.'):
				#print "FILES:",fila
		#print "inboxdone!"
	return
	
def ConnectionInfo():	
	# it takes the username, password and server details from the user as soon as VSoutlook is started and creates directories for this username, if not present.
	global info
	global inboxdir,sentdir,penddir
	info=[]
	#print "Here"
	msg         = "Enter Connection Details"
	title       = "Connect"
	fieldNames  = ["POP/IMAP Server","SMTP Server","Username","Password"]
	defaultvalues1=["pop.mail.yahoo.com","smtp.mail.yahoo.com","f007khan@yahoo.co.in","qwerty"]
	defaultvalues2=["pop.gmail.com","smtp.gmail.com","vsemailclient@gmail.com","vsemail911"]
	defaultvalues3=["newmailhost.cc.iitk.ac.in","smtp.cc.iitk.ac.in","vinitk",""]
	info = vg.passwordbox(msg,title, fieldNames,'',defaultvalues3)
	start=info[2].find("iitk")
	if(start!=-1):		#if it is iitk id and is of form username@iitk.ac.in,strip off '@iitk.ac.in'
		info[2]=info[2][:start-1]
	working=checknet()
	if (info[0].strip()=="") or (info[1].strip()=="") or (info[2].strip()=="" or info[3].strip()==""):
		vg.msgbox("Enter all details correctly!")
		ConnectionInfo()	
	sentdir = "./SENTMAILS/" + info[2]
	inboxdir = "./INBOX/" + info[2]
	if not os.path.exists("./PENDING"):
			os.makedirs("./PENDING")
	penddir = "./PENDING/" + info[2]
	if not os.path.exists(sentdir):
			os.makedirs(sentdir)
	if not os.path.exists(inboxdir):
			os.makedirs(inboxdir)
	if not os.path.exists(penddir):
			os.makedirs(penddir)		
	if(working==0):
		if not os.path.exists(sentdir):
			os.makedirs(sentdir)
		if not os.path.exists(inboxdir):
			os.makedirs(inboxdir)
	else:	#internet working
		if not os.path.exists(sentdir):
			os.makedirs(sentdir)
			#shutil.rmtree(sentdir)
			#os.makedirs(sentdir)
		if not os.path.exists(inboxdir):
			os.makedirs(inboxdir)
			#shutil.rmtree(inboxdir)
			#os.makedirs(inboxdir)
		downloadinbox()	
	return 

#it takes as argument various details of the mail to be sent and sends it and saves a copy of the email in ./SENTMAILS/username/ folder folder with name as details of email (including Date-time which makes sure of unique names) and in that folder, the body of the mail is kept in a file named 'body'.
def sendandsave(Values="",iors=0):
	global info,inobxdir,sentdir,working
	if(Values[0]==""):
		vg.msgbox("Invalid Input!")
		return
	if(working==0):
		vg.msgbox("Not connected to internet!")
		return	
	msg=MIMEMultipart()
	msg['From']=info[2]
	msg['To']=Values[0]
	msg['Cc']=Values[1]
	msg['Bcc']=Values[2]
	msg['Date']=email.Utils.formatdate(localtime=True)
	msg['Subject']=Values[3]
	msg.attach(MIMEText(Values[5]))
	
	if(Values[4]!=""):
		part=MIMEBase('application', "octet-stream")
		part.set_payload(open(Values[4],"rb").read())
		Encoders.encode_base64(part)
		part.add_header('Content-Disposition', 'attachment; filename="%s"' % os.path.basename(Values[4]))
		msg.attach(part)

	server=smtplib.SMTP(info[1])
	EMAIL_TO = [msg['To'],msg['Cc'],msg['Bcc']]
	EMAIL_SERVER=info[1]
	EMAIL_FROM=info[2]
	server = smtplib.SMTP(EMAIL_SERVER)
	server.ehlo()
	if(info[2].find("gmail")!=-1):
		server.starttls()
		server.ehlo()
	server.login(info[2],info[3])
	server.sendmail(EMAIL_FROM, EMAIL_TO, msg.as_string())
	
	os.chdir(sentdir)
	emailfoldername="To:" + Values[0]+"   CC:"+Values[1]+"  BCC:"+Values[2]+ "  "+"Subject:" +Values[3] +"  Date:"+msg['Date']
	if not os.path.exists(emailfoldername):
		os.makedirs(emailfoldername)
		newfile = open(emailfoldername+"/body","w")
		body=Values[5]
		newfile.write(body)
		newfile.close()
		os.chdir("../../")
		return
	else:
		os.chdir("../../")
		return

def pending():
	global working,info,inboxdir,sentdir,penddir
	checknet()
	Values=[]
	for i in range(6):
		Values.append("")
	if(working==0 or len(os.listdir(penddir))==0):
		return
	else:
		for mails in os.listdir(penddir):
			start=mails.find("To:")
			end=mails.find("CC:")
			#print "Start: ",start,"End: ",end, "Mails: ",mails
			Values[0]=(mails[start+3:end])
			start=mails.find("CC:")
			end=mails.find("BCC:")	
			#print "Start: ",start,"End: ",end
			Values[1]=(mails[start+3:end])
			start=mails.find("BCC:")
			end=mails.find("Subject:")
			#print "Start: ",start,"End: ",end
			Values[2]=(mails[start+4:end])
			start=mails.find("Subject:")
			end=mails.find("Date:")
			#print "Start: ",start,"End: ",end
			Values[3]=(mails[start+8:end])
			body=""
			os.chdir(penddir+"/"+mails)
			f1=open("body",'r')
			for i, line in enumerate(f1):
				body+=line
			f1.close()
			Values[5]=body
			os.chdir("../")
			shutil.rmtree(mails)
			os.chdir("../../")
			sendandsave(Values)
	return		
			
			
			
def sendlater(Values=[],iors=0):
	global info,inobxdir,sentdir,working
	if(Values[0]==""):
		vg.msgbox("Invalid Input!")
		return
	
	os.chdir(penddir)
	emailfoldername="To:" + Values[0]+"   CC:"+Values[1]+"  BCC:"+Values[2]+ "  "+"Subject:" +Values[3] +"  Date:"+email.Utils.formatdate(localtime=True)
	if not os.path.exists(emailfoldername):
		os.makedirs(emailfoldername)
		newfile = open(emailfoldername+"/body","w")
		body=Values[5]
		newfile.write(body)
		newfile.close()
		os.chdir("../../")
		return
	else:
		os.chdir("../../")
		return

#it is the central windows' function which controls which function to call and when. Depending upon the last button press, it decides whether to display sent mails or the inbox mails.
def Home(iors=0):
	global info,inboxdir,sentdir,working,penddir
	checknet()
	if(working==1):
		pending()
	if(iors==0 and working==1):
		downloadinbox()
	#OR download sentmails
		
	#print "download done"
	if(iors==0):
		msg = "SHOWING INBOX"
	else:
		msg = "SHOWING SENT MAILS"
		
	title       = "VSoutlook"
	#choices = ["Inbox","Forward", "Reply", "Compose","Display Mail","Connection Settings","Sent Mails","Exit"]
	choices = ["Inbox","Forward", "Reply","Compose","Display Mail","Sent Mails","Exit"]

	#append all the mails to 'mails' array to make radiobuttons for them
	displaynum=0
	if(iors==0):
		mails=[]
		for files in os.listdir(inboxdir):
			mails.append(files)
			displaynum+=1
			if(displaynum==15):
				break
	elif(iors==1):
		mails=[]
		for files in os.listdir(sentdir):
			mails.append(files)		
			displaynum+=1
			if(displaynum==15):
				break
	
	response = vg.buttonbox(msg, title,choices,mails)
	#print "Response0:",response[0]
	#print "Response1:",response[1]
		
	# mycomposebox returns Values as follows:
	#	0:Forward to addr
	#	1:CC
	#	2:BCC
	#	3:SUBJECT
	#	4:Attachment file name
	#	5:Body
	
	if(response[0]=="Exit"):
		checknet()
		if(working==1):
			pending()
		exit()

	if(response[0]=="Inbox"):
		checknet()
		if(working==1):
			pending()
		Home(0)
		
	#if(response[0]=="Connection Settings"):
		#ConnectionSetting()	
		#Home()

	if(response[0]=="Forward"):
		#look at the value of reponse[1] representing mail number
		msg         = "Enter the details:"
		title       = "Forward Mail"
		Names  = ["TO:","CC:","BCC:","SUBJECT:","ATTACH FILE:","BODY:"]
		Values = []  # we start with blanks for the values

		body = ""
		forwardto = ""
		if(mails==[]):#no mail to forward!
			Home(iors)
		filename=mails[response[1]]
		if(iors==0):
			os.chdir(inboxdir+"/"+filename)
		elif(iors==1):
			os.chdir(sentdir+"/"+filename)	
		f1=open("body",'r')
		for i, line in enumerate(f1):
			body+=line
		f1.close()
		os.chdir("../../../")
		#Take the subject and body from the mail itself
		start=filename.find("Subject:")
		end=filename.find("Date:")
		subject = "Fwd: "+ (filename[start+len("Subject:"):end]).strip()	
		
		Values = vg.mycomposebox(msg,title, Names,forwardto,subject,body)
		checknet()
		#print Values
		if(Values!=None and working==1):
			sendandsave(Values)
		elif(Values!=None and working==0):
			sendlater(Values)	
		#in the end, return to home
		
		Home(iors)
		
		
	if(response[0]=="Reply"):
		#look at the value of reponse[1] representing mail number
		msg         = "Enter the details:"
		title       = "Reply to Mail"
		Names  = ["TO:","CC:","BCC:","SUBJECT:","ATTACH FILE:","BODY:"]
		Values = []  # we start with blanks for the values
		
		if(mails==[]):
			Home(iors)
		filename=mails[response[1]]
		if(iors==0):
			os.chdir(inboxdir+"/"+filename)
		elif(iors==1):
			os.chdir(sentdir+"/"+filename)	
		
		replyto=""
		#Take value whom to reply from the mail itself
		if(iors==0):
			start=filename.find("From:")	
			end=filename.find("Subject:")
			if(start!=-1 and end!=-1):
				replyto+=(filename[start+len("From:"):end]).strip()
		else:
			start=filename.find("To:")	
			end=filename.find("CC:")
			if(start!=-1 and end!=-1):
				replyto+=(filename[start+len("To:"):end]).strip()

		forwardto = replyto
		subject = "Re: "
		body = ""
		os.chdir("../../../")	
		Values = vg.mycomposebox(msg,title, Names,forwardto,subject)
		checknet()
		#print Values
		if(Values!=None and working==1):
			sendandsave(Values)
		elif(Values!=None and working==0):
			sendlater(Values)	
		#in the end, return to home
		Home(iors)
		
	#if(response[0]=="Delete"):
		#pass
		#look at the value of reponse[1] representing mail number
		#Home()

	if(response[0]=="Compose"):
		msg         = "Enter the details:"
		title       = "Compose Mail"
		Names  = ["TO:","CC:","BCC:","SUBJECT:","ATTACH FILE:","BODY:"]
		Values = []  # we start with blanks for the values
		Values = vg.mycomposebox(msg,title, Names)
		#print Values
		checknet()	
		if(Values!=None and working==1):
			sendandsave(Values)
		elif(Values!=None and working==0):
			sendlater(Values)	
		#in the end, return to home
		Home(iors)		
	
	if(response[0]=="Display Mail"):
		if(mails==[]):
			Home(iors)
		filename=mails[response[1]]
		if(iors==0):
			os.chdir(inboxdir+"/"+filename)
		elif(iors==1):
			os.chdir(sentdir+"/"+filename)	

		body=""
		a=0
		f=open("body",'r')
		for i, line in enumerate(f):
			body+=line
		f.close()
		if(iors==0):
			vg.textbox("Displaying mail from inbox:\n\n"+filename,"Received Mail",body)
		else:
			vg.textbox("Displaying mail from sentmails:\n\n"+filename,"Sent Mail",body)
		os.chdir("../../../")
		checknet()
		if(working==1):
			pending()
		Home(iors)
		
	if(response[0]=="Sent Mails"):
		checknet()
		if(working==1):
			pending()
		Home(1)

ConnectionInfo()	
#print "Here!"		
Home(0)	
	
#def exit():
	#sys.exit(1)

	





















