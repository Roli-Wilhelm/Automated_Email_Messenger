#!/usr/bin/env python
import smtplib
import os, glob, time, re, sys, getopt
import getpass
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText

Usage = """
Usage:
		./Automated.Mail.Server.py -l input_e_mail_list.tsv -m message.html -c column_name_to_send_to

Example:
		./Automated.Mail.Server.py -l e_mail_list.dec.2016.ts -m NYE.2016.html -c 'Local Party List'

Dependencies:
	- Google Server (which often requires special authorization for sending smpt messages: <https://security.google.com/settings/security/apppasswords>
	- firefox browser (or any browser that can open html files from the command line)

Optional:
	- html formatted messages prepare using http://beefree.io/
	- one can download the html file, modify the file so that images are linked to hosted images on beefree.io or elsewhere.
	- if one wishes to personalize the address, include "MEPHISTO" in the address and this script will automatically substitute that with the client name
"""


## Get Inputs
input_file = ''
send_to_column_name = ''
html_message = ''

try:
	opts, args = getopt.getopt(sys.argv[1:],"l:m:c:")
except getopt.GetoptError:
	print Usage
	sys.exit(2)
for opt, arg in opts:
	if opt == '-h':
		print Usage
		sys.exit()
	elif opt == "-l":
		input_file = arg
	elif opt == "-m":
		html_message = arg
	elif opt == "-c":
		send_to = arg

def get_information_from_input(input_file):
	results = {}
	e_mail_column = 0
	send_to_column = 0

	#Find Column with E-mail
	with open(input_file, 'r') as f:
		header = f.readline()

	header = header.strip()
	header = header.split("\t")
	e_mail_column = header.index("E-mail")
	send_to_column = header.index(send_to)

	print e_mail_column
	with open(input_file, 'r') as f:
		next(f) # Skips first line

		for line in f:
			line = line.strip()
			line = line.split('\t')

			# Filter Sender List
			try:
				line[send_to_column]

				if int(line[send_to_column]) == 1:
					given_name = line[1]
					surname = line[2]
					client_email = line[int(e_mail_column):len(line)]
					results[surname+"_"+given_name] = [surname, given_name, client_email]

			except IndexError:
				pass

	return (results)

def send_email(subject, organization, sender_email, sender_password, client_email, client_information, html_message):
	# Client Information
	client_surname = client_information[0]
	client_firstname = client_information[1]

	msg = MIMEMultipart('alternative')
	msg['Subject'] = subject
	msg['From'] = organization
	msg['To'] = client_firstname+" "+client_surname

	with open(html_message, 'r') as msg_content:
		html = msg_content.read() 

	html = re.sub("MEPHISTO",client_firstname,html)  ## USING A KEY WORD, SUBSITUTE FOR NAME OF CLIENT

	# Record the MIME type and Attach message to container.
	msg.attach(MIMEText(html, 'html'))

	smtpObj = smtplib.SMTP('smtp.gmail.com:587')  #smtp.office365.com:587
	smtpObj.ehlo()
	smtpObj.starttls()
	smtpObj.login(sender_email, sender_password)
	smtpObj.sendmail(sender_email, client_email, msg.as_string())
	smtpObj.quit()
	print 'Email sent to %s' %client_email

def main():
	#sender_email = raw_input("Hello!\nLet's Get Started! Please enter sender's e-mail address:\n") 
	#print('Please enter the sender e-mail\'s password')
	#PASSWORD = getpass.getpass()
	organization = raw_input("Who shall be named as the sender of the message?\n")
	subject = raw_input("What will be the subject of your message?\n")

	## HARD CODE INFORMATION
	sender_email = "roliwilhelm@gmail.com"
	PASSWORD = "grjxeqebltpbfmje"  #To Send E-mail with Gmail, you must first create an 'app-specific' password. Visit <https://security.google.com/settings/security/apppasswords>
	#organization = "Roli Wilhelm - Party Industries"
	#subject = "NYE Party Invitation - Restrospective @ 3546 West 43rd Ave"

	# Get Client Information Spreadsheet
	my_results = get_information_from_input(input_file)

	# Print Information
	print 'Sending email to %i clients' % len(my_results)

	# Show Test html image
        with open(html_message, 'r') as msg_content:
                html_test = msg_content.read()

	html_test = re.sub("MEPHISTO","EACH_CLIENT_NAME_HERE!",html_test) 
	output = open("html_test.html","w")
	output.write(html_test)
	output.close()

	print "\n\nIs The HTML Message Correct (it takes a second to load)? [note: formatting will mirror exactly what you see]\n\n"

	os.system(' '.join([
		"firefox",
		"html_test.html &"
	]))

	complete = raw_input("Write \'Yes\' to continue or Cntrl+C to Cancel.\n")
	os.system('rm html_test.html')

	if complete == "Yes":
		for key, value in my_results.iteritems():
			client_email = value[2]
			client_information = [value[0],value[1]]
			send_email(subject, organization, sender_email, PASSWORD, client_email, client_information, html_message)
	else:
		print "Sending Aborted."

main()

