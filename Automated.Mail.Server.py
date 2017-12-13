#!/home/roli/anaconda3/envs/py27/bin/python
import smtplib
import os, glob, time, re, sys, getopt
import getpass
from email.mime.multipart import MIMEMultipart
from email.mime.text import MIMEText
import argparse

## Input
parser = argparse.ArgumentParser(description='Send a Batch of Emails Using a Spreadhseet of Contacts.')
parser.add_argument('--contact_list', help='Location of .tsv formatted spreadsheet with at least three columns: Given Name, Surname and E-mail Address', action='store', type=str)
parser.add_argument('--html_message', help='Location of .html formatted e-mail message', action='store', type=str)
parser.add_argument('--column', help='a column denoting who on the list shall receive an e-mail (mark 1 for send, mark 0 otherwise', action='store', type=str)
parser.add_argument('--sender', help='this should serve as the name of the sender, but hasn\'t been working (it is cosmetic)', action='store', type=str)
parser.add_argument('--subject', help='provide a short and concise e-mail subject line', action='store', type=str)
parser.add_argument('--place_holder', help='chose a word which will be replaced with the addressee\'s name throughout your message', action='store', type=str)

args=parser.parse_args()

if len(sys.argv[1:])==0:
        parser.print_help()
        sys.exit(0)

## HARD CODE INFORMATION
sender_email = ""
PASSWORD = ""  #To Send E-mail with Gmail, you must first create an 'app-specific' password. Visit <https://security.google.com/settings/security/apppasswords>

"""
Example
	./Automated.Mail.Server.py --contact_list e_mail_list.dec.2017.tsv --html_message NYE.2017.html -c 'Holiday List'

Dependencies:
	- Google Server (which often requires special authorization for sending smpt messages: <https://security.google.com/settings/security/apppasswords>
	- firefox browser (or any browser that can open html files from the command line)

Optional:
	- html formatted messages prepare using http://beefree.io/
	- one can download the html file, modify the file so that images are linked to hosted images on beefree.io or elsewhere.
"""


def get_information_from_input(input_file, send_to):
	results = {}

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

	html = re.sub(args.place_holder,client_firstname,html)  ## USING A KEY WORD, SUBSITUTE FOR NAME OF CLIENT

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
	organization = args.sender
	subject = args.subject
	input_file = args.contact_list
	html_message = args.html_message
	send_to = args.column

	# Get Client Information Spreadsheet
	my_results = get_information_from_input(input_file, send_to)

	# Print Information
	print 'Sending email to %i clients' % len(my_results)

	# Show Test html image
        with open(html_message, 'r') as msg_content:
                html_test = msg_content.read()

	html_test = re.sub(args.place_holder,"EACH_CLIENT_NAME_HERE!",html_test) 
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

