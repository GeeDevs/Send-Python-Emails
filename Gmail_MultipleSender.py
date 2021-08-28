import smtplib, ssl
import pandas
import time
from email.message import EmailMessage


# Get Initial Count Of Emails In The File 
emailCount = 100


# Import CSV with Emails & Passwords
mailsFile = "./emailsDatabase.csv"
columnNames = ['Email', 'SMTP', 'Password', 'Subject', 'Message']

#Read The CSV data and store it in a DataField
mailData = pandas.read_csv(mailsFile, names=columnNames)

# Run A Loop Based On The Number Of Emails Available & Stop Once Emails Hit Zero
while emailCount > 2:

    

    # Store columns In Separate Lists
    rawEmails = mailData.Email.tolist()
    smtpEmails = mailData.SMTP.tolist()
    rawPassword = mailData.Password.tolist()
    rawSubjects = mailData.Subject.tolist()
    emailMessages = mailData.Message.tolist()

    emailCount = len(rawEmails)

    print(emailCount)
    # Get First Row In The Subject & Message Columns
    mailSubject = "" + rawSubjects[1] + ""
    mailBody = "" + emailMessages[1] + ""
    

    # Start Sending Of Emails
    try:
        port = 587  # For starttls
        smtp_server = "smtp.gmail.com"
        server_email = smtpEmails[1]
        sender_email = rawEmails[1]
        # receiver_email = "buyyatab@gmail.com" # Change This The Email Receiving the Email.
        password = rawPassword[1]
        message = EmailMessage()
        message.set_content("{}".format(mailBody))
        message['Subject'] = mailSubject
        message['From'] = sender_email
        message['To'] = "godwinmuthomim7@gmail.com" # Change This The Email Receiving the Email.
    except Exception as e:
        print(e)
        print("The following error has occured. Kindly restart the program. Also check CSV file if empty. CSV with emails has " + emailCount + " emails remaining and the message body file has " + messageCount + "rows of messages remaining.")

    # Start SMTP Server Connection and Send Email 
    try:   
        context = ssl.create_default_context()
        with smtplib.SMTP(smtp_server, port) as server:
            server.ehlo()  # Possible To Work Without
            server.starttls(context=context)
            server.ehlo()  # Possible To Work Without
            server.login(server_email, password)
            # server.sendmail(sender_email, receiver_email, message)
            server.send_message(message)
            server.quit()
            print("Email sent from ", sender_email) # Simple Log To Check Which Email Sent The Message
    except:
        print("Authentication did not take place. Password for " + sender_email + " might be incorrect or mistyped. Skipping To next email...")
      
    
    mailData = mailData.iloc[1:,:] # Remove First Row After Successful Send
    mailData.to_csv('emailsDatabase.csv', index=False) # Save CSV File After Row Deletion To Avoid Repetition
    
    rawEmails = mailData.Email.tolist()


    emailCount = len(rawEmails) # Get the number of rows remaining in Emails CSV
    print("There are " + str(emailCount) + " Unused Emails Remaining!!!")
    time.sleep(120) # Sleep For 1 Minute Before Sending Another Message

if emailCount >= 1:
    print("Done Sending Emails!")
