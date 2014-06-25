onenote-notifier
================

What does it do?

Onenote-notifier can be used to email a set of users with a summary of recent changes in notebooks. You can configure 
the exact noteboos, frequency of notifications and the people to be notified. 

How does it work? 

Onenote-notifier uses COM API's to communicate with OneNote and search for changes, and then with Outlook to send email 
to recipients. It has a partial object model implemented entirely in Python 3.


What are the requirements for onenote-notifier?

* You need to run it on a Windows machine with OneNote & Outlook installed (Office 2010, Office 2013)
* The notebook you want to monitor needs to be open in OneNote
* The email address you want to send notifications from needs to be open in Outlook?
* You need to install Python 3 & bindings
* You need your 


How do I set up notifications?

* Open cronjob.py
* Add receipients to EMAIL_LIST 
* Add the notebook you want to monitor to NBK_NAME
* Schedule cronjob.py to be executed at regular intervals
