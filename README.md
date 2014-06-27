onenote-notifier
================

#### What does it do?

Onenote-notifier can be used to email a set of users with a summary of recent changes in notebooks. You can configure 
the exact notebooks, frequency of notifications and the people to be notified. 

#### How does it work? 

Onenote-notifier uses a COM API to communicate with OneNote and search for changes, and then with Outlook to send email 
to recipients. It has a partial object model implemented entirely in Python 3.


#### What are the requirements for onenote-notifier?

* Windows 7 with Python 3.x
* OneNote 2013 or 2010 with your notebooks open
* Outlook 2013 with your email account set up
* [OnePy 0.2] (https://www.github.com/varunsrin/onepy)


#### How do I set up my environment?

* Install Python 3.4 x86 from [here](https://www.python.org/download/releases/3.4.0/) 
* Install PyWin32 for Python 3.4 x86 from [here](http://sourceforge.net/projects/pywin32/files/pywin32/) 
* Add `C:\Python34\` & `C:\Python34\Scripts\` your PATH variable
* Run `C:\Python34\Lib\site-packages\win32com\client\makepy.py`
* Select `Microsoft OneNote 15.0 Extended Type Library`
* Run `pip install onepy`


Having trouble with Dispatch? Check this [SO Post](http://stackoverflow.com/questions/16287432/python-pywin-onenote-com-onenote-application-15-cannot-automate-the-makepy-p)


#### How do I send notifications?

* Run notify.py from the command line with a list of email addresses and a notebook name as arguments
* You can schedule this script to be invoked at regular intervals using Windows Scheduler. 