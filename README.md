# py-cli-outlook
Python Outlook CLI client

This is a (currently buggy) python script that will get your list of emails, by using a cookie to obtain a token which can then access all emails for ~1 hour. The primary goal of this script is to be faster than the official webpage, which takes ages to load.

Goals:
 - Colours
 - Ability to move and delete emails
 - Marking emails as read
 - Better user interface for input
 - Speediness (Like making outlook actually usable)
 - Better mechanism for finding tokens and refreshing them (it's pretty hacky atm)
 - Support for outlook.live.com (currently only outlook.office.com works)
 - Ability to send emails
