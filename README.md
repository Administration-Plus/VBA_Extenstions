# VBA_Extenstions
VBA Extensions for common administrative tasks

We will be adding a number of VBA modules that I've created with the assistance of AI that make life easier for a host of Administrative tasks that are useful in the workplace.

**Keeping your Mailings tidy**

a) You've sent out an email newsletter using mailmerge.  Not every organisation has access to paid for sophisticated mailing systems so you took it upon yourself to do it with your office skills. The "Unsubscribe" code added a new VBA macro that scans up to **500 recent inbox emails** and exports sender addresses when “unsubscribe” is detected in either the subject or the first 400 characters of the body, capturing the sender’s email, subject, and received time to an Excel sheet. Run as an Outlook VBA module. You may need to tweak the word count it searches through depending on how long your newsletter is and the position of your unsubscribe instruction. 
b) Part 2 of keeping your mailing lists tidy. You send out your mailmerge and immediately get flooded with bouncebacks from email addresses that are no longer active. Delete all these emails and then run the "Bounceback" VBA module. This one searches your deleted items for common delivery failed messages so you can remove them from your mailing list with a Vlookup in excel and keep things clean.
