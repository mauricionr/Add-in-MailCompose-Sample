      __  __       _ _    ______                                   __  
     |  \/  | __ _(_) |  / / ___|___  _ __ ___  _ __   ___  ___  __\ \ 
     | |\/| |/ _` | | | | | |   / _ \| '_ ` _ \| '_ \ / _ \/ __|/ _ \ |
     | |  | | (_| | | | | | |__| (_) | | | | | | |_) | (_) \__ \  __/ |
     |_|  |_|\__,_|_|_| | |\____\___/|_| |_| |_| .__/ \___/|___/\___| |
                         \_\                   |_|                 /_/ 

This sample showcases a simple add-in for Outlook that lets you add yourself to the recipient list of an email message that you're composing. 

The easiest way to run this sample is to open it in the playground for Office Add-ins: http://aka.ms/Nahkwl 

====================================================================================

If you're not using the playground for Office Add-ins, follow these steps to run the sample:

1. Host these files either online (e.g. AWS, Azure, Heroku, etc) or on a local host (e.g. a local node.js/python server, etc). Make sure they are served over HTTPS. In the manifest.xml file, change the default value of the SourceLocation to point to the URL where the index.html file is hosted.

2. Now, we want to make this page show up within Outlook as an add-in. Go to Office 365 Mail (https://outlook.office365.com) and login if you aren't already logged in.

Note: You need to already have a subscription to Office 365. If you don't have one, obtain a free account for 30 days from https://portal.microsoftonline.com/Signup/MainSignUp.aspx?OfferId=6881A1CB-F4EB-4db3-9F18-388898DAF510&DL=DEVELOPERPACK

3. Click on the Settings icon in the top right and select "Manage Apps".

4. Click the "+" button, and select "add from file".

5. Add the manifest.xml file and install the app.

6. You should see "My First Add-in" in the list of apps. Make sure it's turned on.

7. Click on Mail on the top navigation bar to go back to your inbox and press "New" to create a new message. Select "Apps" above the "To" field and you should see "My First Mail Add-in" appear.
