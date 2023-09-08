# multistage_approval_gform
A Google app script that built up a multistage approval system where you can nest any Gform under the system and set up a workflow for the Gform. 

You may refer to the example Google spreadsheets and Google form inside [THIS](https://drive.google.com/drive/folders/1Obo61L_kHTIPymv0amOcTFkgqGQjkgMl?usp=sharing) folder.

# Setting up the Multistage Approval System Spreadsheet
The spreadsheet is the heart❤ of the whole system, you may duplicate the provided example **Multistage Approval System** Google sheet into your own Google Drive. 

Inside the spreadsheet, you will see 4 main tabs, which are:

## Mailing List
Mailing List serves as the lookup table for names, Gmail, and school email (or organisation email) based on the positions.

## Workflow
Inside the workflow tab, you will see the different rows of example workflow that I had created, let's understand the function together. 
Request to be AWESOME is a form that is attached to the system. Once there a **new entries**, the system will read the form content and then based on the defined workflow, send it to different personnel to sign off the workflow. 

For the Request to be AWESOME workflow, it will start with the VP Operations to sign-off, the requester (the one who submitted the form). If VP Operations approves the in the approval receipt (what is an approval receipt? It will be explain in later's part), the email will be sent to the President for Sign-off and the VP Operations, the requester will be kept inside the loop (will be CC-ed), and if the President approved, it will go to the Staff for Sign-off and then President, VP Operations, Requester will be kept inside the loop. 

The requester is being kept inside the loop by DEFAULT. 

You might ask whether the other person can FAKE SIGNATURE (or FAKE APPROVAL)? NO! Because each positions is tied to a Gmail account as defined inside the Mailing List, **IF YOU ARE NOT MEANT TO SIGN DURING THAT STAGE, YOUR SIGNATURE WON'T BE RECORDED AND IT WILL NOT AFFECT THE WORKFLOW**

## Blacklisted List
Just a list of blacklisted Gmail accounts, if any of these Gmail accounts submit any form under the system, it will not be responded to!

## Ongoing Workflow
A database to keep track of the ongoing workflow. In other words, don't touch it!

You will also see other Forms connected, which are the **Approval Receipt** and **Request to be AWESOME**, Those forms are just an example, you can remove the form after you have understood the function.

# Setting up of Approval Receipt
1. Duplicate the Approval receipt into your own folder (same location as the **Multistage Approval System Spreadsheet**), it is suggested to not change the Request ID, Request Type, and Options as this field is heavily involved inside the system spreadsheet!
2. Set the form to collect VERIFIED EMAIL address. (Go to Form settings > Responses > Collect Email Address > Verified)
3. Go to Responses > Link to sheets > Select existing spreadsheet > Multistage Approval System
4. Rename the link sheet (Form Response) to "Approval Receipt" **CASE-SENSITIVE❗️**

👍🏻 You might be wondering why there are Pending options. The Approved is easy to understand, the workflow is approved and then it will move to the next stage PIC while Rejected will cancel the whole workflow. Pending acts are sort of like HOLD, which means you can acknowledge the actions. Some use-case is when you already see the submissions, but you will need to check with some of your colleagues on certain items, you can put the process as Pending and put in your comment.

# Nesting of New Form into the system
0. Create a new form (preferably inside the same folder as the **Multistage Approval System**)
1. Set the form to collect VERIFIED EMAIL address. (Go to Form settings > Responses > Collect Email Address > Verified)
2. Set up the Google form as you want it to be😉 (Anything is okay, multiple sections, or attachments, all works!)
3. At the end, the most important step, is to link the response to the Multistage Approval Spreadsheet. (Go to Form Responses > link to sheets > Select existing spreadsheet > Multistage Approval Spreadsheet)
4. You can rename the link sheet (Form Response) to ANY NAME, and then define the workflow inside the **Workflow** tab.

# Setting Up the System
Action checklist! Please make sure your spreadsheet is now well prepared with
1. Mailing list (All the positions, names, Gmail, school email define?)
2. Workflow (All the Form(s) attached, EXCEPT **Approval Receipt** had defined the workflow? )

If both are ready, you are now ready to setup the CODE!
1. Go to the Multistage Approval System > Extensions > Apps Script.
2. Copy the code from the main.gs and paste it into the Apps Script field. (Tips: Ctrl + a to hightlight all, Ctrl + c to copy and then Ctrl + p to paste 😉)
3. There is only one field that needs to modified (assuming you are following all the previous steps) which is the

   ```const approval_form = FormApp.openByUrl("THE EDITABLE FORM LINK FOR THE APPROVAL RECEIPT"); // in edit format```

4. You will need to replace ```"THE EDITABLE FORM LINK FOR THE APPROVAL RECEIPT"``` into the editable form of the Approval Receipt, how to get it? Go the Open up the Approval Receipt Google Form, copy the URL. And then paste it in! VOILA, you are all good to go!
5. Try to run it once, press RUN to execute the apps script, if no ERROR pops up, most likely everything is setup correctly!

# Setting up Automation
After you test the system, you will need to automate it through setting up the trigger. What you need to do are 
1. Go to **Triggers** on the left side bar
2. On the bottom right, select **+ Add Trigger**
3. Choose the ```main``` function to run
4. Select event source > From spreadsheet
5. Select event type > On form submit

Optional but good to do: Failure notification settings > Notify me immediately

# Testing of the system
If everything is setup correctly, you can now try to submit the form that you had attached in. If you are following along, you can try to submit one entry on the **Request to be AWESOME**, the email will be process by the system and will be send to the name,gmail who are tags to the VP Operations positions and so on. 

# Limitations
Too good to be true? Here's the pitfall, this system is not meant to be large-scaled, cause the email that can be sent out by free Gmail account per day through API call is 100 recipients, so in normal usage, I won't put the school email for those Main Committees, unless it goes to the staff, then I will put in the Staf's school email. 

With that being said, you can leave the school emails empty for certain positions BUT NOT the Gmail. Gmail is important as it represents the identity of different positions. 

You may test the remaining function remaining by running the ```testQuota()``` function, the execution log will tell you how many quota left.

# Improvements
1. Currently, if any of the workflow is stuck, means that the person never respond, then the people will need to go and remind the person, probably in the next update might come up with a function to do that.
2. If the correct PIC response to the same form twice, it will cause the system to crash, need to figure out a method to fix that💣.
3. A hold button to put the process on hold, probably need to build a master dashboard to keep track of the whole system. 
4. Kick back option, which will kick it back one stage
