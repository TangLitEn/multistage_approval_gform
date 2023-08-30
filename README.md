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

You will also see a Form connected, which is the **Request to be AWESOME**, that form is just an example, you can remove the form after you have understood the function.


# Setting up of Approval Receipt

# Nesting of New Form into the system
1. Collect a VERIFIED EMAIL address. (Go to Form settings > Responses > Collect Email Address > Verified)
2. Set up the Google form as you want it to be😉 (Anything is okay, multiple sections, or attachments, all works!)
3. At the end, the most important step, is to link the response to the Multistage Approval Spreadsheet. (Go to Form Responses > link to sheets > Select existing spreadsheet > Multistage Approval Spreadsheet)

# Setting Up the System
From the 
