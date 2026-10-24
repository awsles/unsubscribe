![Outlook Unsubscribe](images/Promo_Banner2.png)

**CURRENT VERSION:** v1.0.0.11  (21 September 2026)

# Outlook Unsubscribe
*Tired of junk email?  Me too!* Even being careful with where and when I share my email address, I was still getting 40+ a day. And there was no "easy" way to go through an unsubscribe. You had to open the latest message from each sender and then hunt for the tiny "unsubscribe" link within the message. *Did you know that most bulk email contains a hidden [RFC-8058](https://certified-senders.org/wp-content/uploads/2017/07/CSA_one-click_list-unsubscribe.pdf) one-click unsubscribe?* One click to that and your remove request is sent!  No login. No questions to answer. If this hidden field is missing, the add-in attempts to find the unsubscribe link embedded in the email body.

I decided to create an Outlook add-in button that would find this hidden link and call it. I enhanced this with a fallback to look for "unsubscribe" within the message body. **Outlook Unsubscribe** is a lightweight Outlook web add-in that adds an **Unsubscribe** command to received email messages. It reads the standard mailing-list unsubscribe headers in the current message and uses the best unsubscribe method advertised by the sender.

The add-in supports **new Outlook**, **classic Outlook**, **Outlook on the web (OWA)**,
and Outlook on **iOS and Android**. Desktop and web clients use a ribbon
command; mobile clients use an Outlook mobile command that opens a compact task
pane. My plan is to eventually make this a free giveaway on Microsoft Marketplace.

## Table of Contents
* [FREQUENTLY ASKED QUESTIONS](#FREQUENTLY-ASKED-QUESTIONS)
* [INSTALLATION - Single User](#INSTALLATION-for-Single-User)
* [INSTALLATION - M365 Tenant-wide](#INSTALLATION-for-M365-Tenant)
* [REVISION HISTORY](#REVISION-HISTORY)
* [License & Privacy Policy](#License) (MIT)
* Source Code: [https://github.com/awsles/unsubscribe/](https://github.com/awsles/unsubscribe/)

## Cleaning up Unwanted Email
Once you have installed the add-in and you can see it, the next step is to go through existing email. The best way I have found to do this is to SORT your email by SENDER. For each sender, select the most recent email and then click UNSUBSCRIBE. In my case, I then deleted that email from that sender and proceeded to the next. You may have different message types from the same sender, so sometimes a couple of targeted UNSUBSCRIBEs are needed. Be careful not to unsubscribe from important email (like notices from your bank). Generally, such important email does not have an unsubscribe option but sometimes it does. And if you're on a roll cleaning up, you may inadvertently unsubscribe from something you want to keep.

Just click the **Unsubscribe** button when reading the email message:
![One-Click Unsubscribe](screenshots/UnsubscribeButton2.png)

The one-click **Unsubscribe** is then confirmed:
![One-Click Submitted](screenshots/One-ClickSubmitted.png)

In the event that there is no hidden RFC-8058 compliant header, the add-in will attempt to locate an unsubscribe link within the message body. This will open a browser page to the associated unsubscribe page.

### Update v1.0.0.13
v1.0.0.13 introduces a change whereby if you click Unsubscribe a second time, it will bring up a dialog asking you which unsubscribe method you would like to use.
This dialog only appears if there is there is both an RFC-8058 compliant option AND a detectable embedded unsubscribe link within the message body.

![Unsubscribe Method](screenshots/ChooseMethodSmall.png)

---

# INSTALLATION for Single User
The add-in is currently installed by *sideloading* a Manifest file (`manifest.xml`) file into Outlook.
This is done by using the WEB version of Outlook (often called OWA for Outlook Web Access). It can take a few hours for the plug-in to propagate to your desktop and mobile Outlook (though often it is a couple of minutes if you fully close and re-open Outlook).

## 1. Download `manifest.xml`
Download the current manifest from the project repository in GitHub:

https://github.com/awsles/unsubscribe/blob/main/manifest.xml

Click the link to open the file and choose **Download raw file**, then save it somewhere convenient on your PC, such as your **Downloads** folder.

You can also view the deployed copy here:
https://awsles.github.io/unsubscribe/manifest.xml

Outlook requires a local XML file for manual sideloading; Microsoft's current sideloading interface does not offer the **Add from URL** option for individual use.

## 2. Open the Outlook add-in sideloading page
In a web browser, open:

https://aka.ms/olksideload

Sign in with the **same Microsoft account/mailbox that you use in Outlook**.
Give the page a moment to fully load the "Add-Ins for Outlook" popup window. 

The **Add-Ins for Outlook** dialog should open.

## 3. Open My add-ins
In the Add-Ins for Outlook dialog:

1. Select **My add-ins**.
2. Scroll to the **Custom Addins** section near the bottom.
3. Select **Add a custom add-in**.
4. Select **Add from File**.

## 4. Select the manifest
Choose the `manifest.xml` file you downloaded in Step 1.

Accept the installation/security prompts.

The add-in is now associated with your Outlook mailbox. Microsoft states that an add-in sideloaded through this dialog should also become available in supported Outlook desktop clients that use the same mailbox.

> **Classic Outlook note:** A manually sideloaded add-in can occasionally take up to 24 hours to appear in classic Outlook because of client caching. New Outlook and Outlook on the web normally reflect the installation much sooner.

## 5. Find the Unsubscribe command
Open a **received email message** in Outlook, based on the version of Outlook you are using:

### New Outlook for Windows / Outlook on the web

1. Select the message.
2. In the message action bar, select **Apps**.
3. Look for **Outlook Unsubscribe** or **Unsubscribe**.

You may be able to pin the command depending on your Outlook configuration.

### Classic Outlook for Windows
Open a received message and look for the **Unsubscribe** command on the message's add-in/ribbon surface. If it does not appear immediately after installation, restart Outlook and allow time for its add-in cache to refresh.

## 6. Test it
For the best test, open a legitimate newsletter or mailing-list message that provides a `List-Unsubscribe` header (i.e., look for some junk mail).

Select **Unsubscribe**.

Depending on the message, the add-in will either:

- Submit an RFC 8058 one-click unsubscribe request
- Open the sender's unsubscribe webpage
- Open a pre-addressed unsubscribe email draft, OR
- Display a message saying that no supported unsubscribe method was found.

## Updating the add-in

Most JavaScript/HTML changes are loaded from the hosted GitHub Pages site automatically. However, if the `manifest.xml` itself changes, Outlook may continue using the older copy that was previously sideloaded.

When the manifest version changes:

1. Download the new `manifest.xml`.
2. Remove the existing custom add-in from **My add-ins** if necessary.
3. Sideload the new manifest using **Add from File** again.

The manifest's `<Version>` value can be used to verify that you have the expected release.

## Removing the add-in

To uninstall the sideloaded add-in:

1. Open https://aka.ms/olksideload 
2. Select **My add-ins**.
3. Locate **Outlook Unsubscribe** under your custom add-ins.
4. Remove the add-in.


---

# INSTALLATION for M365 Tenant
You can deploy an Outlook add-in to all users in your Microsoft 365 tenant using the Microsoft 365 Admin Center. This is called Centralized Deployment and allows administrators to install and manage add-ins across the organization without manually installing them on each user's computer. You will need Exchange Administrator or Global Administrator privileges to do this.

## Steps
1. Sign in to the [Microsoft 365 Admin Center](https://admin.microsoft.com/).
1. Navigate to **[Settings → Integrated apps](https://admin.cloud.microsoft/#/Settings/IntegratedApps)**.
1. Scroll down below the *Popular apps to be deployed* section and click **Upload custom apps**.
1. Select the **App type** to be "Office Add-in" in the drop-down menu.
1. Choose **Provide link to manifest file** and enter: `https://awsles.github.io/unsubscribe/manifest.xml`
1. Click the **Validate** button to verify the manifest. This may take a couple of minutes.
1. Click **Next**
1. Under user assignment, select Everyone to deploy the add-in tenant-wide.
1. Click Deploy and complete the wizard.
1. You can return to the **Integrated apps** menu and see "Outlook Unsubscribe" in the list of apps. The context menu allows you to change the users to which the plug-in is made available.

![Tenant Install](screenshots/M365%20Tenant%20Install.png)

## Verify Installation
1. Open Outlook (desktop or web or mobile).
2. Open a received email.
3. Look for Outlook Unsubscribe in the Outlook toolbar or Apps menu.
NOTE: Deployment may take 24–72 hours to appear for all users. Restarting Outlook may help.

Reference: [Microsoft's official add-in deployment instructions](https://learn.microsoft.com/en-us/microsoft-365/admin/manage/manage-deployment-of-add-ins?view=o365-worldwide).

## uninstall
The add-in can be removed from the **[Settings → Integrated apps](https://admin.cloud.microsoft/#/Settings/IntegratedApps)** screen.


---

# FREQUENTLY ASKED QUESTIONS

## I don't see the Unsubscribe button?
Once you've installed the add-in, you should refresh your Outlook web.
The Outlook Unsubscribe only appears when you are reading an email.
If you don't see the icon, then look for an icon that looks like a stove top with 4 burners (basically a square with 4 small circles within).
Click that and you should then see Outlook Unsubscribe.

## Outlook is prompting "Outlook Unsubscribe wants to display a new window"?
This is normal behavior in cases where Outlook needs to open a browser window to take you to the respective unsubscribe page. Unfortunately, there is no way to bypass this prompt - though I have found that Outlook stops asking after a while (giving me the impression that because I approved enough times, it is okay).

## How does it work?
Email originating via a subscription often includes a hidden **List-Unsubscribe** header that contains a link for unsubscribing from the associated subscription. This may be in addition to a *Click _here_ to unsubscribe* which appears in the email body.

This Outlook add-in provides an **Unsubscribe** button that appears when reading email in Outlook. When clicked, this button
looks for the List-Unsubscribe header and, if present, will launch a mini-browser to that web page. In some cases, the Unsubscribe
does not display a page but simply just accepts the unsubscribe action.

You should only have to do this for each unique sender & subscription that you want to unsubscribe from.
IT is best to choose the most recent email from that sender to use the unsubscribe button in.

For more detail, see [*How it Works*](How%20it%20Works.md).

## How long does it take to be unsubscribed?
Often, the unsubscribe takes effect immediately. In other cases, it can take days for the sender to honor the request.

## Does it work with GMAIL?
This UNSUBSCRIBE plug-in only works with Microsoft 365 personal and work accounts. It does not work with GMAIL. 

## It did not find the Unsubscribe in the email but I can see it?
If there is no one-click unsubscribe option in the email headers, then the add-in attempts to find an unsubscribe option within the body of the message. Often, this is obvious. Sometimes however, it is not. For example, if the text is "Remove me" instead of "unsubscribe". I will periodically update the plug-in to reflect different use cases (you may need to manually update the plug-in to get the latest version if the manifest.xml changes). 


---

# LICENSE
Outlook Unsubscribe is free and open-source software distributed under the MIT License. You may use, modify, and redistribute the source code for any purpose, including commercial use, provided that the original copyright and license notices are retained. See [LICENSE.md](LICENSE.md) for the full license.


## DISCLAIMER OF WARRANTY AND LIMITATION OF LIABILITY

**THIS SOFTWARE IS PROVIDED FREE OF CHARGE, "AS IS" AND "AS AVAILABLE," ENTIRELY AT YOUR OWN RISK.**

This Outlook Unsubscribe Add-in is open-source software provided without warranties of any kind, express or implied, including warranties of merchantability, fitness for a particular purpose, security, reliability, or non-infringement.

**No technical support, maintenance, updates, or continued availability are guaranteed.** The software may contain errors, malfunction, or produce unintended results. In particular, successful unsubscription is not guaranteed, and use of the software may result in unintended email transmissions, disclosure of information to third parties, or other adverse consequences.

**TO THE MAXIMUM EXTENT PERMITTED BY APPLICABLE LAW**, the authors, contributors, and copyright holders shall not be liable for any claims, damages, losses, or liabilities of any kind, whether direct, indirect, incidental, consequential, or otherwise, arising from the installation, use, misuse, or inability to use this software, even if advised of the possibility of such damages.

**BY INSTALLING OR USING THIS SOFTWARE, YOU ACKNOWLEDGE THAT YOU DO SO ENTIRELY AT YOUR OWN RISK AND ASSUME FULL RESPONSIBILITY FOR ANY CONSEQUENCES.**

## PRIVACY POLICY
This add-in does not intentionally collect or transmit any personal data beyond what the email sender has included in their unsubscribe link(s). This data may include your email address and/or other identifiers required by the sender to unsubscribe you. All Javascript processing occurs client side. 

Questions: Please refer any questions to the Unsubscribe project on GitHub.


# REVISION HISTORY

- v1.0.0.4 -- *Initial Version*
- v1.0.0.5 -- Add support for extracting Unsubscribe link from message body
- v1.0.0.6 -- Add promptBeforeOpen:false in Office.context.ui.displayDialogAsync() to eliminate "Outlook Unsubscribe wants to display a new window."
- v1.0.0.7 -- Rolled back v1.0.0.6 changes
- v1.0.0.8 -- Standardize mailto unsubscribe drafts and identify the original recipient in the body
- v1.0.0.9 -- Wait for Outlook to open mailto drafts before completing the command
- v1.0.0.10 -- Add Outlook mobile command surface and mobile unsubscribe task pane
- v1.0.0.11 -- Detect unsubscribe wording immediately before or after a hyperlink
- v1.0.0.13 -- Adds a second-click unsubscribe chooser to select One-Click or embedded link when both are available after the first click

## Known Limitations & Issues

1. There are variations of "unsubscribe" options appearing in emails. This add-in uses a *best effort* approach to find these which doesn't always succeed. The add-in will be periodically updated to support new use cases (this will require you to install the latest version to get those cases).
1. Outlook does not load the primary icon for the add-in in some places.  Unsure why.
1. Users have to accept "Outlook Unsubscribe wants to display a new window." in some cases. If promptBeforeOpen:false is added to Office.context.ui.displayDialogAsync(), then the window open fails.  Not sure if there is another work-around.
1. Extraction of embedded Unsubscribe links is based on keywords, so it can miss some.
1. On mobile clients that don't expose `getAllInternetHeadersAsync()`, header-only `List-Unsubscribe` methods aren't available; the add-in scans the message body instead.

## Feedback
Feedback is welcomed! Raise an issue and/or contribute to the code base. In particular, I am looking for use cases where the unsubscribe link embedded in the message body was not discovered (please include the relevant text surrounding the link - preferably in HTML form when raising an issue).

The published version of this README is [here](https://awsles.github.io/unsubscribe/).
See [this article](https://certified-senders.org/wp-content/uploads/2017/07/CSA_one-click_list-unsubscribe.pdf) about RFC-8058 List-Unsubscribe header.
