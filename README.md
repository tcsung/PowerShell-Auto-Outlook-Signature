# Auto Outlook Signature (PowerShell Edition)

## Preface

This script written since 2015 by using Kixtart which I used to standardize all users Outlook signature, here is what I converted to PowerShell and share it out.  Script enquire the AD user account information to pull out the data, it generated a standard signature from your template and then update the Outlook if any different. Script also can auto enable the signature in the Outlook, but due to registry update contain some delay so the email account is first time setup then user must logout and login the computer after the script run.

## Compatibility

Script is compatibile with MS Outlook 2010 to the latest version, as long as computer is running Windows 7, 8 or 10 environment.


## Requirement

- Knowledge of how Distinguished names (DN) work on Active Directory
- Basic knowledge of Active Directory
- Create your Email signature template from your Outlook.
- netlogon folder shared for script use purpose.

## How to

Script require to create a signature template, here are the guideline how to create your template :
1. Open your Outlook, Click 'File' > 'Options', Under the pop up screen click on 'Mail', in the right side Windows select 'Signatures...'
2. Under the 'Signature and Stationery' screen, assume it has nothing defined before (better nothing, otherwise remove it all), click on 'New' button, type the name of the signature call 'sign' (This is a must).
3. Start draft your signature template under 'Edit signature', and put below parameters instead of the ture info :
- %FULLNAME% for user's full name
- %JOBTITLE% for user's job title
- %PHONE% for user's telephone number
- %MOBILE% for user's mobile phone number
- %EMAIL% for user's email address
- %COMPANYNAME% for company name

Here is an example:

	%FULLNAME%
	%JOBTITLE%
	
	%COMPANYNAME%
	Tel :		%PHONE%
	Email :		%EMAIL%
	Mobile :	%MOBILE%
	www.abc.com


4. Once done above, click 'OK' to save it, you should be able to find the signature saved under 'C:\Users\<username>\AppData\Roaming\Microsoft\Signatures', rename the 'Signatures' folder to 'default' and copy/move it to your server UNC path "\\<File server name>\netlogon\templates\signatures"
5. If you have signature setup for different domain, you rename as the 'Signatures' folder to your domain name.
6. Try run the script on the computer then you should be able to get the standard form of signature ready

### How to update templates:

Template contain 3 files:
- sign.htm use Microsoft Word
- sign.rtf use WordPad
- sign.txt use Notepad

Anything you updated on either of the files, will auto reflect to the email signature when next time run the script.


## Remarks

You probably found that users cannot change their signature and it will be locked, if you want to release it, please remove below registry settings :

	path = 'HKCU\Software\Microsoft\Office\<Version number>\Common\MailSettings'	
	name = 'NewSignature'

	path = 'HKCU\Software\Microsoft\Office\<Version number\Common\MailSettings'
	name = 'ReplySignature'

I can provide an update version of script for create exception list for ignore the registry changes, or you just remove related codes from the script.

## Sample template

Please also check the sample signature template for information.
