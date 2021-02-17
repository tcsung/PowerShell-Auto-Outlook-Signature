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
- %FULLNAME% for users full name
- %JOBTITLE% for users job title
- %PHONE% for users telephone number
- %MOBILE% for users mobile phone number
- %EMAIL% for users email address

Here is an example :
  %FULLNAME%
  %JOBTILE%
  
  %


once done copy the template to your server UNC path "\\<File server name>\netlogon\templates\signatures" 


Users cannot change their signature and it will be locked, if you want to relase that, please remove below registry settings :


path="HKCU\Software\Microsoft\Office\<Version number>\Common\MailSettings"
name=NewSignature

path="HKCU\Software\Microsoft\Office\<Version number\Common\MailSettings"
name=ReplySignature

#  Prerequesties
Prepare the signature template is a must, and save it under "\\<File server name>\netlogon\templates\signatures\<domain name>" or "\\<File server name\netlogon\templates\signatures\default"
