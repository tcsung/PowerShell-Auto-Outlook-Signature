# Auto Outlook Signature (PowerShell Edition)

## Preface

This script written since 2015 by using Kixtart which I used to standardize all users Outlook signature, here is what I converted to PowerShell and share it out.  Script enquire the AD user account information to pull out the data, it generated a standard signature from your template and then update the Outlook if any different. Script also can auto enable the signature in the Outlook, but due to registry update contain some delay so the email account is first time setup then user must logout and login the computer after the script run.

## Compatibility

Script is compatibile with MS Outlook 2010 till the latest version, as long as computer is running Windows 7, 8 or 10 environment.





Users cannot change their signature and it will be locked, if you want to relase that, please remove below registry settings :


path="HKCU\Software\Microsoft\Office\<Version number>\Common\MailSettings"
name=NewSignature

path="HKCU\Software\Microsoft\Office\<Version number\Common\MailSettings"
name=ReplySignature

#  Prerequesties
Prepare the signature template is a must, and save it under "\\<File server name>\netlogon\templates\signatures\<domain name>" or "\\<File server name\netlogon\templates\signatures\default"
