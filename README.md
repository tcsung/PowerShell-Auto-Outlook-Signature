# Auto Outlook Signature (PowerShell Edition)

## Preface

To make all your Outlook computer has a standard email signature auto setup, this script was written since 2015 by using Kixtart, here is what I converted to PowerShell and share it out.  Script based on the AD user account information to pull out the data, it generate a signature from your template and then update the user Outlook if any different. Script can auto apply to your Outlook, but due to registry update contain some delay so user must logout and login the computer to get the signature auto attached.

## Compatibility

Script is compatibile with MS Outlook 2010 till the latest version, as long as computer is running Windows base environment.



Users cannot change their signature and it will be locked, if you want to relase that, please remove below registry settings :


path="HKCU\Software\Microsoft\Office\<Version number>\Common\MailSettings"
name=NewSignature

path="HKCU\Software\Microsoft\Office\<Version number\Common\MailSettings"
name=ReplySignature

#  Prerequesties
Prepare the signature template is a must, and save it under "\\<File server name>\netlogon\templates\signatures\<domain name>" or "\\<File server name\netlogon\templates\signatures\default"
