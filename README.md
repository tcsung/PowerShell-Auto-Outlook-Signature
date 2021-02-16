# PowerShell-Auto-Outlook-Signature
Make your Outlook signature fully auto deploy to the computers

# Feature

This script is to help you automate the Outlook signature for each computers in your domain, the feature is almost the same as  raymix /
PowerShell-Outlook-Signatures, script enqire the AD about the user information and generate the signature,the different is to auto apply the change for each computer by change it's registry settings.  During run the script, better close
the Outlook, but you may find that the it may not reflect to your Outlook immediately but after restart the computer it should be there.

Users cannot change their signature and it will be locked, if you want to relase that, please remove below registry settings :


path="HKCU\Software\Microsoft\Office\<Version number>\Common\MailSettings"
name=NewSignature

path="HKCU\Software\Microsoft\Office\<Version number\Common\MailSettings"
name=ReplySignature

#  Prerequesties
Prepare the signature template is a must, and save it under "\\<File server name>\netlogon\templates\signatures\<domain name>" or "\\<File server name\netlogon\templates\signatures\default"
