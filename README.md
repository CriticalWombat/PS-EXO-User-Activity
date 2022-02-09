# PS-EXO-User-Activity
PowerShell Exchange Online User Activity report

This script is designed to run against an office365 tenant with admin credentials. This generates a report in excel with 3 columns being:
Display name, Last logon time(to the mailbox), and last interaction time(LastInteractionTime does not always work correctly and gives bogus dates...).
