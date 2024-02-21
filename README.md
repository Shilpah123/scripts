# scripts
This repo contains scripts written in different languages and not in Python.

#To execute these powershell scripts, you MUST install AD package. 
Open a command prompt and type powershell. 

C:\Users\sh001\Desktop\script>powershell
PS C:\Users\sh001\Desktop\script>

PS indicates that you are in powershell.

Install AD module as follows:

PS C:\Users\sh001\Desktop\script> Install-Module -Name AzureAD -Scope CurrentUser

After this, run the scripts.

To execute the powershell script, navigate to the location where you have placed your scripts.
PS C:\Users\sh001\Desktop\script>

Then run ./scriptname.ps1

Example:
PS C:\Users\sh001\Desktop\script>./single_manager.ps1
