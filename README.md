# rmm-web-bot
A web script to run PowerShell commands on a list of computers

# Goal of the bot
Our weekly policies that were pushed via our remote management system was not consistent on newer verions of windows, so I developed a bot that can read through a list of devices, imported by the user. The script includes several pop-up windows for the user to enter commands to be executed in PowerShell on the list of devices. The bot then loops through the list of devices, running the entered commands on each.

# Features
I converted this script into an .exe that allows any end-user to use it easily. I added logging to track each step and monitor where the bot currently is and what issues it may be hung up on. Once the bot finishes looping through all devices, an automated email is sent to the user with basic stats of what devices had these commands successfully run, which devices had issues with the commands, and what devices did not exist.
