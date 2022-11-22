# CTSinator
AHK automated helping tool for CTS team for best results use chrome or change all instances of chrome.exe in code to your browser of choice

functions:

ctrl+windows+alt+x - panic button, press to stop script

ctrl+windows+alt+g - message generator ran from a ticket main page in snow - incident or catalog task generate a message based on a ticket in service-now, works for me, might be buggy but if ironed out saves huge amounts of time

ctrl+windows+alt+h - ticket holder same as above, used for putting tickets on hold , open text box allows for options:
u - hold for user unavailable with message "Contacted user on Teams, waiting for reply"

ctrl+windows+alt+f - finder, the most used by me - use instead of ctrl+c on a string you want to search for in snow. opens chrome, works only if you're logged in to snow in another tab

ctrl+windows+alt+v - paste the message from the generator

ctrl+windows+alt+shift+v - paste the message from the generator into clipboard

ctrl+windows+alt+e - paste the email address from the message generator

ctrl+windows+alt+c - type in "Contacted user on Teams, waiting for reply"

ctrl+windows+alt+i - type in "my computer info" instruction

ctrl+windows+alt+t - type in "user stated issue has been resolved already, closing ticket"

ctrl+windows+alt+s - type in "ok, I will be connecting shortly"

ctrl+windows+alt+b - type in "I'll close the ticket now, have a great day :)"

ctrl+windows+alt+k - type in "user has confirmed - software installed successfully"

ctrl+windows+alt+a - action menu type in options:

laps - inputboxes for hostname and DC, runs powershell and finds local admin password

dcu - puts a note on how to update drivers using dell comand update into your clipbaord, paste in chat

cmrc - runs config manager remote control viewer, copy a hostname before use
