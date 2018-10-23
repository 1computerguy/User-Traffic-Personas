# User-Traffic-Personas
User Traffic based on personas in a CSV that outlines what user activities should be performed by each user.

#### Setup Procedures:
  - Download all files to the c:\scripts directory (hard-coded directory path currently used in the script)
  - Extract the user-activity-data-files.7z to the same c:\scripts directory
  - Disable PowerShell script execution limitations (this will vary depending on Windows version)

#### TODO:
  - Make script determine location at run-time and use current working directory instead of hardcoded script location
  - Add "intelligence" to script personas to prevent predictability
    - Consider AI similar to Lariat or other
  - Migrate from PowerShell to C# or some other Microsoft compiled language (to prevent overrun of PowerShell log traffic)
  - Add more realistic data corpus for message (email and document) content
    - Possibly use something like a Reddit chatbot
    - Consider some form of dynamic data
    - Pull/push data from/to file share instead of accessing all data locally in the c:\scripts\documents directory
