# PAPER (PAP Email Responder)
Author: Jan Janiczek

A set of custom tools to simplify work specific to my team:

1. A "moveIT" autoresponder - searches the team's mailbox for "moveIT" requests (auto-generated by a different system operated by another team), parses them to extract the required data (since the emails are auto-generated from the same template every single time, the position of said data is more predictable than its content, so it would not make sense to use regular expressions), creates a new email by combining this data with a template, and dispatches it on-behalf-of the group mailbox. The need to send it on-behalf-of in a tightly controlled corporate network necessitates the use of Microsoft Outlook as an intermediary.

2. An inbox cleaner moving auto-generated email typical for our work (it has to be stored for audit) to a folder dedicated for it.

3. A simple generator of pseudo-random passwords.

Originally compiled using pyinstaller, with the command: "pyinstaller --onefile --windowed PAPER.py"
