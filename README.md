# Proof of Delivery ETL Automation
## Uploaded to GitHub: 8/22/2024

This was a Python project that I was working on for a previous employer that had the intent of automating the workflow for a particularly monotonous task. 

In the US, in the home medical equipment business, it is necessary for the companies that are delivering equipment to prove that their deliveries have 
been completed before they can be reimbursed by insurance companies. This code sought to automate that process in a variety of ways:

- Download a report of internal customer order information and transform it into a tidy DataFrame 
- Download hundreds of PDFs of receipts from a third-party website read all of those PDFs, and turn them into a tidy DataFrame
- Merge those spreadsheets together to get a list of all orders that have been officially delivered
- Complete internal order processing steps by automating input into company software

I leveraged a variety of libraries for this project, including but not limited to:

- selenium
- selenium-stealth
- pywin32
- pyautogui
- pytesseract
- pydirectinput
- pandas
- and more...

Currently, when the code is supplied with the appropriate links and file path constant variables, it works for the first three steps of the above mentioned
process. I ran into issues with the fourth step, seemingly because I was trying to use pyautogui/pydirectinput to manipulate the processes of a virtual machine.

Anywho, I no longer work for this company and don't have access to the appropriate links, so this will just serve as a reference for future web automation projects. 
