# email-order-scanner
Python program for scanning email for customer orders

## Description  
This is a very simple hacky solution to a local business I am helping out. They run a restaurant and have 3rd party integrations just fine, with one exception - Dine In online, a local 3rd party delivery service. This service provides a tablet for orders, and orders also come to an email inbox, but are unable to come through their POS with any embedded integrations. For years, this business has been manually printing orders from the email that are placed through this service. This process is cumbersome and caused orders to be late on occasion or caused an inconvenience for employees to have to print orders this way. This hacky solution is intended to very simply solve this issue as fast as possible. I have been running this without issue on their system for one year. The only problem is the requirement to reauthorize the user token roughly once every week, but this is a limitation of the Google GMail API.

This program is split into two as detailed below, but could optionally be restructured to function together.

## What is inside
1. gmail printer  
The purpose of the gmail printer is to scan a gmail inbox looking for emails with a specific subject string and print the attached pdf file. It should also save the pdf file to a folder using the order number as the file name.

2. Fetch all dine-in orders
The purpose of the fetch all is to scan only the main inbox and save any old orders that are still present and move them to a separate folder for organizational purposes.