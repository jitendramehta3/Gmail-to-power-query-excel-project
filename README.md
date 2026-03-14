# Gmail to Power Query Excel Automation

## Project Overview
This project demonstrates how to automatically fetch email attachment data from Gmail and load it into Excel using Power Query.

Since Power Query cannot directly connect to Gmail, Google Apps Script is used as a bridge to expose Gmail data through a Web API.

## Tools Used
- Microsoft Excel
- Power Query
- Gmail IMAP
- Google Apps Script
- AI

## Project Workflow

1. Enable IMAP access in Gmail
2. Create Google Apps Script to fetch Gmail attachments
3. Deploy the script as a Web App
4. Connect Power Query using:
   Data → Get Data → From Web
5. Convert Base64 data to Binary
6. Load attachments (Excel / CSV / JSON) into tables

## Project Purpose

This project demonstrates how Power Query can be used for automated data extraction from email systems, which is useful in real-world data automation workflows.
