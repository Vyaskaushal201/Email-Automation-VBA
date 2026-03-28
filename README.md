# Email Automation System using Excel VBA

## Project Overview

This project automates the process of generating payment reminder letters and sending them via email using Excel VBA and Gmail SMTP.

It eliminates manual work of creating individual letters and sending emails one by one.

---

## Problem Statement

In many organizations, sending payment reminders is a manual and repetitive task:
- Creating individual letters for each customer
- Saving documents manually
- Sending emails one by one
- Tracking which emails are sent

This process is time-consuming and error-prone.

---

## Solution

I developed an automated system using Excel VBA that:

✔ Generates personalized PDF letters  
✔ Stores file paths automatically  
✔ Sends emails with attachments using Gmail SMTP  
✔ Updates email status dynamically  

---

## Tools & Technologies Used

- Microsoft Excel (VBA)
- Microsoft Word (for PDF generation)
- Gmail SMTP (for email automation)

---

## Data Structure

The Excel sheet contains:

- Sr No
- Customer Name
- Email ID
- Amount Pending
- Due Date
- Days Overdue
- Payment Status 
- Email Status 
- File Path

---

## Process Flow

1. Input customer data in Excel  
2. VBA generates personalized letters in Word  
3. Letters are saved as PDF  
4. File path is stored in Excel  
5. Emails are sent automatically with attachment  
6. Email status is updated  

---

## Key Features

- Dynamic PDF generation
- Conditional email sending
- Gmail SMTP integration (no Outlook dependency)
- Error handling (File Missing, Failed)
- Status tracking system

---

## Screenshots

### Excel Data Before Email Sent

<img src="https://github.com/Vyaskaushal201/Email-Automation-VBA/blob/6eb0525b07cf3e45d83e86943fb4be2625d6bdc3/Images/Before%20Email%20Sent.png" alt="Image Description" width="600">

### Excel Data After Email Sent

<img src="https://github.com/Vyaskaushal201/Email-Automation-VBA/blob/09c1d84f7e3ba2bf3f2f74107551dc48d0287749/Images/After%20Email%20Sent.png" alt="Image Description" width="600">

### Generated PDF

<img src="https://github.com/Vyaskaushal201/Email-Automation-VBA/blob/02c9ce7c1bc5cfcdc695c288acb7a3a552451231/Images/PDF%20Latter.png" alt="Image Descriptioon" width="600">

### Email Sent

<img src="https://github.com/Vyaskaushal201/Email-Automation-VBA/blob/b9ceb9d1e9e46a3c1862b5ffc1caaed8d7371b14/Images/Gmail.png" alt="Image Descriptioon" width="600">

### VBA Code For Generates Latters

<img src="https://github.com/Vyaskaushal201/Email-Automation-VBA/blob/2f023886ff9919938422707b1c944abf8438bc73/Images/Generates%20Latters_VBA.png" alt="Image Descriptioon" width="600">

### VBA Code For FillFilePath

<img src="https://github.com/Vyaskaushal201/Email-Automation-VBA/blob/020e445a260cea500573596843ca9c418af15551/Images/FillFilePath_VBA.png" alt="Image Descriptioon" width="600">

### VBA Code For SendEmail

<img src="" alt="Image Descriptioon" width="600">

---

## Outcome

- Reduced manual effort by 90%
- Faster communication with customers
- Improved tracking and accuracy
- Scalable for large datasets

---

## Author

Created by [Your Name]




