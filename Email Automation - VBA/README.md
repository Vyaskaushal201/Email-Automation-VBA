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

### Excel Data
![Excel](images/excel.png)

### Generated PDF
![PDF](images/pdf.png)

### Email Sent
![Email](images/email.png)

### VBA Code
![Code](images/code.png)

---

## Outcome

- Reduced manual effort by 90%
- Faster communication with customers
- Improved tracking and accuracy
- Scalable for large datasets

---

## Author

Created by [Your Name]



