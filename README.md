# Automated Certificate Generation Using Excel VBA

## Project Overview

This project automates the generation of individual PDF certificates using **Excel VBA, Microsoft Word and PDF export**.

The objective was to eliminate the repetitive manual process of creating certificates for multiple candidates. Instead of manually entering candidate information into a Word certificate template and exporting each certificate individually, the VBA automation handles the complete process in a single run.

## How It Works

Candidate information is maintained in an Excel worksheet with the following columns:

* Candidate ID
* Candidate Name
* Date
* Certificate ID

The Word certificate template contains three placeholders:

* `<<CANDIDATE NAME>>`
* `<<DATE>>`
* `<<CERTIFICATE ID>>`

When the macro is executed, it:

1. Prompts the user to select the Word certificate template.
2. Prompts the user to select the folder where certificates should be saved.
3. Reads candidate information from Excel row by row.
4. Opens a fresh copy of the Word template for each candidate.
5. Replaces the placeholders with the corresponding candidate information.
6. Exports the completed certificate as a PDF.
7. Names the PDF using the Candidate ID.
8. Repeats the process for every candidate in the Excel dataset.
9. Displays a completion summary after processing.

## Automation Workflow

**Excel Candidate Data → VBA → Word Certificate Template → Dynamic Text Replacement → PDF Export → Individual Certificate**

### Example

If the Excel data contains:

`C001 | Rahul Sharma | 23/08/2026 | CERT-001`

The generated certificate will contain:

**Candidate Name:** Rahul Sharma
**Date:** 23 August 2026
**Certificate ID:** CERT-001

The resulting file will automatically be saved as:

`C001.pdf`

## Key Features

* Automated Word template selection
* Automated output folder selection
* Batch certificate generation
* Dynamic replacement of text-box placeholders
* Individual PDF creation
* Automatic file naming using Candidate ID
* Windows filename validation
* Error handling
* Completion summary
* No manual Word editing required

## Technologies Used

* Microsoft Excel
* VBA
* Microsoft Word
* PDF Export

## Business Value

This automation can be particularly useful for organizations that regularly generate:

* Training certificates
* Course completion certificates
* Employee certificates
* Workshop certificates
* Event participation certificates
* Academic certificates
* Recognition certificates

The project demonstrates how **Excel VBA can be used beyond spreadsheets to automate document creation and repetitive administrative workflows**.

#Excel #VBA #Automation #ExcelAutomation #WordAutomation #PDFAutomation #OfficeAutomation #DataAutomation

<img src="https://github.com/irenegoswami20/Automated-Certificate-Generation-Using-Excel-VBA/blob/15ff513d8e7389d897afb55a781d948a6874be80/VBA%20automation%20window.png" width="600"> 
VBA automation
<br>
