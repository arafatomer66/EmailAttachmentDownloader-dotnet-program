# Outlook Email Processing Program

# Commands
# Step 1: Build the project
dotnet build

# Step 2: Run the project
dotnet run

## Overview

This program checks and downloads emails from the Outlook inbox at regular intervals (every 30 minutes). It specifically looks for emails from `your_email` with subjects containing "Email" and saves up to two attachments from each matching email to the local `you directory` directory. The program logs all activities and errors into a `log.txt` file.

## Prerequisites

- Microsoft Outlook installed on your machine.
- Microsoft.Office.Interop.Outlook library.
- .NET Framework.

## Setup

1. **Install Microsoft Office Interop for Outlook**:
   - You can install the `Microsoft.Office.Interop.Outlook` package via NuGet:
     ```sh
     Install-Package Microsoft.Office.Interop.Outlook
     ```

2. **Ensure the `C:\Network` directory exists**:
   - Create the `C:\Network` directory if it doesn't already exist to save attachments.
     ```sh
     mkdir C:\Network
     ```

3. **Grant necessary permissions**:
   - Ensure the application has the required permissions to access and manipulate Outlook emails and the file system.

## Running the Program

1. **Build and run the program**:
   - Use your preferred IDE or command line to compile and run the program.

2. **Logs**:
   - The program logs all its activities and errors to a file named `log.txt` in the working directory. You can monitor this file for any issues or information about the program's operations.

## Code Explanation

### Main Method
- The `Main` method initializes the logging and sets up an infinite loop that runs every 30 minutes to check and process emails.

### CheckAndDownloadEmails Method
- Initializes Outlook application and accesses the inbox folder.
- Synchronizes the inbox to get the latest emails.
- Iterates through all emails, checking for specific criteria (subject contains "Email" and sender is `arafatomer66@gmail.com`).
- Saves up to two attachments from matching emails to the `your directory` directory.your_email

### Log Method
- Writes log messages to both the console and a `log.txt` file for debugging and monitoring purposes.

## Handling Errors

- Any errors encountered during the email checking and attachment downloading process are logged with detailed messages to help identify and fix issues.

## Customization

- **Email Criteria**: Modify the conditions in the `CheckAndDownloadEmails` method to change the criteria for processing emails.
- **Attachment Limit**: Change the limit of saved attachments by modifying the condition `if (i >= 2)` in the loop.
- **Log File Location**: Change the `log.txt` path in the `Log` method if you want to save logs to a different location.

## Important Notes

- This program is intended to run in an environment where Outlook is installed and properly configured.
- Ensure that the program has appropriate permissions to interact with Outlook and the file system.
