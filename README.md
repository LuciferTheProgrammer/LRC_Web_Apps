# LRC Web Apps

Applications are contained in the Projects and Carasolva folders.

## Internal IT Automation Web Apps

This repository contains Python and Flask-based automation tools created for internal IT workflows at Living Resources. The applications were built to help automate repetitive account creation tasks, reduce manual data entry, and provide simple web interfaces for running user creation jobs from uploaded spreadsheets.

The main applications included in this repository are an Active Directory user creation web app and a Carasolva user creation web app. Both applications use spreadsheet data as input and automate user setup tasks that would normally require manual entry.

## Features

* Flask web application interface
* Spreadsheet upload support
* Supports `.xlsx`, `.xls`, and `.csv` files
* Active Directory administrator credential validation
* LDAP connection and bind testing
* Target OU lookup by distinguished name or OU name
* Background worker process for user creation jobs
* Live job log streaming in the browser
* Active Directory user account creation
* User attribute population from spreadsheet data
* Supervisor lookup support
* Default Active Directory group assignment
* PowerShell post-processing
* Microsoft 365 Office 365 E1 license assignment script generation
* ZIP output bundle generation
* Carasolva user creation web app
* Carasolva browser-based automation workflow
* Selenium backend automation for Carasolva
* Microsoft Edge WebDriver support
* Duplicate user checking
* Error logging and screenshot capture
* Web interface branding with the Living Resources logo

## Project Files

* `app.py` - main Flask web application for the Active Directory user creation tool
* `worker.py` - backend worker script that processes uploaded spreadsheets and creates Active Directory users
* `index.html` - main web page for uploading spreadsheets and entering administrator credentials
* `job.html` - live job console page that streams output while an Active Directory job is running
* `App.py` - Flask web application for running the Carasolva user creation automation from a browser
* `Carasolva UserCreation.py` - Selenium backend automation script used by the Carasolva Flask web app
* `livingresources-logo.png` - logo used in the web application interface
* `summary.xlsx` - generated output file containing processed user information
* `AssignOffice365E1.ps1` - generated PowerShell script used to assign Office 365 E1 licenses
* `post_process.ps1` - generated PowerShell script used for Active Directory post-processing
* `created_users.json` - generated file containing newly created Active Directory user records
* `run_log.txt` - generated log file for reviewing job output
* `ERROR.txt` - generated error file if the worker fails
* `WORKER_STDERR.txt` - generated file containing backend worker error output when applicable

## Active Directory User Creation Web App

The Active Directory user creation web app allows an IT administrator to upload a spreadsheet of new users and create accounts in Active Directory through a browser-based interface.

The app validates the administrator credentials, verifies the target OU, saves the uploaded spreadsheet, starts a background worker process, streams live logs to the browser, and provides a downloadable ZIP file after the job completes.

## Active Directory App Features

* Upload a spreadsheet of new users
* Validate AD administrator credentials
* Accept usernames with or without a domain prefix
* Resolve the target OU by full distinguished name or OU name
* Create users in a temporary target OU
* Generate usernames and email addresses from first and last names
* Set user attributes such as title, department, location, employee ID, and supervisor
* Force password change on next login
* Enable created accounts
* Add users to default Active Directory groups
* Run PowerShell post-processing after account creation
* Move users to the configured destination OU
* Set mail-related Active Directory attributes
* Create a summary spreadsheet of processed users
* Generate a PowerShell script for assigning Office 365 E1 licenses
* Provide a downloadable ZIP bundle of output files

## Required Active Directory Spreadsheet Columns

The Active Directory user creation app expects the uploaded spreadsheet to contain the following columns:

* `First Name`
* `Last Name`
* `Employee ID`
* `Job Title`
* `Department`
* `Location`
* `Supervisor`

The spreadsheet may also include the following optional column:

* `Supervisor Username`

The `Supervisor Username` column can help the script find the correct supervisor account more accurately. If this column is not provided, the script attempts to resolve the supervisor from the `Supervisor` name field.

## How the Active Directory App Works

The user opens the Flask web app and uploads a spreadsheet containing employee information. The administrator enters their Active Directory admin username and password, along with the target OU where the accounts should initially be created.

The Flask app checks the administrator credentials against the configured domain controller. It then verifies the target OU and creates a background job for the uploaded spreadsheet. The browser is redirected to a live job console page, where logs are streamed in real time using Server-Sent Events.

The backend worker script reads the spreadsheet and processes each row. For every user, the worker builds the account username, email address, and display name. It checks whether the user already exists, creates the account if needed, fills in the required attributes, assigns default groups, and adds the user to the summary output.

After the user creation process finishes, the worker runs PowerShell post-processing. This step updates additional Active Directory attributes, sets mail-related fields, moves newly created users to the destination OU, and prepares the generated Office 365 E1 licensing script.

When the job is complete, the web app creates a ZIP file containing the summary spreadsheet, logs, generated PowerShell scripts, and other output files for review.

## Active Directory Generated Output Files

The Active Directory app can generate the following output files:

* `summary.xlsx` - spreadsheet containing processed users and email addresses. This file can later be used for the Carasolva automated account creation process.
* `run_log.txt` - log file containing the job output
* `ERROR.txt` - error file created if the job fails
* `WORKER_STDERR.txt` - backend worker error output
* `created_users.json` - JSON file containing newly created user records
* `post_process.ps1` - PowerShell script used for post-processing
* `AssignOffice365E1.ps1` - PowerShell script used to assign Office 365 E1 licenses
* `README.txt` - short explanation of the generated ZIP bundle contents

## Carasolva User Creation Web App

The Carasolva user creation web app allows an IT user to upload a spreadsheet, enter Carasolva login credentials, choose a Carasolva role, provide the Microsoft Edge WebDriver path, and start the user creation automation from a browser.

The web app uses Flask for the front end and backend routing. It saves the uploaded spreadsheet, starts the Selenium automation script in the background, streams live output to the browser, and displays a summary when the job finishes.

## Carasolva Web App Features

* Browser-based Flask web interface
* Upload a spreadsheet of users
* Enter Carasolva username and password through the web form
* Enter or confirm the Carasolva role to assign
* Enter or confirm the Microsoft Edge WebDriver path
* Start the automation from the browser
* Run Selenium automation in the background
* Stream live output back to the browser
* Read user data from `.xlsx`, `.xls`, or `.csv` files
* Supports flexible spreadsheet column headers
* Uses Microsoft Edge WebDriver
* Logs into the Carasolva MedSupport site
* Navigates to the Users section
* Searches for existing users before creating new accounts
* Skips duplicate users
* Fills in first name, last name, initials, username, employee number, and title
* Sets a default password
* Assigns the selected Carasolva role
* Saves each new user account
* Returns to the user list after each user is processed
* Saves screenshots when an error occurs
* Displays final user summary in the browser

## Required Carasolva Spreadsheet Columns

The Carasolva web app expects the uploaded spreadsheet to contain at least the following columns:

* `First Name`
* `Last Name`
* `Email`

The spreadsheet may also include the following optional column:

* `Employee ID`

If an Employee ID is included, it is used for the SSO Login and Employee Number fields in Carasolva.

## How the Carasolva Web App Works

The user opens the Carasolva Flask web app and uploads a spreadsheet containing the users that need to be created in Carasolva. The user enters their Carasolva username and password, chooses the role to assign, confirms the Microsoft Edge WebDriver path, and starts the automation from the browser.

The Flask app saves the uploaded spreadsheet into an upload folder for that run. It reads the spreadsheet information for the final summary and creates a unique run ID. The app then starts `Carasolva UserCreation.py` as a background process and streams the script output back to the browser using Server-Sent Events.

The Selenium backend automation opens Microsoft Edge, logs into the Carasolva MedSupport website, expands the company section, opens the Users page, and processes each user from the spreadsheet.

For each user, the automation searches for an existing record by first name, last name, and email address. If the user already exists, the automation skips that user and returns to the user list. If no existing record is found, the automation continues with account creation. It fills in the required fields, sets the username, enters employee information if available, assigns the selected role, saves the user, and moves on to the next record.

If an error occurs while processing a user, the automation saves a screenshot for troubleshooting and then attempts to return to the user list.

When the automation finishes, the web app displays a summary of the users from the uploaded spreadsheet.

## Technologies Used

* Python
* Flask
* pandas
* openpyxl
* Selenium
* Microsoft Edge WebDriver
* ldap3
* pyad
* PowerShell
* Microsoft Graph PowerShell
* HTML
* CSS
* JavaScript
* Server-Sent Events
* Active Directory
* Microsoft 365
* Carasolva MedSupport

## Running the Active Directory Web App

Install the required Python packages:

```bash
pip install flask pandas openpyxl ldap3 pyad
```

Run the Active Directory Flask app:

```bash
python app.py
```

Open the app in a browser:

```text
http://localhost:5050
```

If the app is running on a server or another machine on the network, use the server IP address:

```text
http://SERVER-IP:5050
```

## Running the Carasolva Web App

Install the required Python packages:

```bash
pip install flask pandas openpyxl selenium
```

Make sure Microsoft Edge WebDriver is installed on the machine running the app. The Edge WebDriver path can be entered in the web form or configured in the Flask app.

Run the Carasolva Flask web app:

```bash
python App.py
```

Open the app in a browser:

```text
http://localhost:5000
```

If the app is running on a server or another machine on the network, use the server IP address:

```text
http://SERVER-IP:5000
```

## Carasolva Web App Workflow

1. Start the Flask app with `python App.py`
2. Open `http://localhost:5000`
3. Enter the Carasolva username and password
4. Upload the spreadsheet
5. Enter the Carasolva role, such as `Non Med Cert Staff`
6. Confirm or update the Microsoft Edge WebDriver path
7. Click `Start Script`
8. Review the live output in the browser
9. Review the final summary after the automation finishes

## Recommended Repository Structure

```text
LRC_Web_Apps/
│
├── Projects/
│   ├── app.py
│   ├── worker.py
│   ├── templates/
│   │   ├── index.html
│   │   └── job.html
│   └── static/
│       └── livingresources-logo.png
│
├── Carasolva/
│   ├── App.py
│   ├── Carasolva UserCreation.py
│   ├── templates/
│   │   └── index.html
│   └── static/
│       └── livingresources-logo.png
│
├── README.md
└── requirements.txt
```

Depending on the final project organization, the Active Directory web app and the Carasolva web app may be separated into their own folders.

## Files That Should Not Be Committed

Generated files, spreadsheets, logs, screenshots, and sensitive data should not be committed to the repository.

Recommended files and folders to exclude from Git:

```text
uploads/
*.xlsx
*.xls
*.csv
*.zip
summary.xlsx
run_log.txt
ERROR.txt
WORKER_STDERR.txt
created_users.json
post_process.ps1
AssignOffice365E1.ps1
carasolva_errors/
__pycache__/
.env
```

## Security Notes

These tools are intended for internal IT use only.

Do not commit real passwords, administrator credentials, employee spreadsheets, production logs, downloaded ZIP bundles, generated output files, or screenshots containing sensitive information.

Administrator passwords and Carasolva passwords are entered at runtime through the web interfaces or command line and should not be saved in the repository.

The Flask web apps should be run only in a trusted internal environment. If they are deployed for regular production use, HTTPS and additional access controls should be used.

## Purpose

This project was created to support internal IT account creation workflows at Living Resources.

The Active Directory user creation web app helps automate new user setup by creating AD accounts from spreadsheet data, assigning standard attributes and groups, generating summaries, and preparing follow-up Microsoft 365 licensing scripts.

The Carasolva user creation web app helps reduce manual data entry by allowing users to upload a spreadsheet and run Carasolva account creation automation from a browser.

Together, these tools help make user creation faster, more consistent, easier to track, and easier to troubleshoot.
