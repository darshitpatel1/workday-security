# Automate Workday Domain Permissions

A Chrome extension that automates bulk assignment of Domain Security Policies in Workday security groups.

This tool is built for Workday administrators who frequently manage View, Modify, Get, and Put permissions and want to avoid repetitive manual configuration.

---

## Features

* Bulk add Domain Security Policies in Workday
* Supports View, Modify, Get, and Put access
* Accepts structured block input or plain lists
* Upload a Workday security export file and auto build the policy list
* Automatically skips policies that are already selected
* Correctly handles both auto selected and popup based selections
* Run and Stop controls
* Configurable delay between actions
* Designed specifically for the Maintain Domain Permissions for Security Group task

---

## Supported Workday Page

This extension works only on the following Workday task:

**Maintain Domain Permissions for Security Group**

If the page is not open, the extension will prompt you to navigate to it.

---

## Installation

### Option 1: Download as ZIP

1. Click **Code** on this GitHub repository
2. Select **Download ZIP**
3. Extract the ZIP file to a folder on your computer

### Option 2: Clone the Repository

```bash
git clone https://github.com/darshitpatel1/workday-security.git
```

---

### Load the Extension in Chrome

1. Open Google Chrome
2. Navigate to `chrome://extensions`
3. Enable **Developer mode** (top right)
4. Click **Load unpacked**
5. Select the extracted project folder

The extension will now be available in your Chrome toolbar.

---

## How It Works

1. Open the Maintain Domain Permissions for Security Group task in Workday
2. Open the extension popup
3. Paste the domain security policies
4. Click Run
5. The extension will automatically select and verify each policy

---

## Input Formats

### Block Format

```text
VIEW:
Worker Data: Current Job Profile Information
Affordable Care Act (ACA) Administration - USA

MODIFY:
Worker Data: Workers

GET:
Worktag REST API
```

### Plain List

```text
Worker Data: Current Job Profile Information
Affordable Care Act (ACA) Administration - USA
Reports: Headcount Plan
```

When using a plain list, select the target operation in the UI.

---

## File Upload

You can upload the exported Workday security file and let the extension build the input automatically.

Supported file

* Workday export Excel file

Expected columns

* Operation
* Domain Security Policy

How it maps

* View Only goes to VIEW
* Modify Only goes to MODIFY
* Get Only goes to GET
* Put Only goes to PUT
* View Modify goes to VIEW and MODIFY
* Get Put goes to GET and PUT

Steps

1. Open the extension popup
2. Click Choose File and upload the Workday export
3. The extension will parse the file and populate the lists
4. Review the generated input and click Run

---

## Options

* Skip if already selected
* Delay (ms)
* Stop on first error

---

## Disclaimer

This is an unofficial helper tool and is not affiliated with Workday, Inc.

Always test in non production tenants first.

---

## Purpose

Workday security administration is powerful but repetitive.

This tool exists to save time, reduce errors, and make bulk security updates safer and faster.
