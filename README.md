# RAM REV Automation Lab

Operational refund routing system built using **Google Sheets + Google Apps Script automation**.

This project demonstrates how manual customer operations workflows can be transformed into automated, event-driven systems that improve speed, accuracy, and scalability during live customer interactions.

---

## Overview

RAM REV Automation Lab simulates a real-world customer operations environment where agents process refund requests while speaking with customers.

Instead of manually copying data across multiple sheets and deciding routing rules, the system automates the entire workflow.

When an agent completes a request, automation validates data, determines routing logic, and moves the record to the correct processing queue instantly.

---

## Problem

Before automation, refund handling required:

- Manual copying between sheets
- Human routing decisions
- Multi-step operational updates
- High risk of errors and delays
- Slow processing during live calls

This created operational friction and inconsistent execution.

---

## Solution

A trigger-driven automation system that:

✅ Detects workflow completion  
✅ Validates required fields  
✅ Applies business routing logic  
✅ Moves records automatically  
✅ Prevents duplicate processing  

All actions occur instantly after status completion.

---

## Demo Workflow

1. Go to **Command_Center**
2. Enter: 1@test.com


into **Lookup Input**

3. Paste any address into **Raw Address**
4. Click **Verify Address** (Google Maps link)
5. Change **Status → Completed**

### Automation executes automatically:

- Data validation runs
- Routing decision executes
- Record moves to:
  - **Stripe Refunds**, OR
  - **Monthly Week Processing Queue**
- Notes column stamped with `AUTO-MOVED` timestamp

---

## Routing Logic

When Status becomes **Completed**:

### Stripe Routing
Record moves to **Stripe Refunds** if:

- Refund Type contains "Stripe", OR
- Member ID begins with `CAR`

### Monthly Processing Routing
Otherwise:

- Record routes to active month sheet
- Correct week calculated automatically
- First available row selected safely

---

## Technical Architecture

**Platform**
- Google Sheets
- Google Apps Script (JavaScript runtime)

**Automation Type**
- Installable `onEdit` trigger
- Event-driven workflow execution

**Core Features**
- Header-based column mapping
- Dynamic sheet routing
- Week-of-month computation
- Data validation safeguards
- Duplicate prevention markers

---

## Key Automation Concepts Demonstrated

- Workflow automation design
- Operational intelligence systems
- Event-driven scripting
- Data validation pipelines
- Process scalability engineering
- Customer operations automation

---

## Live Demo

Google Sheet (sanitized demo data):

👉 https://docs.google.com/spreadsheets/d/1pkzLjlBe166fFBBq7fm5BrWnzBvHx5cndvP5fpBzs-A/edit?usp=sharing

---

## Author

**Doc Pulliam**

Customer Experience & Automation Manager  
Python Workflow Automation | Operational Intelligence | SaaS Customer Operations

LinkedIn: https://linkedin.com/in/deandre-pulliam-b347b73b/

---

## Purpose

This project is part of a portfolio demonstrating automation-driven customer operations systems designed to reduce manual work and enable scalable service delivery.
