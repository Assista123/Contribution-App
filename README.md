# Contribution Group System (V2)

A robust, full-featured group financial management and contribution tracking platform designed for cooperatives, savings circles (Ajo/Esusu), associations, and community groups. Version 2 introduces advanced batch management, automated bank alert reconciliation, sequential payout tracking, and strict administrative slot controls.

## 🚀 Key V2 Features

### 1. Advanced Batch & Slot Management
* **Multi-Batch Creation:** Administer multiple parallel financial batches with custom frequencies (Daily, Weekly, Bi-weekly, Monthly) and adjustable slot prices.
* **Member Slot Requests & Approvals:** Members can browse available slots, request assignments, and await secure administrative lock-in.
* **Leading-Zero Protection:** Strict formatting handling for Nigerian phone numbers and bank account numbers to preserve critical leading zeros.

### 2. Automated & Manual Payment Reconciliation
* **Smart Bank Alert Auto-Matcher:** Integrates with partner banks (Kuda, Moniepoint, ALAT) via Make.com and Google Apps Script webhooks to automatically parse credit alerts and match them against pending transaction reference tags (`TXN-`).
* **Partial Payment Tracking:** Automatically calculates slot balances, tracking whether contributions are fully paid or partially fulfilled.
* **Manual Review Queue:** Un-tagged or ambiguous incoming payments are securely logged in an `Bank_Alerts_Log` for manual administrative review.

### 3. Sequential Payouts & Financial Summaries
* **Due Payout Calculation:** Automatically determines the next eligible slot and member due for disbursement based on fulfillment rules.
* **Batch Financial Summaries:** Real-time breakdown of grand total paid, total outstanding balances, and active schedule pacing.
* **Admin Portal & Security:** Secure PIN-based admin authentication and role management.

---

## 🛠️ Tech Stack & Architecture

* **Frontend:** [Cloudflare Pages](https://pages.cloudflare.com/) (Responsive, fast web portal optimized for mobile access).
* **Backend:** Google Apps Script Web App (Polymorphic `POST`/`GET` handling, data parsing, and auto-matching logic).
* **Database:** Google Sheets (Relational data structure covering Members, Batches, Member Slots, Contributions, Payouts, Bank Alerts, and Admins).
* **Automation:** Make.com (Mailhook processing and HTTP payload forwarding).

---

## ⚙️ Quick Setup Guide

### 1. Backend Setup (Google Apps Script)
1. Create a Google Sheet named `Contribution_Group_System`.
2. Open **Extensions > Apps Script**, paste the backend code, and save.
3. Run `setupDatabase()` to auto-initialize all required sheets with professional headers and formatting.
4. Click **Deploy > New deployment**, select **Web app**, set execution to **Me**, and access to **Anyone**. Copy your Web App URL.

### 2. Frontend Setup (Cloudflare Pages)
1. Push this repository to your GitHub account (`Assista123/Contribution-App`).
2. Log into Cloudflare Pages, connect your repository, and deploy.
3. Update your frontend configuration to point to your live Google Apps Script Web App URL.

### 3. Automation Setup (Make.com)
1. Configure a Mailhook to capture bank alert notification emails.
2. Add a text filter to process credit alerts.
3. Connect an **HTTP POST** module (`application/x-www-form-urlencoded`) pointing to your Web App URL with keys: `action=logBankAlert`, `bankName`, and `senderDetails`.

---

## 🔒 Security & Privacy Best Practices
When connecting Gmail for bank alert automation, we recommend setting up a dedicated free Google account used exclusively for receiving bank notifications to maintain clean data separation and strict privacy.
