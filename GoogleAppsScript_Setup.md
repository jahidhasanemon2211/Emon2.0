# Google Apps Script Setup Guide for Emon 2.0 NOC CMS

This guide provides the complete, copy-paste-ready Google Apps Script backend code for your website's **Live Visitor Logs**, **Contact Messages**, and **AI Chatbot logs**.

---

## 🛠️ Step-by-Step Setup Instructions

### Step 1: Create a Google Spreadsheet
1. Go to [Google Sheets](https://sheets.google.com) and create a **blank spreadsheet**.
2. Name your spreadsheet (e.g., `Emon 2.0 Database`).
3. You do **not** need to create any sheets manually. The script below will automatically create the required sheets (`VisitorLogs`, `ContactMessages`, and `ChatLogs`) with headers on their first run!

---

### Step 2: Open Apps Script Editor
1. In the top menu of your Google Spreadsheet, go to **Extensions** > **Apps Script**.
2. Delete any default code in the editor (e.g., `function myFunction() { ... }`).

---

### Step 3: Paste the Backend Code
Copy the entire block of code below and paste it into the editor (`Code.gs`):

```javascript
// Google Apps Script backend for Emon 2.0 NOC CMS
// Handles GET (fetching visitor logs) and POST (logging visitors, chats, & contact form)

function doGet(e) {
  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    let sheet = doc.getSheetByName("VisitorLogs");
    
    // If VisitorLogs sheet doesn't exist yet, return an empty array
    if (!sheet) {
      return ContentService.createTextOutput(JSON.stringify([]))
        .setMimeType(ContentService.MimeType.JSON);
    }
    
    const rows = sheet.getDataRange().getValues();
    const headers = rows[0];
    const data = [];
    
    // Loop through rows (skip header row at index 0)
    for (let i = 1; i < rows.length; i++) {
      const row = rows[i];
      const log = {};
      for (let j = 0; j < headers.length; j++) {
        const headerName = headers[j];
        let val = row[j];
        
        // Format dates correctly to ISO string
        if (val instanceof Date) {
          val = val.toISOString();
        }
        log[headerName] = val;
      }
      data.push(log);
    }
    
    // Return logs as a clean JSON response (Google automatically adds CORS headers)
    return ContentService.createTextOutput(JSON.stringify(data))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}

function doPost(e) {
  try {
    const doc = SpreadsheetApp.getActiveSpreadsheet();
    let payload = {};
    
    // Parse incoming request body
    if (e.postData && e.postData.contents) {
      // Request sent with Content-Type: application/json
      payload = JSON.parse(e.postData.contents);
    } else {
      // Request sent via standard FormData parameters
      payload = e.parameter;
    }
    
    const timestamp = new Date();
    
    // Case 1: Chatbot Logs
    if (payload.type === "chat") {
      let sheet = doc.getSheetByName("ChatLogs");
      if (!sheet) {
        sheet = doc.insertSheet("ChatLogs");
        sheet.appendRow(["Timestamp", "Name", "Query", "Reply", "IP", "City", "Country", "ISP"]);
        sheet.getRange(1, 1, 1, 8).setFontWeight("bold").setBackground("#e0e7ff");
      }
      sheet.appendRow([
        timestamp,
        payload.name || "Anonymous",
        payload.query || "",
        payload.reply || "",
        payload.ip || "",
        payload.city || "",
        payload.country || "",
        payload.org || ""
      ]);
    }
    // Case 2: Contact Form Submissions
    else if (payload.name && payload.email && payload.message) {
      let sheet = doc.getSheetByName("ContactMessages");
      if (!sheet) {
        sheet = doc.insertSheet("ContactMessages");
        sheet.appendRow(["Timestamp", "Name", "Email", "Message"]);
        sheet.getRange(1, 1, 1, 4).setFontWeight("bold").setBackground("#fcd34d");
      }
      sheet.appendRow([
        timestamp,
        payload.name,
        payload.email,
        payload.message
      ]);
    }
    // Case 3: Live Visitor Tracking
    else if (payload.ip) {
      let sheet = doc.getSheetByName("VisitorLogs");
      if (!sheet) {
        sheet = doc.insertSheet("VisitorLogs");
        sheet.appendRow(["Timestamp", "IP", "ISP", "City", "Country", "Browser", "OS"]);
        sheet.getRange(1, 1, 1, 7).setFontWeight("bold").setBackground("#a7f3d0");
      }
      sheet.appendRow([
        timestamp,
        payload.ip,
        payload.isp || "",
        payload.city || "",
        payload.country || "",
        payload.browser || "",
        payload.os || ""
      ]);
    }
    
    return ContentService.createTextOutput(JSON.stringify({ result: "success" }))
      .setMimeType(ContentService.MimeType.JSON);
      
  } catch (error) {
    return ContentService.createTextOutput(JSON.stringify({ result: "error", error: error.message }))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
```

---

### Step 4: Deploy as a Web App (CRITICAL)
For the script to work publicly, you must deploy it with correct permissions:
1. In the top-right corner of the Apps Script editor, click **Deploy** > **New deployment**.
2. Click the gear icon next to "Select type" and choose **Web app**.
3. Fill in the deployment details:
   - **Description**: `Emon 2.0 NOC CMS Backend`
   - **Execute as**: **Me (your-email@gmail.com)** *(This is critical so it can write to your Sheet)*
   - **Who has access**: **Anyone** *(This is critical so visitors and your admin panel can access it without Google sign-in)*
4. Click **Deploy**.
5. Google will ask you to authorize access. Click **Authorize Access**, select your Google account, click **Advanced** (at the bottom), and then click **Go to Untitled project (unsafe)** or **Allow**.
6. Once deployed, copy the **Web app URL** (it ends with `/exec`).

---

### Step 5: Save and Update your Website
Now that you have the Web App URL, configure it in your website admin panel:
1. Open your Emon 2.0 Web Admin Panel (`admin.html`).
2. Go to the **Contact & Socials** tab.
3. Paste your new Google Apps Script Web App URL into the **Visitor Log API URL** field.
4. Click **Save & Deploy Content** to publish this update to your live repository.

Now click **Visitor Logs** and hit **Refresh Logs**. Your live visitor logs will load perfectly! 🚀
