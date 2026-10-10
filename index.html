/**
 * Contribution Group System - Complete Backend
 * Features: Auto-Matcher, Offline Payments, Settings, Manual Match, Frequency Tracking
 */

const SS = SpreadsheetApp.openById("1xH8nBzRfW0VUoB2ZXxtbLudpdHm664wvKoZLDWPudXU");
const SHEETS = {
  MEMBERS: "Members",
  BATCHES: "Batches",
  SLOTS: "Member_Slots",
  CONTRIBUTIONS: "Contributions",
  PAYOUTS: "Payouts",
  ALERTS: "Bank_Alerts_Log",
  ADMINS: "Admins",
  SETTINGS: "Settings"
};

function setupDatabase() {
  const sheetConfigs = [
    { name: SHEETS.MEMBERS, headers: ["Member_ID", "Full_Name", "Phone_Number", "Payout_Bank_Name", "Payout_Account_Number", "Payout_Account_Name", "Date_Joined", "Status", "Notes"] },
    { name: SHEETS.BATCHES, headers: ["Batch_ID", "Batch_Name", "Total_Slots", "Slot_Price", "Frequency", "Start_Date", "Slots_Fulfilled", "Status", "Notes"] },
    { name: SHEETS.SLOTS, headers: ["Slot_ID", "Batch_ID", "Member_ID", "Status", "Balance_Paid", "Notes"] },
    { name: SHEETS.CONTRIBUTIONS, headers: ["Transaction_ID", "Slot_ID", "Batch_ID", "Amount", "Payment_Method", "Sender_Account_Name", "Sender_Account_Last4", "Status", "Timestamp", "Verified_By"] },
    { name: SHEETS.PAYOUTS, headers: ["Payout_ID", "Batch_ID", "Slot_ID", "Amount", "Date_Disbursed", "Status", "Disbursed_By", "Notes"] },
    { name: SHEETS.ALERTS, headers: ["Alert_ID", "Timestamp", "Bank_Name", "Raw_Amount", "Sender_Name", "Sender_Details", "Match_Status", "Matched_Transaction_ID"] },
    { name: SHEETS.SETTINGS, headers: ["Setting_Key", "Setting_Value", "Updated_At", "Updated_By"] },
    { name: SHEETS.ADMINS, headers: ["Admin_ID", "Full_Name", "Phone_Number", "Bank_Name", "Pin_Code", "Role"] }
  ];

  sheetConfigs.forEach(config => {
    let sheet = SS.getSheetByName(config.name);
    if (!sheet) sheet = SS.insertSheet(config.name);
    sheet.clear();
    sheet.appendRow(config.headers);
    sheet.getRange(1, 1, 1, config.headers.length)
      .setBackground("#1a73e8").setFontColor("#ffffff").setFontWeight("bold");
    sheet.setFrozenRows(1);
  });

  try {
    SS.getSheetByName(SHEETS.MEMBERS).getRange("C:C").setNumberFormat('@');
    SS.getSheetByName(SHEETS.MEMBERS).getRange("E:E").setNumberFormat('@');
    SS.getSheetByName(SHEETS.ADMINS).getRange("C:C").setNumberFormat('@');
  } catch (e) {}

  try {
    SpreadsheetApp.getUi().alert("Database initialized successfully!");
  } catch (err) {
    Logger.log("UI alert skipped (non-UI context).");
  }
}

function onEdit(e) {
  const sheet = e.range.getSheet();
  const sheetName = sheet.getName();
  const row = e.range.getRow();
  const col = e.range.getColumn();

  if (sheetName === SHEETS.MEMBERS && col === 2 && row > 1) {
    const idCell = sheet.getRange(row, 1);
    if (!idCell.getValue()) idCell.setValue("MEM-" + (1000 + sheet.getLastRow() - 1));
    if (!sheet.getRange(row, 7).getValue()) sheet.getRange(row, 7).setValue(new Date());
    if (!sheet.getRange(row, 8).getValue()) sheet.getRange(row, 8).setValue("Active");
  }

  if (sheetName === SHEETS.BATCHES && col === 3 && row > 1) {
    const batchIdCell = sheet.getRange(row, 1);
    const totalSlots = e.value;
    if (batchIdCell.getValue() && totalSlots > 0) {
      const batchId = batchIdCell.getValue();
      const slotsSheet = SS.getSheetByName(SHEETS.SLOTS);
      for (let i = 1; i <= totalSlots; i++) {
        slotsSheet.appendRow([`${batchId}_${String(i).padStart(2, '0')}`, batchId, "", "Unassigned", 0, ""]);
      }
    }
  }
}

function doGet(e) {
  if (e.parameter && e.parameter.action) {
    const action = e.parameter.action;
    try {
      if (action === "getMemberData") return jsonResponse(getMemberPortalData(e.parameter.memberId, e.parameter.phone));
      if (action === "getAdminData") return jsonResponse(getAdminDashboardData());
      if (action === "getCurrentPayoutDue") return jsonResponse(getCurrentPayoutDue(e.parameter.batchId));
      if (action === "getAvailableSlots") return jsonResponse(getAvailableSlots());
      if (action === "getPendingSlots") return jsonResponse(getPendingSlotRequests());
      if (action === "getBatchSummary") return jsonResponse(getBatchPaymentSummary(e.parameter.batchId));
      if (action === "getUnmatchedAlerts") return jsonResponse(getUnmatchedAlerts());
      if (action === "getActiveBatches") return jsonResponse(getActiveBatches());
      if (action === "getCollectionAccount") return jsonResponse({ status: "success", collectionAccount: getCollectionAccount() });
      return jsonResponse({ status: "error", message: "Invalid GET action." });
    } catch (err) {
      return jsonResponse({ status: "error", message: err.toString() });
    }
  }
  return HtmlService.createHtmlOutputFromFile('Index')
    .setTitle('Contribution Group Portal')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}

function doPost(e) {
  try {
    let data = {};
    if (e.parameter && Object.keys(e.parameter).length > 0) {
      data = e.parameter;
    } else if (e.postData && e.postData.contents) {
      try {
        data = JSON.parse(e.postData.contents);
      } catch (err) {
        let params = {};
        e.postData.contents.split("&").forEach(part => {
          let item = part.split("=");
          if (item.length === 2) params[decodeURIComponent(item[0])] = decodeURIComponent(item[1]);
        });
        data = params;
      }
    }

    const action = data.action;
    Logger.log("Incoming Action: " + action);

    if (action === "submitPaymentIntent") return jsonResponse(recordPaymentIntent(data));
    if (action === "confirmPayment") return jsonResponse(updatePaymentStatus(data.transactionId, "Confirmed", data.adminName));
    if (action === "rejectPayment") return jsonResponse(updatePaymentStatus(data.transactionId, "Failed", data.adminName));
    if (action === "adminLogin") return jsonResponse(verifyAdminLogin(data.phone, data.pin));
    if (action === "registerMember") return jsonResponse(registerMember(data));
    if (action === "logBankAlert") return jsonResponse(logBankAlert(data));
    if (action === "requestSlot") return jsonResponse(requestSlot(data.memberId, data.slotId));
    if (action === "reviewSlot") return jsonResponse(reviewSlotRequest(data.slotId, data.decision));
    if (action === "createBatch") return jsonResponse(createBatch(data));
    if (action === "recordOfflinePayment") return jsonResponse(recordOfflinePayment(data));
    if (action === "updateSetting") return jsonResponse(updateSetting(data.key, data.value, data.adminName));
    if (action === "manualMatchAlert") return jsonResponse(manualMatchAlert(data.alertId, data.transactionId, data.adminName));

    return jsonResponse({ status: "error", message: "Invalid POST action: " + action });
  } catch (err) {
    return jsonResponse({ status: "error", message: err.toString() });
  }
}

function jsonResponse(payload) {
  return ContentService.createTextOutput(JSON.stringify(payload))
    .setMimeType(ContentService.MimeType.JSON);
}

/* ===================== MEMBERS ===================== */

function registerMember(data) {
  try {
    const membersSheet = SS.getSheetByName(SHEETS.MEMBERS);
    const rows = membersSheet.getDataRange().getValues();
    const cleanPhoneInput = String(data.phone || "").trim();

    for (let i = 1; i < rows.length; i++) {
      if (String(rows[i][2] || "").replace(/^['"]/, "").trim() === cleanPhoneInput) {
        return { status: "error", message: "This phone number is already registered." };
      }
    }

    const nextId = "MEM-" + (1000 + rows.length - 1);
    const targetRow = membersSheet.getLastRow() + 1;

    membersSheet.getRange(targetRow, 3).setNumberFormat('@');
    membersSheet.getRange(targetRow, 5).setNumberFormat('@');

    membersSheet.getRange(targetRow, 1, 1, 9).setValues([[
      nextId,
      data.fullName || "",
      "'" + cleanPhoneInput,
      data.payoutBankName || "",
      "'" + String(data.payoutAccountNumber || "").trim(),
      data.payoutAccountName || "",
      new Date(),
      "Active",
      "Self-registered"
    ]]);

    return { status: "success", message: `Registration successful! Member ID: ${nextId}`, memberId: nextId };
  } catch (err) {
    return { status: "error", message: err.toString() };
  }
}

function getMemberPortalData(memberId, phone) {
  const membersSheet = SS.getSheetByName(SHEETS.MEMBERS);
  const data = membersSheet.getDataRange().getValues();
  let verifiedMember = null;
  const cleanInputPhone = String(phone || "").trim();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === memberId && String(data[i][2] || "").replace(/^'/, "").trim() === cleanInputPhone) {
      verifiedMember = {
        memberId: data[i][0],
        fullName: data[i][1],
        bankName: data[i][3],
        accountNumber: String(data[i][4] || "").replace(/^'/, ""),
        accountName: data[i][5],
        status: data[i][7]
      };
      break;
    }
  }

  if (!verifiedMember) return { status: "error", message: "Invalid Member ID or Phone." };

  const slotData = SS.getSheetByName(SHEETS.SLOTS).getDataRange().getValues();
  let mySlots = [];
  for (let i = 1; i < slotData.length; i++) {
    if (slotData[i][2] === memberId) {
      mySlots.push({
        slotId: slotData[i][0],
        batchId: slotData[i][1],
        status: slotData[i][3],
        balancePaid: slotData[i][4] || 0
      });
    }
  }

  return { status: "success", member: verifiedMember, slots: mySlots };
}

/* ===================== SLOTS ===================== */

function getAvailableSlots() {
  const slotData = SS.getSheetByName(SHEETS.SLOTS).getDataRange().getValues();
  let available = [];
  for (let i = 1; i < slotData.length; i++) {
    const status = String(slotData[i][3]).trim();
    const memberId = slotData[i][2];
    if (!memberId || status === "Unassigned" || status === "") {
      available.push({ slotId: slotData[i][0], batchId: slotData[i][1] });
    }
  }
  return { status: "success", availableSlots: available };
}

function requestSlot(memberId, slotId) {
  const slotsSheet = SS.getSheetByName(SHEETS.SLOTS);
  const data = slotsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === slotId) {
      if (data[i][3] !== "Unassigned" && data[i][3] !== "") {
        return { status: "error", message: "Slot is no longer available." };
      }
      slotsSheet.getRange(i + 1, 3).setValue(memberId);
      slotsSheet.getRange(i + 1, 4).setValue("Pending Approval");
      return { status: "success", message: `Slot ${slotId} requested successfully. Awaiting admin approval.` };
    }
  }
  return { status: "error", message: "Slot not found." };
}

function getPendingSlotRequests() {
  const slotsSheet = SS.getSheetByName(SHEETS.SLOTS);
  const memberData = SS.getSheetByName(SHEETS.MEMBERS).getDataRange().getValues();
  const slotData = slotsSheet.getDataRange().getValues();

  const memberMap = {};
  for (let m = 1; m < memberData.length; m++) memberMap[memberData[m][0]] = memberData[m][1];

  let pending = [];
  for (let i = 1; i < slotData.length; i++) {
    if (slotData[i][3] === "Pending Approval") {
      pending.push({
        slotId: slotData[i][0],
        batchId: slotData[i][1],
        memberId: slotData[i][2],
        memberName: memberMap[slotData[i][2]] || "Unknown"
      });
    }
  }
  return { status: "success", pendingSlots: pending };
}

function reviewSlotRequest(slotId, decision) {
  const slotsSheet = SS.getSheetByName(SHEETS.SLOTS);
  const data = slotsSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === slotId) {
      if (decision === "Approve") {
        slotsSheet.getRange(i + 1, 4).setValue("Active");
        return { status: "success", message: `Slot ${slotId} approved and locked in!` };
      } else {
        slotsSheet.getRange(i + 1, 3).setValue("");
        slotsSheet.getRange(i + 1, 4).setValue("Unassigned");
        return { status: "success", message: `Slot ${slotId} request rejected.` };
      }
    }
  }
  return { status: "error", message: "Slot not found." };
}

/* ===================== BATCHES ===================== */

function createBatch(data) {
  const batchesSheet = SS.getSheetByName(SHEETS.BATCHES);
  const slotsSheet = SS.getSheetByName(SHEETS.SLOTS);

  batchesSheet.appendRow([
    data.batchId,
    data.batchName,
    data.totalSlots,
    data.slotPrice,
    data.frequency || "Weekly",
    data.startDate || new Date(),
    0,
    "Active",
    "Created via Admin Dashboard"
  ]);

  for (let i = 1; i <= Number(data.totalSlots); i++) {
    const slotId = `${data.batchId}_${String(i).padStart(2, '0')}`;
    slotsSheet.appendRow([slotId, data.batchId, "", "Unassigned", 0, ""]);
  }

  return {
    status: "success",
    message: `Batch ${data.batchId} created with ${data.totalSlots} slots (${data.frequency || "Weekly"} frequency)!`
  };
}

function getActiveBatches() {
  const batchData = SS.getSheetByName(SHEETS.BATCHES).getDataRange().getValues();
  let batches = [];
  for (let i = 1; i < batchData.length; i++) {
    if (String(batchData[i][7]).toLowerCase() === "active") {
      batches.push({
        batchId: batchData[i][0],
        batchName: batchData[i][1],
        totalSlots: batchData[i][2],
        slotPrice: batchData[i][3],
        frequency: batchData[i][4] || "Weekly"
      });
    }
  }
  return { status: "success", activeBatches: batches };
}

/* ===================== PAYMENTS ===================== */

function recordPaymentIntent(data) {
  const contribSheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);
  const txId = "TXN-" + Math.floor(100000 + Math.random() * 900000);
  const slotsList = Array.isArray(data.slotIds) ? data.slotIds.join(", ") : String(data.slotIds || "");

  contribSheet.appendRow([
    txId, slotsList, data.batchId || "Multi-Batch", data.amount,
    "Bank Transfer", data.senderName, "", "Pending", new Date(), ""
  ]);

  const collection = getCollectionAccount();

  return {
    status: "success",
    message: `Reference Tag generated: ${txId}. Include this tag in your bank transfer narration for ₦${Number(data.amount).toLocaleString()}.`,
    transactionId: txId,
    amount: Number(data.amount),
    collectionAccount: collection
  };
}

function applyPaymentToSlot(slotId, paymentAmount) {
  const slotsSheet = SS.getSheetByName(SHEETS.SLOTS);
  const batchData = SS.getSheetByName(SHEETS.BATCHES).getDataRange().getValues();
  const slotData = slotsSheet.getDataRange().getValues();

  let slotRowIndex = -1, batchId = "", currentBalance = 0;
  for (let i = 1; i < slotData.length; i++) {
    if (slotData[i][0] === slotId) {
      slotRowIndex = i + 1;
      batchId = slotData[i][1];
      currentBalance = Number(slotData[i][4]) || 0;
      break;
    }
  }
  if (slotRowIndex === -1) return;

  let slotPrice = 0;
  for (let b = 1; b < batchData.length; b++) {
    if (batchData[b][0] === batchId) {
      slotPrice = Number(batchData[b][3]) || 0;
      break;
    }
  }

  const newBalance = currentBalance + Number(paymentAmount);
  const newStatus = (slotPrice > 0 && newBalance >= slotPrice) ? "Confirmed" : "Partially Paid";

  slotsSheet.getRange(slotRowIndex, 4).setValue(newStatus);
  slotsSheet.getRange(slotRowIndex, 5).setValue(newBalance);
}

function updatePaymentStatus(transactionId, newStatus, adminName) {
  const contribSheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);
  const data = contribSheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === transactionId) {
      const slotIds = String(data[i][1]).split(",").map(s => s.trim());
      const totalAmount = Number(data[i][3]);
      const amountPerSlot = totalAmount / (slotIds.length || 1);
      const oldStatus = data[i][7];

      contribSheet.getRange(i + 1, 8).setValue(newStatus);
      contribSheet.getRange(i + 1, 10).setValue(adminName || "Admin");

      if (newStatus === "Confirmed" && oldStatus !== "Confirmed") {
        slotIds.forEach(sId => {
          if (sId) applyPaymentToSlot(sId, amountPerSlot);
        });
      }

      return { status: "success", message: `Transaction ${transactionId} marked as ${newStatus}.` };
    }
  }
  return { status: "error", message: "Transaction ID not found." };
}

function recordOfflinePayment(data) {
  const contribSheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);
  const txId = "OFF-" + Math.floor(100000 + Math.random() * 900000);
  const slotsList = Array.isArray(data.slotIds) ? data.slotIds.join(", ") : String(data.slotIds || "");
  const amount = Number(data.amount) || 0;
  const adminName = data.adminName || "Admin";

  if (!slotsList || amount <= 0) {
    return { status: "error", message: "Slot ID(s) and a valid amount are required." };
  }

  contribSheet.appendRow([
    txId, slotsList, data.batchId || "Offline", amount,
    data.paymentMethod || "Offline / Cash / Direct Transfer",
    data.senderName || "Offline Payer", "", "Confirmed", new Date(), adminName
  ]);

  const slotIds = slotsList.split(",").map(s => s.trim());
  const amountPerSlot = amount / slotIds.length;
  slotIds.forEach(sId => {
    if (sId) applyPaymentToSlot(sId, amountPerSlot);
  });

  return {
    status: "success",
    message: `Offline payment recorded as ${txId}. ₦${amount.toLocaleString()} applied to ${slotsList}.`,
    transactionId: txId
  };
}

/* ===================== BANK ALERTS ===================== */

function runAutoMatcher(alertId, bankName, rawAmount, senderName, rawNarration) {
  const targetBank = String(bankName || "").toUpperCase();
  if (!targetBank.includes("KUDA") && !targetBank.includes("MONIEPOINT") && !targetBank.includes("ALAT")) {
    updateAlertStatus(alertId, "Ignored (Non-Target Bank)", "");
    return { status: "success", matched: false, message: "Alert ignored. Only supported partner banks are accepted." };
  }

  const contribSheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);
  const contribData = contribSheet.getDataRange().getValues();
  let matchedTxId = null, matchedRowIndex = -1;
  const alertAmount = Number(rawAmount);
  const narration = String(rawNarration || "").toUpperCase();

  for (let i = 1; i < contribData.length; i++) {
    const txId = contribData[i][0];
    if (contribData[i][7] === "Pending" && alertAmount === Number(contribData[i][3]) && narration.includes(txId.toUpperCase())) {
      matchedTxId = txId;
      matchedRowIndex = i + 1;
      break;
    }
  }

  if (matchedTxId && matchedRowIndex !== -1) {
    const slotIds = String(contribData[matchedRowIndex - 1][1]).split(",").map(s => s.trim());
    const amountPerSlot = alertAmount / slotIds.length;

    contribSheet.getRange(matchedRowIndex, 8).setValue("Confirmed");
    contribSheet.getRange(matchedRowIndex, 10).setValue(`Auto-Matched (${bankName})`);
    updateAlertStatus(alertId, "Matched", matchedTxId);

    slotIds.forEach(sId => {
      if (sId) applyPaymentToSlot(sId, amountPerSlot);
    });
    return { status: "success", matched: true, transactionId: matchedTxId };
  }

  updateAlertStatus(alertId, "Unmatched", "");
  return { status: "success", matched: false, message: "No matching reference code found. Available for manual review." };
}

function logBankAlert(data) {
  const alertSheet = SS.getSheetByName(SHEETS.ALERTS);
  const alertId = "ALT-" + Math.floor(100000 + Math.random() * 900000);

  const rawText = data.senderDetails || data.text || data.body || JSON.stringify(data);
  const bankName = data.bankName || "ALAT";

  const amountMatch = rawText.match(/(?:NGN|₦)\s*([0-9,]+\.[0-9]+|[0-9,]+)/i) ||
                      rawText.match(/Amount[:\s]*([0-9,.]+)/i) ||
                      rawText.match(/([0-9,]+\.[0-9]{2})/);

  let parsedAmount = amountMatch ? Number(amountMatch[1].replace(/,/g, '')) : (Number(data.amount) || 0);

  const txMatch = rawText.match(/TXN-[A-Za-z0-9]+/i);
  const extractedTxId = txMatch ? txMatch[0].toUpperCase() : null;

  // Improved sender extraction – prioritise text after "NIP:"
  let senderName = data.senderName || "Bank User";
  const nipMatch = rawText.match(/NIP[:\s]*([A-Za-z0-9\s\.\-]+)/i);
  if (nipMatch) {
    senderName = nipMatch[1].trim().split(/\s+/).slice(0, 3).join(" ");
  } else {
    const senderMatch = rawText.match(/Sender[:\s]*([^\r\n]+)/i) || rawText.match(/From[:\s]*([^\r\n]+)/i);
    if (senderMatch) senderName = senderMatch[1].trim();
  }

  const noteMatch = rawText.match(/Note[:\s]*([^\r\n]+)/i) || rawText.match(/Narration[:\s]*([^\r\n]+)/i);
  const parsedNarration = noteMatch ? noteMatch[1].trim() : rawText;

  const initialStatus = extractedTxId ? "Unmatched" : "No Reference (Manual Review)";

  alertSheet.appendRow([
    alertId, new Date(), bankName, parsedAmount, senderName,
    parsedNarration.substring(0, 500), initialStatus, ""
  ]);

  if (extractedTxId) {
    return runAutoMatcher(alertId, bankName, parsedAmount, senderName, parsedNarration);
  } else {
    return {
      status: "success",
      matched: false,
      alertId: alertId,
      message: "Logged without reference tag. Available for manual matching."
    };
  }
}

function updateAlertStatus(alertId, matchStatus, matchedTxId) {
  const alertSheet = SS.getSheetByName(SHEETS.ALERTS);
  const data = alertSheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === alertId) {
      alertSheet.getRange(i + 1, 7).setValue(matchStatus);
      alertSheet.getRange(i + 1, 8).setValue(matchedTxId || "");
      break;
    }
  }
}

function manualMatchAlert(alertId, transactionId, adminName) {
  if (!alertId || !transactionId) {
    return { status: "error", message: "Both alertId and transactionId are required." };
  }

  const alertSheet = SS.getSheetByName(SHEETS.ALERTS);
  const contribSheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);

  const alertData = alertSheet.getDataRange().getValues();
  let alertRow = -1;
  for (let i = 1; i < alertData.length; i++) {
    if (alertData[i][0] === alertId) {
      alertRow = i + 1;
      if (String(alertData[i][6]).toLowerCase().includes("matched")) {
        return { status: "error", message: `Alert ${alertId} is already matched.` };
      }
      break;
    }
  }
  if (alertRow === -1) return { status: "error", message: "Alert ID not found." };

  const contribData = contribSheet.getDataRange().getValues();
  let txRow = -1, txAmount = 0, slotIdsString = "";
  for (let i = 1; i < contribData.length; i++) {
    if (contribData[i][0] === transactionId) {
      if (contribData[i][7] !== "Pending") {
        return { status: "error", message: `Transaction ${transactionId} is not in Pending status.` };
      }
      txRow = i + 1;
      txAmount = Number(contribData[i][3]) || 0;
      slotIdsString = String(contribData[i][1]);
      break;
    }
  }
  if (txRow === -1) return { status: "error", message: "Transaction ID not found." };

  const slotIds = slotIdsString.split(",").map(s => s.trim());
  const amountPerSlot = txAmount / (slotIds.length || 1);

  contribSheet.getRange(txRow, 8).setValue("Confirmed");
  contribSheet.getRange(txRow, 10).setValue(adminName || "Manual Match");

  slotIds.forEach(sId => {
    if (sId) applyPaymentToSlot(sId, amountPerSlot);
  });

  updateAlertStatus(alertId, "Manually Matched", transactionId);

  return {
    status: "success",
    message: `Alert ${alertId} manually matched to ${transactionId}. Payment confirmed.`
  };
}

function getUnmatchedAlerts() {
  const alertSheet = SS.getSheetByName(SHEETS.ALERTS);
  const data = alertSheet.getDataRange().getValues();
  let unmatched = [];

  for (let i = 1; i < data.length; i++) {
    const status = String(data[i][6] || "");
    if (status === "No Reference (Manual Review)" || status === "Unmatched") {
      unmatched.push({
        alertId: data[i][0],
        timestamp: data[i][1],
        bankName: data[i][2],
        amount: data[i][3],
        senderName: data[i][4],
        details: data[i][5],
        matchStatus: status
      });
    }
  }
  return { status: "success", unmatchedAlerts: unmatched };
}

/* ===================== ADMIN ===================== */

function getAdminDashboardData() {
  const contribSheet = SS.getSheetByName(SHEETS.CONTRIBUTIONS);
  const data = contribSheet.getDataRange().getValues();
  let pendingPayments = [];
  for (let i = 1; i < data.length; i++) {
    if (data[i][7] === "Pending") {
      pendingPayments.push({
        transactionId: data[i][0],
        slotId: data[i][1],
        batchId: data[i][2],
        amount: data[i][3],
        method: data[i][4],
        senderName: data[i][5],
        timestamp: data[i][8]
      });
    }
  }
  return { status: "success", pendingPayments: pendingPayments };
}

function verifyAdminLogin(phone, pin) {
  const adminSheet = SS.getSheetByName(SHEETS.ADMINS);
  if (!adminSheet) return { status: "error", message: "Admins sheet missing." };

  const data = adminSheet.getDataRange().getValues();
  const cleanInputPhone = String(phone || "").replace(/\D/g, "").trim();
  const cleanInputPin = String(pin || "").trim();

  for (let i = 1; i < data.length; i++) {
    const sheetPhone = String(data[i][2] || "").replace(/^'/, "").replace(/\D/g, "").trim();
    const sheetPin = String(data[i][4] || "").trim();

    if (sheetPhone === cleanInputPhone && sheetPin === cleanInputPin) {
      return {
        status: "success",
        admin: {
          name: data[i][1],
          role: data[i][5] || "Admin"
        }
      };
    }
  }
  return { status: "error", message: "Invalid Admin Credentials." };
}

/* ===================== SETTINGS ===================== */

function getSetting(key) {
  const sheet = SS.getSheetByName(SHEETS.SETTINGS);
  if (!sheet) return "";
  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      return String(data[i][1] || "").trim();
    }
  }
  return "";
}

function getCollectionAccount() {
  return {
    bankName: getSetting("Collection_Bank_Name") || "Not set",
    accountNumber: getSetting("Collection_Account_Number") || "Not set",
    accountName: getSetting("Collection_Account_Name") || "Not set"
  };
}

function updateSetting(key, value, updatedBy) {
  const sheet = SS.getSheetByName(SHEETS.SETTINGS);
  if (!sheet) return { status: "error", message: "Settings sheet missing." };

  const data = sheet.getDataRange().getValues();
  for (let i = 1; i < data.length; i++) {
    if (String(data[i][0]).trim() === key) {
      sheet.getRange(i + 1, 2).setValue(value);
      sheet.getRange(i + 1, 3).setValue(new Date());
      sheet.getRange(i + 1, 4).setValue(updatedBy || "Admin");
      return { status: "success", message: `${key} updated.` };
    }
  }
  sheet.appendRow([key, value, new Date(), updatedBy || "Admin"]);
  return { status: "success", message: `${key} created.` };
}

/* ===================== REPORTS ===================== */

function getCurrentPayoutDue(batchId) {
  const slotData = SS.getSheetByName(SHEETS.SLOTS).getDataRange().getValues();
  const memberData = SS.getSheetByName(SHEETS.MEMBERS).getDataRange().getValues();
  const batchData = SS.getSheetByName(SHEETS.BATCHES).getDataRange().getValues();

  let slotPrice = 0;
  for (let b = 1; b < batchData.length; b++) {
    if (batchData[b][0] === batchId) {
      slotPrice = Number(batchData[b][3]) || 0;
      break;
    }
  }

  const memberMap = {};
  for (let m = 1; m < memberData.length; m++) {
    memberMap[memberData[m][0]] = {
      fullName: memberData[m][1],
      phone: String(memberData[m][2] || "").replace(/^'/, "").trim(),
      bankName: memberData[m][3],
      accountNumber: String(memberData[m][4] || "").replace(/^'/, ""),
      accountName: memberData[m][5]
    };
  }

  let targetSlot = null;
  for (let i = 1; i < slotData.length; i++) {
    if (slotData[i][1] === batchId && slotData[i][3] !== "Paid Out") {
      const memberId = slotData[i][2];
      const balancePaid = Number(slotData[i][4]) || 0;
      const memberInfo = memberMap[memberId] || {
        fullName: "Unassigned", phone: "", bankName: "", accountNumber: "", accountName: ""
      };

      targetSlot = {
        slotId: slotData[i][0],
        batchId: slotData[i][1],
        memberId: memberId,
        memberName: memberInfo.fullName,
        memberPhone: memberInfo.phone,
        payoutBankName: memberInfo.bankName,
        payoutAccountNumber: memberInfo.accountNumber,
        payoutAccountName: memberInfo.accountName,
        balancePaid: balancePaid,
        slotPrice: slotPrice,
        eligibleForPayout: balancePaid >= slotPrice,
        status: slotData[i][3]
      };
      break;
    }
  }
  return targetSlot
    ? { status: "success", duePayout: targetSlot }
    : { status: "success", message: "All slots fulfilled and paid out." };
}

function getBatchPaymentSummary(batchId) {
  const slotsSheet = SS.getSheetByName(SHEETS.SLOTS);
  const membersSheet = SS.getSheetByName(SHEETS.MEMBERS);
  const batchesSheet = SS.getSheetByName(SHEETS.BATCHES);

  const slotData = slotsSheet.getDataRange().getValues();
  const memberData = membersSheet.getDataRange().getValues();
  const batchData = batchesSheet.getDataRange().getValues();

  let batchName = "", slotPrice = 0, totalSlots = 0, frequency = "Weekly", startDate = new Date();

  for (let b = 1; b < batchData.length; b++) {
    if (batchData[b][0] === batchId) {
      batchName = batchData[b][1];
      totalSlots = Number(batchData[b][2]) || 0;
      slotPrice = Number(batchData[b][3]) || 0;
      frequency = batchData[b][4] || "Weekly";
      startDate = new Date(batchData[b][5] || new Date());
      break;
    }
  }

  if (!batchName && slotPrice === 0) return { status: "error", message: "Batch ID not found." };

  const now = new Date();
  const diffDays = Math.ceil(Math.abs(now - startDate) / (1000 * 60 * 60 * 24));

  let elapsedCycles = 1;
  if (frequency === "Daily") elapsedCycles = Math.max(1, diffDays);
  else if (frequency === "Weekly") elapsedCycles = Math.max(1, Math.ceil(diffDays / 7));
  else if (frequency === "Bi-weekly") elapsedCycles = Math.max(1, Math.ceil(diffDays / 14));
  else if (frequency === "Monthly") elapsedCycles = Math.max(1, Math.ceil(diffDays / 30));

  elapsedCycles = Math.min(elapsedCycles, totalSlots || elapsedCycles);
  const expectedPaidByNow = totalSlots > 0 ? Math.round((elapsedCycles / totalSlots) * slotPrice) : slotPrice;

  const memberMap = {};
  for (let m = 1; m < memberData.length; m++) memberMap[memberData[m][0]] = memberData[m][1];

  let paidList = [], owingList = [], grandTotalPaid = 0, currentCycleSlot = "N/A";

  for (let i = 1; i < slotData.length; i++) {
    if (slotData[i][1] === batchId) {
      const slotId = slotData[i][0];
      const memberId = slotData[i][2];
      const status = slotData[i][3];
      const balancePaid = Number(slotData[i][4]) || 0;
      const memberName = memberMap[memberId] || "Unassigned";

      grandTotalPaid += balancePaid;
      if (status !== "Paid Out" && currentCycleSlot === "N/A") currentCycleSlot = slotId;

      if (balancePaid >= slotPrice && slotPrice > 0) {
        paidList.push({ slotId, memberName, balancePaid, status: "Fully Paid" });
      } else {
        const outstanding = Math.max(0, slotPrice - balancePaid);
        const shortfall = Math.max(0, expectedPaidByNow - balancePaid);
        let pace = "On Track";
        if (shortfall > 0) pace = "Behind Schedule";
        else if (balancePaid === 0) pace = "Not Started";

        owingList.push({
          slotId, memberName, balancePaid,
          owingAmount: outstanding,
          expectedByNow: expectedPaidByNow,
          shortfall, pace
        });
      }
    }
  }

  return {
    status: "success",
    summary: {
      batchId, batchName, slotPrice, totalSlots, frequency, elapsedCycles,
      expectedPaidByNow, grandTotalPaid,
      totalBatchPayout: totalSlots * slotPrice,
      currentCycleSlot, paidList, owingList
    }
  };
}
