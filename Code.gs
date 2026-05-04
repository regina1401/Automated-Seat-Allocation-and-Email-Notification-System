function doGet(e) {
  if (e && e.parameter.page === "admin") {
    return HtmlService.createHtmlOutputFromFile("admin")
      .setTitle("Admin Dashboard");
  }

  if (e && e.parameter.page === "scanner") {
    return HtmlService.createHtmlOutputFromFile("scanner")
      .setTitle("Entry Scanner");
  }

  return HtmlService.createHtmlOutputFromFile("index")
    .setTitle("Seat Booking");
}

/* Get seats */
function getSeats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seats");
  const data = sheet.getDataRange().getValues();
  data.shift();
  return data;
}

/* Book up to 3 seats */
function bookSeat(
  seatIDs,
  name,
  roll,
  email,
  studentPhone,
  branch,
  guardian1Name,
  guardian1Phone,
  guardian1Email,
  guardian2Name,
  guardian2Phone,
  guardian2Email,
  eventName,
  hallName,
  organizedBy
) {
  const lock = LockService.getScriptLock();
  lock.waitLock(10000);

  try {
    const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seats");
    const data = sheet.getDataRange().getValues();

    if (!Array.isArray(seatIDs) || seatIDs.length === 0) {
      return { status: "error", message: "Please select at least one seat." };
    }

    if (seatIDs.length > 3) {
      return { status: "error", message: "You can select maximum 3 seats." };
    }

    /* Prevent same roll booking twice */
    for (let i = 1; i < data.length; i++) {
      if (String(data[i][4]).trim() === String(roll).trim()) {
        return { status: "error", message: "Roll already booked" };
      }
    }

    /* Validate that all requested seats are still available before writing anything */
    const seatRowMap = {};
    for (let i = 1; i < data.length; i++) {
      seatRowMap[String(data[i][0]).trim()] = i;
    }

    for (const seatID of seatIDs) {
      const rowIndex = seatRowMap[String(seatID).trim()];
      if (rowIndex === undefined) {
        return { status: "error", message: "Seat " + seatID + " not found" };
      }
      if (String(data[rowIndex][3]).trim() === "Booked") {
        return { status: "error", message: "Seat " + seatID + " already booked" };
      }
    }

    /* Book each seat */
    const bookedSeats = [];

    for (const seatID of seatIDs) {
      const i = seatRowMap[String(seatID).trim()];

      sheet.getRange(i + 1, 4).setValue("Booked");
      sheet.getRange(i + 1, 5).setValue(roll);
      sheet.getRange(i + 1, 6).setValue(name);
      sheet.getRange(i + 1, 7).setValue(email);
      sheet.getRange(i + 1, 8).setValue(studentPhone);
      sheet.getRange(i + 1, 9).setValue(branch);
      sheet.getRange(i + 1, 10).setValue(guardian1Name);
      sheet.getRange(i + 1, 11).setValue(guardian1Phone);
      sheet.getRange(i + 1, 12).setValue(guardian1Email);
      sheet.getRange(i + 1, 13).setValue(guardian2Name);
      sheet.getRange(i + 1, 14).setValue(guardian2Phone);
      sheet.getRange(i + 1, 15).setValue(guardian2Email);
      sheet.getRange(i + 1, 16).setValue(eventName);
      sheet.getRange(i + 1, 17).setValue(hallName);
      sheet.getRange(i + 1, 18).setValue(organizedBy);

      bookedSeats.push(seatID);
    }

    /* QR DATA */
    const qrData =
      "EVENT ENTRY PASS\n" +
      "Student : " + name + "\n" +
      "Roll    : " + roll + "\n" +
      "Seats   : " + bookedSeats.join(", ") + "\n" +
      "Event   : " + eventName + "\n" +
      "Hall    : " + hallName;

    const qrURL =
      "https://api.qrserver.com/v1/create-qr-code/?size=200x200&data=" +
      encodeURIComponent(qrData);

/* ---------------- STUDENT EMAIL ---------------- */

const studentHTML = `
<div style="font-family:Arial, sans-serif; background:#f2f4f7; padding:30px;">

<div style="max-width:650px; margin:auto; background:white; border-radius:8px; overflow:hidden; box-shadow:0 6px 18px rgba(0,0,0,0.1);">

<!-- Header -->

<div style="background:#1f4e79; color:white; padding:18px; text-align:center;">
<h2 style="margin:0;">Event Seat Confirmation</h2>
<p style="margin:4px 0 0 0; font-size:14px;">Official Registration Receipt</p>
</div>

<div style="padding:25px;">

<p style="font-size:15px;">
Dear <b>${name}</b>,
</p>

<p style="font-size:14px; line-height:1.6;">
Your seat registration for the event has been <b style="color:#1aad68;">successfully confirmed</b>.
Please find your registration details below.
</p>

<!-- Details Table -->

<table style="width:100%; border-collapse:collapse; margin-top:15px; font-size:14px;">

<tr>
<td style="border:1px solid #444; padding:10px; width:40%;"><b>Student Name</b></td>
<td style="border:1px solid #444; padding:10px;">${name}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Roll Number</b></td>
<td style="border:1px solid #444; padding:10px;">${roll}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Seat Number(s)</b></td>
<td style="border:1px solid #444; padding:10px;">${bookedSeats.join(", ")}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Branch</b></td>
<td style="border:1px solid #444; padding:10px;">${branch}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Event Name</b></td>
<td style="border:1px solid #444; padding:10px;">${eventName}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Hall</b></td>
<td style="border:1px solid #444; padding:10px;">${hallName}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Organized By</b></td>
<td style="border:1px solid #444; padding:10px;">${organizedBy}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Guardian 1</b></td>
<td style="border:1px solid #444; padding:10px;">${guardian1Name} (${guardian1Phone})</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Guardian 2</b></td>
<td style="border:1px solid #444; padding:10px;">${guardian2Name || "N/A"}${guardian2Phone ? " ("+guardian2Phone+")" : ""}</td>
</tr>

</table>

<!-- QR Section -->

<div style="margin-top:30px; padding:20px; border:1px dashed #aaa; border-radius:6px; text-align:center;">

<h3 style="margin-top:0;">Event Entry QR Code</h3>

<p style="font-size:14px; color:#555;">
Please present the QR code below at the event entrance for verification.
</p>

<img src="${qrURL}&color=0-102-204" width="200" style="margin-top:10px;">

</div>

<!-- Instructions -->

<div style="margin-top:20px; font-size:13px; color:#555; line-height:1.6;">

<b>Important Instructions:</b>

<ul style="text-align:left; padding-left:18px;">
<li>Please arrive at the venue at least 15 minutes before the event.</li>
<li>Carry a valid student ID card.</li>
<li>Show this QR code at the event entry gate.</li>
<li>This email serves as your official entry confirmation.</li>
</ul>

</div>

</div>

<!-- Footer -->

<div style="background:#f0f0f0; padding:15px; text-align:center; font-size:12px; color:#666;">

<p style="margin:5px 0;">
This is an automated email confirmation generated by the event registration system.
</p>

<p style="margin:5px 0;">
© ${new Date().getFullYear()} ${organizedBy}. All rights reserved.
</p>

</div>

</div>
</div>
`;

    MailApp.sendEmail({
      to: email,
      subject: "Seat Confirmation - " + eventName,
      htmlBody: studentHTML
    });

    /* ---------------- GUARDIAN 1 EMAIL ---------------- */
    if (guardian1Email) {

const guardian1HTML = `
<div style="font-family:Arial, sans-serif; background:#f2f4f7; padding:30px;">

<div style="max-width:650px; margin:auto; background:white; border-radius:8px; overflow:hidden; box-shadow:0 6px 18px rgba(0,0,0,0.1);">

<!-- Header -->

<div style="background:#2c7a7b; color:white; padding:18px; text-align:center;">
<h2 style="margin:0;">Student Event Registration Confirmation</h2>
<p style="margin:4px 0 0 0; font-size:14px;">Notification for Guardian</p>
</div>

<div style="padding:25px; font-size:14px; line-height:1.6;">

<p>
Dear Guardian,
</p>

<p>
This email is to inform you that the following student has successfully registered for the event mentioned below.
</p>

<!-- Details Table -->

<table style="width:100%; border-collapse:collapse; margin-top:15px; font-size:14px;">

<tr>
<td style="border:1px solid #444; padding:10px; width:40%;"><b>Student Name</b></td>
<td style="border:1px solid #444; padding:10px;">${name}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Roll Number</b></td>
<td style="border:1px solid #444; padding:10px;">${roll}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Seat Number(s)</b></td>
<td style="border:1px solid #444; padding:10px;">${bookedSeats.join(", ")}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Branch</b></td>
<td style="border:1px solid #444; padding:10px;">${branch}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Event</b></td>
<td style="border:1px solid #444; padding:10px;">${eventName}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Venue / Hall</b></td>
<td style="border:1px solid #444; padding:10px;">${hallName}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Organized By</b></td>
<td style="border:1px solid #444; padding:10px;">${organizedBy}</td>
</tr>

</table>

<br>

<div style="text-align:center;">

<h3 style="margin-top:10px;">Event Entry QR Code</h3>

<p style="font-size:13px; color:#555;">
The student will present this QR code at the event entrance.
</p>

<img src="${qrURL}&color=0-102-204" width="200">

</div>

<br>

<p style="font-size:13px; color:#555;">
If you have any questions regarding this event, please contact the organizing team.
</p>

<p>
Regards,<br>
<b>${organizedBy} Team</b>
</p>

</div>

<div style="background:#f0f0f0; padding:15px; text-align:center; font-size:12px; color:#666;">
This is an automated notification from the event registration system.
</div>

</div>
</div>
`;

      MailApp.sendEmail({
        to: guardian1Email,
        subject: "Seat Booking Notification",
        htmlBody: guardian1HTML
      });
    }

    /* ---------------- GUARDIAN 2 EMAIL ---------------- */
    if (guardian2Email) {

const guardian2HTML = `
<div style="font-family:Arial, sans-serif; background:#f2f4f7; padding:30px;">

<div style="max-width:650px; margin:auto; background:white; border-radius:8px; overflow:hidden; box-shadow:0 6px 18px rgba(0,0,0,0.1);">

<!-- Header -->

<div style="background:#2c7a7b; color:white; padding:18px; text-align:center;">
<h2 style="margin:0;">Student Event Registration Confirmation</h2>
<p style="margin:4px 0 0 0; font-size:14px;">Notification for Guardian</p>
</div>

<div style="padding:25px; font-size:14px; line-height:1.6;">

<p>
Dear Guardian,
</p>

<p>
This email is to inform you that the following student has successfully registered for the event mentioned below.
</p>

<!-- Details Table -->

<table style="width:100%; border-collapse:collapse; margin-top:15px; font-size:14px;">

<tr>
<td style="border:1px solid #444; padding:10px; width:40%;"><b>Student Name</b></td>
<td style="border:1px solid #444; padding:10px;">${name}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Roll Number</b></td>
<td style="border:1px solid #444; padding:10px;">${roll}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Seat Number(s)</b></td>
<td style="border:1px solid #444; padding:10px;">${bookedSeats.join(", ")}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Branch</b></td>
<td style="border:1px solid #444; padding:10px;">${branch}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Event</b></td>
<td style="border:1px solid #444; padding:10px;">${eventName}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Venue / Hall</b></td>
<td style="border:1px solid #444; padding:10px;">${hallName}</td>
</tr>

<tr>
<td style="border:1px solid #444; padding:10px;"><b>Organized By</b></td>
<td style="border:1px solid #444; padding:10px;">${organizedBy}</td>
</tr>

</table>

<br>

<div style="text-align:center;">

<h3 style="margin-top:10px;">Event Entry QR Code</h3>

<p style="font-size:13px; color:#555;">
The student will present this QR code at the event entrance.
</p>

<img src="${qrURL}&color=0-102-204" width="200">

</div>

<br>

<p style="font-size:13px; color:#555;">
If you have any questions regarding this event, please contact the organizing team.
</p>

<p>
Regards,<br>
<b>${organizedBy} Team</b>
</p>

</div>

<div style="background:#f0f0f0; padding:15px; text-align:center; font-size:12px; color:#666;">
This is an automated notification from the event registration system.
</div>

</div>
</div>
`;

      MailApp.sendEmail({
        to: guardian2Email,
        subject: "Seat Booking Notification",
        htmlBody: guardian2HTML
      });
    }

    return {
      status: "success",
      seats: bookedSeats
    };

  } catch (err) {
    return {
      status: "error",
      message: err.message || "Booking failed"
    };
  } finally {
    lock.releaseLock();
  }
}

/* Reset seats */
function resetAllSeats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seats");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    for (let c = 4; c <= 18; c++) {
      sheet.getRange(i + 1, c).setValue("");
    }
    sheet.getRange(i + 1, 4).setValue("Available");
  }

  return "Seats Reset";
}

/* Seat Stats */
function getSeatStats() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seats");
  const data = sheet.getDataRange().getValues();

  let booked = 0;

  for (let i = 1; i < data.length; i++) {
    if (data[i][3] == "Booked") booked++;
  }

  return {
    total: data.length - 1,
    booked: booked,
    available: (data.length - 1) - booked
  };
}

/* CSV Export */
function exportSeatData() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seats");
  const data = sheet.getDataRange().getValues();

  return data.map(r => r.join(",")).join("\n");
}

/* QR Entry Verification */
function verifyEntry(roll, seat) {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Seats");
  const data = sheet.getDataRange().getValues();

  for (let i = 1; i < data.length; i++) {
    if (data[i][4] == roll && data[i][0] == seat) {
      sheet.getRange(i + 1, 13).setValue("Checked In");
      return "Entry Verified";
    }
  }

  return "Invalid Ticket";
}
