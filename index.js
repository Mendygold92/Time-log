import express from "express";
import { google } from "googleapis";
import twilio from "twilio";
import dotenv from "dotenv";

dotenv.config();

const app = express();
app.use(express.urlencoded({ extended: false }));
app.use(express.json());

const auth = new google.auth.GoogleAuth({
  credentials: JSON.parse(process.env.GOOGLE_CREDENTIALS_JSON),
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

// Parse "IN | Acme Corp | Task | Creative | Yes | Notes"
function parseMessage(text) {
  const parts = text.split("|").map(s => s.trim());
  const type = parts[0]?.toUpperCase(); // IN or OUT
  return {
    type,
    client:    parts[1] || "",
    task:      parts[2] || "",
    workType:  parts[3] || "",
    billable:  parts[4] || "",
    notes:     parts[5] || "",
  };
}

function nowFormatted(tz) {
  return new Date().toLocaleString("en-US", {
    timeZone: tz,
    hour: "numeric",
    minute: "2-digit",
    hour12: true,
  });
}

function dateFormatted(tz) {
  return new Date().toLocaleDateString("en-US", {
    timeZone: tz,
    month: "2-digit",
    day: "2-digit",
    year: "numeric",
  });
}

async function logToSheet(parsed) {
  const sheets = google.sheets({ version: "v4", auth });
  const tz     = process.env.TIMEZONE || "America/New_York";
  const time   = nowFormatted(tz);
  const date   = dateFormatted(tz);

  // Columns: Date | Day (formula) | Time In | Time Out | Total Hours (formula) | Client | Task | Work Type | Billable | Notes
  // We write: A=date, C or D=time, F=client, G=task, H=workType, I=billable, J=notes
  // B (Day) and E (Total Hours) are formulas already in the sheet

  if (parsed.type === "IN") {
    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "Time Log!A:J",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [[
          date,           // A - Date
          "",             // B - Day (formula in sheet)
          time,           // C - Time In
          "",             // D - Time Out (fill later with OUT)
          "",             // E - Total Hours (formula)
          parsed.client,  // F
          parsed.task,    // G
          parsed.workType,// H
          parsed.billable,// I
          parsed.notes,   // J
        ]],
      },
    });
    return `✅ Clocked IN at ${time}`;
  }

  if (parsed.type === "OUT") {
    // Find the last row with this client that has no Time Out yet
    const res = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "Time Log!A:J",
    });
    const rows = res.data.values || [];
    // Find last row where col C (Time In) has a value and col D (Time Out) is empty
    let targetRow = -1;
    for (let i = rows.length - 1; i >= 0; i--) {
      const row = rows[i];
      const hasTimeIn  = row[2] && row[2].trim() !== "";
      const noTimeOut  = !row[3] || row[3].trim() === "";
      const clientMatch = !parsed.client || (row[5] || "").toLowerCase().includes(parsed.client.toLowerCase());
      if (hasTimeIn && noTimeOut && clientMatch) {
        targetRow = i + 1; // 1-indexed for Sheets
        break;
      }
    }

    if (targetRow === -1) {
      return "⚠️ Couldn't find an open clock-in for that client. Did you send IN first?";
    }

    // Update Time Out in col D (index 4 = column D)
    await sheets.spreadsheets.values.update({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: `Time Log!D${targetRow}`,
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[time]] },
    });

    // Optionally update notes if provided
    if (parsed.notes) {
      await sheets.spreadsheets.values.update({
        spreadsheetId: process.env.GOOGLE_SHEET_ID,
        range: `Time Log!J${targetRow}`,
        valueInputOption: "USER_ENTERED",
        requestBody: { values: [[parsed.notes]] },
      });
    }

    return `✅ Clocked OUT at ${time} — row ${targetRow} updated`;
  }

  return `⚠️ Start your message with IN or OUT.\nExample: IN | Acme Corp | Task | Creative | Yes`;
}

app.post("/webhook", async (req, res) => {
  const body = (req.body.Body || "").trim();
  console.log("📩 Received:", body);

  const twiml = new twilio.twiml.MessagingResponse();
  try {
    const parsed = parseMessage(body);
    const reply  = await logToSheet(parsed);
    twiml.message(reply);
  } catch (err) {
    console.error("❌ Error:", err.message);
    twiml.message("⚠️ Something went wrong. Check your server logs.");
  }

  res.type("text/xml").send(twiml.toString());
});

app.get("/", (req, res) => res.send("WhatsApp logger running ✅"));

const PORT = process.env.PORT || 3000;
app.listen(PORT, () => console.log(`🚀 Listening on port ${PORT}`));
