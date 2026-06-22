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

function nowTime(tz) {
  return new Date().toLocaleString("en-US", {
    timeZone: tz, hour: "numeric", minute: "2-digit", hour12: true,
  });
}

function nowDate(tz) {
  return new Date().toLocaleDateString("en-US", {
    timeZone: tz, month: "2-digit", day: "2-digit", year: "numeric",
  });
}

async function logToSheet(type) {
  const sheets = google.sheets({ version: "v4", auth });
  const tz     = process.env.TIMEZONE || "America/New_York";
  const time   = nowTime(tz);
  const date   = nowDate(tz);

  if (type === "IN") {
    await sheets.spreadsheets.values.append({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "Clock Log!A:E",
      valueInputOption: "USER_ENTERED",
      requestBody: { values: [[date, "", time, "", ""]] },
    });
    return `✅ Clocked in at ${time}`;
  }

  if (type === "OUT") {
    const res  = await sheets.spreadsheets.values.get({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
      range: "Clock Log!A:E",
    });
    const rows = res.data.values || [];

    let targetRow = -1;
    for (let i = rows.length - 1; i >= 0; i--) {
      const hasTimeIn = rows[i][2] && rows[i][2].trim() !== "";
      const noTimeOut = !rows[i][3] || rows[i][3].trim() === "";
      if (hasTimeIn && noTimeOut) { targetRow = i + 1; break; }
    }

    if (targetRow === -1) {
      return "⚠️ No open clock-in found. Send IN first!";
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: process.env.GOOGLE_SHEET_ID,
