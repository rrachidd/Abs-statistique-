import express from "express";
import { createServer as createViteServer } from "vite";
import path from "path";
import { google } from "googleapis";
import dotenv from "dotenv";

dotenv.config();

const app = express();
const PORT = 3000;

app.use(express.json());

// Google Sheets Setup
const auth = new google.auth.GoogleAuth({
  credentials: {
    client_email: process.env.GOOGLE_SERVICE_ACCOUNT_EMAIL,
    private_key: process.env.GOOGLE_PRIVATE_KEY?.replace(/\\n/g, "\n"),
  },
  scopes: ["https://www.googleapis.com/auth/spreadsheets"],
});

const sheets = google.sheets({ version: "v4", auth });
const SPREADSHEET_ID = process.env.GOOGLE_SHEET_ID;

// API Routes
app.get("/api/sheets/data", async (req, res) => {
  try {
    if (!SPREADSHEET_ID) {
      return res.status(400).json({ error: "GOOGLE_SHEET_ID is not configured" });
    }

    const response = await sheets.spreadsheets.values.get({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sheet1!A:Z", // Adjust range as needed
    });

    const rows = response.data.values || [];
    res.json({ rows });
  } catch (error: any) {
    console.error("Error fetching data:", error);
    res.status(500).json({ error: error.message });
  }
});

app.post("/api/sheets/data", async (req, res) => {
  try {
    const { values } = req.body;
    if (!SPREADSHEET_ID) {
      return res.status(400).json({ error: "GOOGLE_SHEET_ID is not configured" });
    }

    await sheets.spreadsheets.values.append({
      spreadsheetId: SPREADSHEET_ID,
      range: "Sheet1!A1",
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [values],
      },
    });

    res.json({ success: true });
  } catch (error: any) {
    console.error("Error appending data:", error);
    res.status(500).json({ error: error.message });
  }
});

app.put("/api/sheets/data/:row", async (req, res) => {
  try {
    const { row } = req.params;
    const { values } = req.body;
    if (!SPREADSHEET_ID) {
      return res.status(400).json({ error: "GOOGLE_SHEET_ID is not configured" });
    }

    await sheets.spreadsheets.values.update({
      spreadsheetId: SPREADSHEET_ID,
      range: `Sheet1!A${row}`,
      valueInputOption: "USER_ENTERED",
      requestBody: {
        values: [values],
      },
    });

    res.json({ success: true });
  } catch (error: any) {
    console.error("Error updating data:", error);
    res.status(500).json({ error: error.message });
  }
});

// Vite middleware setup
async function startServer() {
  if (process.env.NODE_ENV !== "production") {
    const vite = await createViteServer({
      server: { middlewareMode: true },
      appType: "spa",
    });
    app.use(vite.middlewares);
  } else {
    const distPath = path.join(process.cwd(), "dist");
    app.use(express.static(distPath));
    app.get("*", (req, res) => {
      res.sendFile(path.join(distPath, "index.html"));
    });
  }

  app.listen(PORT, "0.0.0.0", () => {
    console.log(`Server running on http://localhost:${PORT}`);
  });
}

startServer();
