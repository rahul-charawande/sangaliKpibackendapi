// server.js
require("dotenv").config();
const express = require("express");
const cors = require("cors");
const bodyParser = require("body-parser");
const { google } = require("googleapis");
const xlsx = require("xlsx");
const axios = require("axios");

const app = express();
app.use(cors());
app.use(bodyParser.json());

// 🔑 Google Drive Auth
const auth = new google.auth.GoogleAuth({
  credentials: {
    type: process.env.GOOGLE_TYPE,
    project_id: process.env.GOOGLE_PROJECT_ID,
    private_key_id: process.env.GOOGLE_PRIVATE_KEY_ID,
    private_key: process.env.GOOGLE_PRIVATE_KEY.replace(/\\n/g, "\n"),
    client_email: process.env.GOOGLE_CLIENT_EMAIL,
    client_id: process.env.GOOGLE_CLIENT_ID,
    auth_uri: process.env.GOOGLE_AUTH_URI,
    token_uri: process.env.GOOGLE_TOKEN_URI,
    auth_provider_x509_cert_url: process.env.GOOGLE_AUTH_PROVIDER_X509_CERT_URL,
    client_x509_cert_url: process.env.GOOGLE_CLIENT_X509_CERT_URL,
    universe_domain: process.env.GOOGLE_UNIVERSE_DOMAIN,
  },
  scopes: ["https://www.googleapis.com/auth/drive.readonly"],
});
const drive = google.drive({ version: "v3", auth });

/* ------------------------------------------------------
   Utility Functions
------------------------------------------------------ */
const summarizeSheetBasic = (sheetJson) => {
  if (!sheetJson || sheetJson.length === 0) return { rowCount: 0, columns: [] };
  return { rowCount: sheetJson.length, columns: Object.keys(sheetJson[0]) };
};

// Utility: summarize sheet
const summarizeSheet = (sheetJson) => {
  if (!sheetJson || sheetJson.length === 0) return { rowCount: 0, columns: [] };
  const columns = Object.keys(sheetJson[0]);
  return { rowCount: sheetJson.length, columns };
};

const summarizeSheetDetailed = (sheetJson) => {
  if (!sheetJson || sheetJson.length === 0) return { rowCount: 0, columns: [] };
  return {
    rowCount: sheetJson.length,
    columns: Object.keys(sheetJson[0]).map((col) => {
      const values = sheetJson.map((row) => row[col]);
      const numericValues = values.filter((v) => typeof v === "number");
      const uniqueValues = [...new Set(values.filter((v) => v != null))];
      const missingCount = values.filter((v) => v == null).length;
      return {
        column: col,
        type: numericValues.length === values.length ? "numeric" : "categorical",
        missing: missingCount,
        uniqueCount: uniqueValues.length,
        sum: numericValues.length ? numericValues.reduce((a, b) => a + b, 0) : null,
        avg: numericValues.length ? numericValues.reduce((a, b) => a + b, 0) / numericValues.length : null,
      };
    }),
  };
};

const summarizeSheetSamples = (sheetJson) => {
  if (!sheetJson || sheetJson.length === 0) return { rowCount: 0, columns: [] };
  return {
    rowCount: sheetJson.length,
    columns: Object.keys(sheetJson[0]).map((col) => ({
      column: col,
      sampleValues: sheetJson.slice(0, 5).map((row) => row[col]),
    })),
  };
};

/* ------------------------------------------------------
   1. Final Summary API (/api/final-summary)
------------------------------------------------------ */
app.post("/api/final-summary", async (req, res) => {
   try {
      const { url } = req.body;
      if (!url) return res.status(400).json({ error: "Google Drive URL required" });
  
      const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) return res.status(400).json({ error: "Invalid Google Drive URL" });
      const fileId = match[1];
  
      const file = await drive.files.get({ fileId, fields: "id, name, mimeType" });
  
      let buffer;
      if (file.data.mimeType === "application/vnd.google-apps.spreadsheet") {
        const resExport = await drive.files.export(
          { fileId, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
          { responseType: "arraybuffer" }
        );
        buffer = Buffer.from(resExport.data);
      } else {
        const resDownload = await drive.files.get({ fileId, alt: "media" }, { responseType: "arraybuffer" });
        buffer = Buffer.from(resDownload.data);
      }
  
      const workbook = xlsx.read(buffer, { type: "buffer" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      let json = xlsx.utils.sheet_to_json(sheet, { defval: null }).slice(0, 10);
  
      const sheetSummary = summarizeSheet(json);
  
      const prompt = `
        You are a data assistant.
        Given this Excel summary: ${JSON.stringify(sheetSummary)}
        Generate a JSON object strictly with this format:
        {
          "limitations": [{"id": number, "text": "string", "fixed": boolean}],
          "takeaway": "string"
        }
        - "limitations" should be 3-6 items max
        - Each item must have id, text, fixed (true/false)
        - takeaway should be 1 actionable insight.
        Return ONLY valid JSON.
      `;
  
      const gptResponse = await axios.post(
        "https://api.openai.com/v1/chat/completions",
        {
          model: "gpt-4.1-mini",
          messages: [{ role: "user", content: prompt }],
          temperature: 0.3,
        },
        { headers: { Authorization: `Bearer ${process.env.OPENAI_API_KEY}` } }
      );
  
      let rawContent = gptResponse.data.choices[0].message.content.trim();
      rawContent = rawContent.replace(/```json|```/g, "").trim();
  
      let structuredData;
      try {
        structuredData = JSON.parse(rawContent);
      } catch (e) {
        console.error("❌ JSON Parse Error:", e.message, rawContent);
        return res.status(500).json({ error: "Invalid JSON from GPT" });
      }
  
      res.json(structuredData);
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: "Failed to fetch Final Summary" });
    }
});

/* ------------------------------------------------------
   2. GPT Analytics API (/api/gpt)
------------------------------------------------------ */
app.post("/api/gpt", async (req, res) => {
    try {
          const { url, testMode } = req.body; // add a testMode flag
          if (!url) return res.status(400).json({ error: "Google Drive URL required" });
  
          const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
          if (!match) return res.status(400).json({ error: "Invalid Google Drive URL" });
          const fileId = match[1];
  
          const file = await drive.files.get({ fileId, fields: "id, name, mimeType" });
  
          let buffer;
          if (file.data.mimeType === "application/vnd.google-apps.spreadsheet") {
              const resExport = await drive.files.export(
                  { fileId, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
                  { responseType: "arraybuffer" }
              );
              buffer = Buffer.from(resExport.data);
          } else {
              const resDownload = await drive.files.get({ fileId, alt: "media" }, { responseType: "arraybuffer" });
              buffer = Buffer.from(resDownload.data);
          }
  
          const workbook = xlsx.read(buffer, { type: "buffer" });
  
          const sheetsSummarya = {};
          workbook.SheetNames.forEach((sheetName) => {
              const sheet = workbook.Sheets[sheetName];
              let json = xlsx.utils.sheet_to_json(sheet, { defval: null });
  
              // TEST MODE: Limit rows to 50 only
              if (testMode) json = json.slice(0, 8);
  
              sheetsSummarya[sheetName] = summarizeSheet(json);
  
          });
  
          const sheetsSummary = {};
          // Take only the first 2 sheets
          workbook.SheetNames.slice(0, 2).forEach((sheetName) => {
              const sheet = workbook.Sheets[sheetName];
              let json = xlsx.utils.sheet_to_json(sheet, { defval: null });
  
              // Limit rows to first 10
              json = json.slice(0, 5);
  
              sheetsSummary[sheetName] = summarizeSheet(json);
          });
  
          console.log("sheetSummary : ", sheetsSummary);
          const prompt = `
        You are an analytics assistant. 
        Given these Excel sheet summaries: ${JSON.stringify(sheetsSummary)} 
        Generate a JSON object with the structure:
        {
          "kpis": [{"label": "string", "value": "string"}],
          "executiveSummary": ["string"],
          "dataQualityNotes": ["string"]
        }
        Return only JSON.
      `;
  
          const gptResponse = await axios.post(
              "https://api.openai.com/v1/chat/completions",
              {
                  model: "gpt-4.1-mini",
                  messages: [{ role: "user", content: prompt }],
                  temperature: 0.3,
              },
              { headers: { Authorization: `Bearer ${process.env.OPENAI_API_KEY}` } }
          );
  
          // ✅ Fix: Strip ```json fences before parsing
          let rawContent = gptResponse.data.choices[0].message.content.trim();
          rawContent = rawContent.replace(/```json|```/g, "").trim();
  
          let structuredData;
          try {
              structuredData = JSON.parse(rawContent);
          } catch (e) {
              console.error("❌ JSON Parse Error:", e.message, rawContent);
              return res.status(500).json({ error: "Invalid JSON from GPT" });
          }
  
          res.json(structuredData);
      } catch (err) {
          console.error(err);
          res.status(500).json({ error: "Failed to fetch Excel sheet data" });
      }
});

/* ------------------------------------------------------
   3. Recommendations API (/api/recommendations)
------------------------------------------------------ */
app.post("/api/recommendations", async (req, res) => {
  try {
      const { url, testMode } = req.body;
      if (!url) return res.status(400).json({ error: "Google Drive URL required" });
  
      const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) return res.status(400).json({ error: "Invalid Google Drive URL" });
      const fileId = match[1];
  
      const file = await drive.files.get({ fileId, fields: "id, name, mimeType" });
  
      let buffer;
      if (file.data.mimeType === "application/vnd.google-apps.spreadsheet") {
        const resExport = await drive.files.export(
          { fileId, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
          { responseType: "arraybuffer" }
        );
        buffer = Buffer.from(resExport.data);
      } else {
        const resDownload = await drive.files.get({ fileId, alt: "media" }, { responseType: "arraybuffer" });
        buffer = Buffer.from(resDownload.data);
      }
  
      const workbook = xlsx.read(buffer, { type: "buffer" });
  
      const sheetsSummary = {};
      workbook.SheetNames.slice(0, 2).forEach((sheetName) => {
        const sheet = workbook.Sheets[sheetName];
        let json = xlsx.utils.sheet_to_json(sheet, { defval: null });
  
        if (testMode) json = json.slice(0, 10);
  
        sheetsSummary[sheetName] = summarizeSheet(json);
      });
  
      const prompt = `
        You are a city analytics assistant.
        Based on these data summaries: ${JSON.stringify(sheetsSummary)}
        Generate recommendations in strictly valid JSON:
        {
          "recommendations": {
            "immediate": [{"id": 1, "text": "string", "done": false}],
            "medium": [{"id": 1, "text": "string", "done": false}],
            "strategic": [{"id": 1, "text": "string", "done": false}]
          },
          "monitoringPlan": [
            {"id": 1, "label": "Daily", "desc": "string"},
            {"id": 2, "label": "Weekly", "desc": "string"},
            {"id": 3, "label": "Monthly", "desc": "string"}
          ]
        }
        Do not include explanations or extra text.
      `;
  
      const gptResponse = await axios.post(
        "https://api.openai.com/v1/chat/completions",
        {
          model: "gpt-4.1-mini",
          messages: [{ role: "user", content: prompt }],
          temperature: 0.3,
        },
        { headers: { Authorization: `Bearer ${process.env.OPENAI_API_KEY}` } }
      );
  
      let rawContent = gptResponse.data.choices[0].message.content.trim();
      rawContent = rawContent.replace(/```json|```/g, "").trim();
  
      let structuredData;
      try {
        structuredData = JSON.parse(rawContent);
      } catch (e) {
        console.error("❌ JSON Parse Error:", e.message, rawContent);
        return res.status(500).json({ error: "Invalid JSON from GPT" });
      }
  
      res.json(structuredData);
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: "Failed to fetch recommendations" });
    }
});

/* ------------------------------------------------------
   4. Fetch Excel API (/api/fetch-excel)
------------------------------------------------------ */
app.post("/api/fetch-excel", async (req, res) => {
  try {
    const { url, sheetIndex = 0 } = req.body; // sheetIndex from client
    if (!url) return res.status(400).json({ error: "Google Drive URL required" });

    // Extract fileId from URL
    const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
    if (!match) return res.status(400).json({ error: "Invalid Google Drive URL" });
    const fileId = match[1];

    // Get file metadata
    const file = await drive.files.get({ fileId, fields: "id, name, mimeType" });

    let buffer;
    if (file.data.mimeType === "application/vnd.google-apps.spreadsheet") {
      // Export Google Sheet as XLSX in memory
      const resExport = await drive.files.export(
        { fileId, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
        { responseType: "arraybuffer" }
      );
      buffer = Buffer.from(resExport.data);
    } else if (
      file.data.mimeType === "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" ||
      file.data.mimeType === "application/vnd.ms-excel"
    ) {
      // Direct Excel download
      const resDownload = await drive.files.get({ fileId, alt: "media" }, { responseType: "arraybuffer" });
      buffer = Buffer.from(resDownload.data);
    } else {
      return res.status(400).json({ error: "Not a Google Sheet or Excel file" });
    }

    // Read Excel into workbook
    const workbook = xlsx.read(buffer, { type: "buffer" });
    const sheetName = workbook.SheetNames[sheetIndex];
    if (!sheetName) return res.status(400).json({ error: "Invalid sheet index" });

    // Return raw sheet data as array of arrays (not converting to JSON keys/values)
    const sheet = workbook.Sheets[sheetName];
    //const rawData = xlsx.utils.sheet_to_json(sheet, { header: 1 }); // array of arrays

    // With the improved version:
    const rawData = [];
    const range = xlsx.utils.decode_range(sheet['!ref']);
    for (let R = range.s.r; R <= range.e.r; ++R) {
      const row = [];
      for (let C = range.s.c; C <= range.e.c; ++C) {
        const cellAddress = xlsx.utils.encode_cell({ r: R, c: C });
        const cell = sheet[cellAddress];
        row.push({ cell: cellAddress, value: cell ? cell.v : null });
      }
      rawData.push(row);
    }


    return res.json({ sheetName, data: rawData });
  } catch (err) {
    console.error("❌ API Error:", err.message);
    return res.status(500).json({ error: "Failed to fetch Excel sheet data" });
  }
});

/* ------------------------------------------------------
   5. Alerts API (/api/gpt-alerts)
------------------------------------------------------ */
app.post("/api/gpt-alerts", async (req, res) => {
   try {
      const { url } = req.body;
      if (!url) return res.status(400).json({ error: "Google Drive URL required" });
  
      const match = url.match(/\/d\/([a-zA-Z0-9-_]+)/);
      if (!match) return res.status(400).json({ error: "Invalid Google Drive URL" });
      const fileId = match[1];
  
      const file = await drive.files.get({ fileId, fields: "id, mimeType" });
  
      let buffer;
      if (file.data.mimeType === "application/vnd.google-apps.spreadsheet") {
        const resExport = await drive.files.export(
          { fileId, mimeType: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet" },
          { responseType: "arraybuffer" }
        );
        buffer = Buffer.from(resExport.data);
      } else {
        const resDownload = await drive.files.get({ fileId, alt: "media" }, { responseType: "arraybuffer" });
        buffer = Buffer.from(resDownload.data);
      }
  
      const workbook = xlsx.read(buffer, { type: "buffer" });
      const sheet = workbook.Sheets[workbook.SheetNames[0]];
      const json = xlsx.utils.sheet_to_json(sheet, { defval: null }).slice(0, 20);
  
      const summary = summarizeSheet(json);
  
      // Prompt GPT
      const prompt = `
        You are a data insights assistant.
        Given this sheet summary: ${JSON.stringify(summary)}
        Generate a JSON object with the structure:
        {
          "alerts": [
            {"id": 1, "type": "critical|warning", "text": "string"}
          ],
          "correlations": [
            {"id": 1, "pair": "string", "value": number, "meaning": "string", "trend": "up|down"}
          ]
        }
        Return only valid JSON.
      `;
  
      const gptResponse = await axios.post(
        "https://api.openai.com/v1/chat/completions",
        {
          model: "gpt-4.1-mini",
          messages: [{ role: "user", content: prompt }],
          temperature: 0.3,
        },
        { headers: { Authorization: `Bearer ${process.env.OPENAI_API_KEY}` } }
      );
  
      let rawContent = gptResponse.data.choices[0].message.content.trim();
      rawContent = rawContent.replace(/```json|```/g, "").trim();
  
      let structuredData;
      try {
        structuredData = JSON.parse(rawContent);
      } catch (e) {
        console.error("❌ JSON Parse Error:", e.message, rawContent);
        return res.status(500).json({ error: "Invalid JSON from GPT" });
      }
  
      res.json(structuredData);
    } catch (err) {
      console.error(err);
      res.status(500).json({ error: "Failed to fetch alerts and correlations" });
    }


});

/* ------------------------------------------------------
   Start Server (single port)
------------------------------------------------------ */
const PORT = process.env.PORT || 5000;
app.listen(PORT, () => console.log(`🚀 Unified API running on port ${PORT}`));
