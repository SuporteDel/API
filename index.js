const fs = require("fs");
const { google } = require("googleapis");
const { ConfidentialClientApplication } = require("@azure/msal-node");
const { Document, Packer, Paragraph, HeadingLevel } = require("docx");

const SPREADSHEET_ID = "1R1SPOcmvqElPUjSJHgQLG7h9vdW05bNgxgRVSDK_A7k";
const RANGE = "MT!L:L"; // aba MT, coluna L
const STATE_FILE = "processed.json";

function loadState() {
  if (!fs.existsSync(STATE_FILE)) return { rows: [] };
  return JSON.parse(fs.readFileSync(STATE_FILE, "utf8"));
}

function saveState(state) {
  fs.writeFileSync(STATE_FILE, JSON.stringify(state, null, 2));
}

async function getSheetRows() {
  const auth = new google.auth.GoogleAuth({
    credentials: JSON.parse(process.env.GOOGLE_SERVICE_ACCOUNT_JSON),
    scopes: ["https://www.googleapis.com/auth/spreadsheets.readonly"]
  });

  const sheets = google.sheets({ version: "v4", auth });
  const res = await sheets.spreadsheets.values.get({
    spreadsheetId: SPREADSHEET_ID,
    range: RANGE
  });

  return res.data.values || [];
}

async function createWordBuffer(texto, linha) {
  const doc = new Document({
    sections: [{
      children: [
        new Paragraph({
          text: "Audiência",
          heading: HeadingLevel.HEADING_1
        }),
        new Paragraph(`Linha da planilha: ${linha}`),
        new Paragraph(""),
        new Paragraph(texto)
      ]
    }]
  });

  return await Packer.toBuffer(doc);
}

async function getMicrosoftToken() {
  const app = new ConfidentialClientApplication({
    auth: {
      clientId: process.env.MS_CLIENT_ID,
      authority: `https://login.microsoftonline.com/${process.env.MS_TENANT_ID}`,
      clientSecret: process.env.MS_CLIENT_SECRET
    }
  });

  const result = await app.acquireTokenByClientCredential({
    scopes: ["https://graph.microsoft.com/.default"]
  });

  return result.accessToken;
}

async function uploadToOneDrive(filename, buffer) {
  const token = await getMicrosoftToken();

  const path = `/Audiencias/${filename}`;

  const res = await fetch(
    `https://graph.microsoft.com/v1.0/users/${process.env.MS_USER_ID}/drive/root:${path}:/content`,
    {
      method: "PUT",
      headers: {
        Authorization: `Bearer ${token}`,
        "Content-Type": "application/vnd.openxmlformats-officedocument.wordprocessingml.document"
      },
      body: buffer
    }
  );

  if (!res.ok) {
    throw new Error(await res.text());
  }

  return await res.json();
}

async function main() {
  const state = loadState();
  const rows = await getSheetRows();

  for (let i = 1; i < rows.length; i++) {
    const linha = i + 1;
    const texto = rows[i]?.[0];

    if (!texto || !texto.trim()) continue;
    if (state.rows.includes(linha)) continue;

    const filename = `audiencia_linha_${linha}.docx`;
    const buffer = await createWordBuffer(texto, linha);

    await uploadToOneDrive(filename, buffer);

    state.rows.push(linha);
    console.log(`Gerado: ${filename}`);
  }

  saveState(state);
}

main().catch(err => {
  console.error(err);
  process.exit(1);
});
