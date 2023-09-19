const functions = require("firebase-functions");
const express = require("express");
const admin = require("firebase-admin");
const xlsx = require("xlsx");

admin.initializeApp();

const app = express();

// Use express.json() middleware to parse JSON request bodies
app.use(express.json());

app.post("/csvtoexcel", async (req, res) => {
  console.log("Received request with headers:", req.headers);
  
  const data = req.body.data;
  let filename = req.body.filename || `converted-${Date.now()}`;
  filename = filename + `.xlsx`;  // Add the extension

  console.log("Received data:", data);

  console.log("Starting Excel generation...");
  const workbook = xlsx.utils.book_new();
  const worksheet = xlsx.utils.aoa_to_sheet(data);
  xlsx.utils.book_append_sheet(workbook, worksheet, "Sheet 1");

  const buffer = xlsx.write(workbook, { type: "buffer", bookType: "xlsx" });
  console.log("Finished Excel generation. Starting upload to Firebase Storage...");

  const bucket = admin.storage().bucket();
  const file = bucket.file(filename);
  const stream = file.createWriteStream({
      metadata: {
          contentType: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet'
      }
  });

  stream.on('error', (err) => {
      console.error("Error during Firebase Storage upload:", err);
      res.status(500).end();
  });

  stream.on('finish', async () => {
      console.log("Finished uploading to Firebase Storage. Making file public...");
      await file.makePublic();
      const url = `https://storage.googleapis.com/${bucket.name}/${filename}`;
      console.log("File is now public. Sending response with URL:", url);
      res.status(200).send({ url });
  });

  stream.end(buffer);
});


app.get("/listexcel", async (req, res) => {
  const bucket = admin.storage().bucket();
  const [files] = await bucket.getFiles();

  const fileNames = files.map(file => file.name);
  res.status(200).send({ fileNames });
});

app.get("/exceltoarray/:filename", async (req, res) => {
  const { filename } = req.params;
  const bucket = admin.storage().bucket();
  const file = bucket.file(filename);

  const buffer = await file.download();
  const workbook = xlsx.read(buffer[0], { type: "buffer" });
  const worksheet = workbook.Sheets[workbook.SheetNames[0]];
  let data = xlsx.utils.sheet_to_json(worksheet, { header: 1 });

  // Convert every cell to string
  data = data.map(row => row.map(cell => String(cell)));

  res.status(200).send({ data });
});

console.log("Starting Firebase Functions...");

exports.api = functions.https.onRequest(app);
