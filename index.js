const express = require("express");
const randomstring = require("randomstring");
const XLSX = require("xlsx");

const app = express();

app.listen(3000, () => {
    console.log("Server started on port 3000");
  });

app.use(express.json());
app.use(express.urlencoded({ extended: true }));

app.get("/", (req, res) => {
  res.sendFile(__dirname + "/index.html");
});

app.post("/generate-codes", (req, res) => {
  // Basic options
  const filename = req.body.filename || "nkzm_random_code";
  const count = parseInt(req.body.count) || 100;
  const codelength = parseInt(req.body.codelength) || 9;
  // Advance options
  const prefix = req.body.prefix || "";
  const suffix = req.body.suffix || "";
  const charset = req.body.charset || "";

  const _data = [];
  let i = 0;

  while (i < count) {
    if (charset=="") {
      _ran = prefix + randomstring.generate(codelength) + suffix;
    }
    if (charset!=="") {
      _ran = prefix + randomstring.generate({
        length: codelength,
        charset: charset
      }) + suffix;
    }
    
    const _code = { code: _ran };

    if (!_data.includes(_code)) {
      _data.push(_code);
      i++;
    }
  }

  const workSheet = XLSX.utils.json_to_sheet(_data);
  const workBook = XLSX.utils.book_new();

  XLSX.utils.book_append_sheet(workBook, workSheet, "Sheet 1");

  // Generate the Excel file in memory
  const excelFile = XLSX.write(workBook, { type: "buffer", bookType: "xlsx" });

  // Set appropriate headers for the response
  res.setHeader("Content-Disposition", "attachment; filename="+filename+".xlsx");
  res.setHeader("Content-Type", "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");

  // Send the Excel file as a response
  res.send(excelFile);
});

module.exports = app;
