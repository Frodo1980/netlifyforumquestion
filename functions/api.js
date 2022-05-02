const express = require("express");
const XlsxPopulate = require("xlsx-populate");
var cors = require("cors");
const serverless = require("serverless-http");

const app = express();
const router = express.Router();
app.use(cors());
app.options("*", cors());

router.get("/", (req, res) => {
  // Load an existing workbook
  XlsxPopulate.fromFileAsync("file_example.xlsx")
    .then((workbook) => {
      const worksheet = workbook.sheet("Sheet1");

      const params = req.query;
      console.table(params);

      // Modify the workbook.
      Object.keys(params).map((key, index) => {
        worksheet.cell(key).value(Object.values(params)[index]);
      });

      return workbook.outputAsync("base64");
    })
    .then((data) => {
      // Set the output file name.
      //  res.attachment("output.xlsx");

      // Send the workbook.
      res.send(data);
    });
});

app.use("/.netlify/functions/api", router);

module.exports.handler = serverless(app);
