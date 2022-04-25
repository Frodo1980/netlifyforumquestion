const express = require("express");
const XlsxPopulate = require("xlsx-populate");

const serverless = require("serverless-http");

const app = express();
const router = express.Router();

router.get("/", (req, res) => {
  // Load an existing workbook
  XlsxPopulate.fromFileAsync("file_example.xlsx")
    .then((workbook) => {
      const worksheet = workbook.sheet("Sheet1");

      const params = req.query;

      // Modify the workbook.
      Object.keys(params).map((key, index) => {
        console.log(Object.values(params)[index]);
        worksheet.cell(key).value(Object.values(params)[index]);
      });

      return workbook.outputAsync();
    })
    .then((data) => {
      // Set the output file name.
      res.attachment("output.xlsx");

      // Send the workbook.
      res.send(data);
    });
});

app.use("/.netlify/functions/api", router);

module.exports.handler = serverless(app);
