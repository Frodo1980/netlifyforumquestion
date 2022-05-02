const express = require("express");
const XlsxPopulate = require("xlsx-populate");
var cors = require("cors");
const serverless = require("serverless-http");

const app = express();
const router = express.Router();
app.use(cors());
app.options("*", cors());

router.get("/", (req, res) => {
  console.log(req);
  // Load an existing workbook
  XlsxPopulate.fromFileAsync("ExcelBlanko.xlsx")
    .then((workbook) => {
      const worksheet = workbook.sheet("Kalkulationstool");

      const params = req.query;

      //repair types
      (params.D7 = parseInt(params.D7)),
        (params.D8 = parseInt(params.D8)),
        (params.D9 = parseInt(params.D9)),
        (params.D10 = parseFloat(params.D10)),
        (params.D11 = parseFloat(params.D11)),
        (params.D12 = parseFloat(params.D12)),
        (params.D14 = parseInt(params.D14));

      console.table(params);

      // Modify the workbook.
      Object.keys(params).map((key, index) => {
        worksheet.cell(key).value(Object.values(params)[index]);
      });

      return workbook.outputAsync("base64");
    })
    .then((data) => {
      // Send the workbook as base64-string
      res.send(data);
    });
});

app.use("/.netlify/functions/api", router);

module.exports.handler = serverless(app);
