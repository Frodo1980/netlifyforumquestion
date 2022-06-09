const express = require("express");
const XlsxPopulate = require("xlsx-populate");
var cors = require("cors");
const serverless = require("serverless-http");

const app = express();
const router = express.Router();

//CORS im Express-Server einrichten
app.use(cors());
app.options("*", cors());

router.get("/", (req, res) => {
  // Load an existing workbook
  XlsxPopulate.fromFileAsync("ExcelBlanko.xlsx")
    .then((workbook) => {
      //Nur ausführen, wenn "Passwort" korrekt
      //TODO Hier evtl. mal eine vernünftige Sicherheitslösung implementieren

      const kalkulationstool = workbook.sheet("Kalkulationstool");
      const wertbeitrag = workbook.sheet("Wertbeitrag");
      const bewertungAufstockung = workbook.sheet("Bewertung Aufstockung");

      const params = req.query;

      //Datentypen anpassen
      params.D7 = parseFloat(params.D7);
      params.D8 = parseInt(params.D8);
      params.D9 = parseInt(params.D9);
      params.D10 = parseFloat(params.D10);
      if (params.D11) {
        params.D11 = parseFloat(params.D11);
      } else {
        params.D11 = 0;
      }
      if (params.D12) {
        params.D12 = parseFloat(params.D12);
      } else {
        params.D12 = 0;
      }
      params.D14 = parseInt(params.D14);

      console.table(params);

      // Arbeitsblatt anpassen
      Object.keys(params).map((key, index) => {
        // der Parameter "key" dient nur als Voraussetzung zum Ausführen der Funktion
        if (key != "key") {
          // Alle auf dem Tabellenblatt "Kalkulationstool" zu ändernden Zellen befinden sich in Spalte "D"
          if (key.charAt(0) === "D") {
            kalkulationstool.cell(key).value(Object.values(params)[index]);
          }
          // Alle auf dem Tabellenblatt "Wertbeitrag" zu ändernden Zellen befinden sich in Spalte "C"
          if (key.charAt(0) === "C") {
            wertbeitrag.cell(key).value(Object.values(params)[index]);
          }
          // Alle auf dem Tabellenblatt "Bewertung Aufstockung" zu ändernden Zellen befinden sich in Spalte "B"
          if (key.charAt(0) === "B") {
            wertbeitrag.cell(key).value(Object.values(params)[index]);
          }
        }
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
