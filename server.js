const express = require("express");
const bodyParser = require("body-parser");
const Docxtemplater = require("docxtemplater");
const PizZip = require("pizzip");
const fs = require("fs");
const path = require("path");
const cors = require("cors");

const app = express();
const port = 3000;

app.use(cors());
app.use(bodyParser.json());

app.use(express.static(path.join(__dirname, "build")));

function generateDocumentContentChecker(formData) {
  console.log("formData:", formData);
  let content = "{nombreSolicitante} y {day} y {hour}";

  console.log("Before replacement:", content);

  content = content.replace(
    new RegExp("{nombreSolicitante}", "g"),
    formData.nombreSolicitante
  );
  content = content.replace(new RegExp("{day}", "g"), formData.day);
  content = content.replace(new RegExp("{hour}", "g"), formData.hour);

  console.log("After replacement:", content);

  return content;
}

app.post("/generate-word", async (req, res) => {
  const formData = await req.body;

  const templatePath = path.join(__dirname, "/templates/word-template.docx");
  const content = fs.readFileSync(templatePath, "binary");

  const zip = new PizZip(content);

  const doc = new Docxtemplater(zip, {
    paragraphLoop: true,
    linebreaks: true,
  });

  doc.render(formData);

  const buf = doc.getZip().generate({
    type: "nodebuffer",
    compression: "DEFLATE",
  });

  const outputPath = path.resolve(
    __dirname,
    `Orden Nº${formData.codigoNumerico} de Permanencia.docx`
  );
  fs.writeFileSync(outputPath, buf);

  res.sendFile(outputPath);
});

app.listen(port, () => {
  console.log(`Server is running on port ${port}`);
});
