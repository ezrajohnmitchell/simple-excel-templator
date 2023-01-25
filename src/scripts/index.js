import "../styles/index.scss";
import { saveAs } from "file-saver";
import { csvParse } from "d3";
const ExcelJS = require("exceljs");

window.generate = generate;

async function generate() {
  //validate
  let templateFile = document.getElementById("template").files[0];
  let csvFile = document.getElementById("data").files[0];

  if (!templateFile) {
    setError("*Must provide template");
    return;
  }
  if (!csvFile) {
    setError("*Must provide csv");
    return;
  }

  // Read template file
  const workbook = new ExcelJS.Workbook();
  await workbook.xlsx.load(fileToBuffer(templateFile));

  const template = workbook.worksheets[0];
  if (!template) {
    setError("*Must have at least 1 worksheet");
  }
  //Read csv file
  const reader = new FileReader();

  reader.onload = (e) => {
    const text = e.target.result;
    const csv = csvParse(text);

    if (csv.length <= 0) {
      setError("*csv must have data");
      return;
    }

    const nameKeys = Object.keys(csv[0]);

    const keytoTemplatePos = new Map();

    template.eachRow((row, rowNum) => {
      row.eachCell((cell, cellNum) => {
        if (nameKeys.includes(cell.value)) {
          keytoTemplatePos.set(cell.value, { x: rowNum, y: cellNum });
        }
      });
    });

    const validKeys = [...keytoTemplatePos.keys()];

    //generate copies of template and fill with row data
    csv.forEach((value, index) => {
      const generatedSheet = workbook.addWorksheet();

      generatedSheet.model = Object.assign(template.model, {
        mergeCells: template.model.merges,
      });
      generatedSheet.name = "" + (index + 1);

      validKeys.forEach((key) => {
        let pos = keytoTemplatePos.get(key);
        generatedSheet.getCell(pos.x, pos.y).value = value[key];
      });
    });

    // download file
    workbook.xlsx.writeBuffer().then((buffer) => {
      const blob = new Blob([buffer], {
        type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
      });
      saveAs(blob, "output.xlsx");
    });
  };

  reader.readAsText(csvFile);
}

async function fileToBuffer(file) {
  const blob = new Blob([file], {
    type: "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet;charset=utf-8",
  });
  return await blob.arrayBuffer();
}

function setError(message) {
  let error = document.getElementById("error");
  error.style.display = "block";
  error.innerHTML = message;
}
