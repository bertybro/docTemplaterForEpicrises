import fs from "fs";
import xlsx from "xlsx";
import Docxtemplater from "docxtemplater";
import JSZip from "jszip";
import * as utils from "./utils.js";

const data = xlsx.readFile("data.xlsx"); // Change to the actual name fo your .xlsx file
const sheetName = data.SheetNames[0];
const worksheet = data.Sheets[sheetName];
const rows = xlsx.utils.sheet_to_json(worksheet);

const templateFile = fs.readFileSync("template.docx", "binary");

rows.forEach((row, index) => {
  const doc = new Docxtemplater();
  doc.loadZip(new JSZip(templateFile));

  // Mapping the placeholders in the template with data from the Excel file

  doc.setData({
    cell1: row["Жетон"],
    cell2: row["Позывной"],
    cell3: utils.ExcelDateToJSDate(row["Дата рождения"]),
    cell4: utils.ExcelDateToJSDate(row["Дата ранения"]),
    cell5: utils.ExcelDateToJSDate(row["Дата исхода"]),
    cell6: row["Характер ранения"],
    cell7: utils.ExcelDateToJSDateLong(row["Дата исхода"]),
  });

  try {
    // render the document (replace all occurences of {Cell1,2,3...}
    doc.render();
  } catch (error) {
    function replaceErrors(key, value) {
      if (value instanceof Error) {
        return Object.getOwnPropertyNames(value).reduce(function (error, key) {
          error[key] = value[key];
          return error;
        }, {});
      }
      return value;
    }
    console.log(JSON.stringify({ error: error }, replaceErrors));

    if (error.properties && error.properties.errors instanceof Array) {
      const errorMessages = error.properties.errors
        .map(function (error) {
          return error.properties.explanation;
        })
        .join("\n");
      console.log("errorMessages", errorMessages);
    }
    throw error;
  }

  const buffer = doc.getZip().generate({ type: "nodebuffer" });

  if (fs.existsSync(`output/${row["Жетон"]}, ${row["Позывной"]}.docx`) === true) {
    console.log(`${row["Жетон"]}.docx already exists!`);

    for (let i = 1; i < 30; i++) {
      if (fs.existsSync(`output/duplicates/${row["Жетон"]}, ${row["Позывной"]} - ${i}.docx`) === true) {
        console.log(`${row["Жетон"]} - ${i}.docx already exists!`);
      } else {
        fs.writeFileSync(
          `output/duplicates/${row["Жетон"]}, ${row["Позывной"]} - ${i}.docx`,
          buffer
        );
        console.log(`DONE - ${row["Жетон"]} - ${i}`);
        break;
      }
    }
  } else {
    fs.writeFileSync(`output/${row["Жетон"]}, ${row["Позывной"]}.docx`, buffer);
    console.log(`DONE - ${row["Жетон"]}`);
  }
});
