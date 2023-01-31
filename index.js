const fs = require("fs");
const path = require("path");
const inputPath = "./users.xlsx";
const outputFolder = "./output";
const xlsParser = new (require("simple-excel-to-json").XlsParser)();
const createCsvWriter = require("csv-writer").createObjectCsvWriter;
const chunkSize = 5;
const csvHeader = [
  { id: "ContactId", title: "ContactId" },
  { id: "Email", title: "Email" },
  { id: "FirstName", title: "FirstName" },
  { id: "LastName", title: "LastName" },
  { id: "CountryCode", title: "CountryCode" },
  { id: "CountryName", title: "CountryName" },
  { id: "OrganizationName", title: "OrganizationName" },
  { id: "LicenseQuantity", title: "LicenseQuantity" },
  { id: "NextPaymentDate", title: "NextPaymentDate" },
];

async function writeToCSV(header = [], data = [], path) {
  const csvWriter = createCsvWriter({
    path,
    header,
  });

  await csvWriter.writeRecords(data);
}

function deleteAllFiles(directoryPath) {
  if (fs.existsSync(directoryPath)) {
    fs.readdirSync(directoryPath).forEach((file) => {
      const curPath = path.join(directoryPath, file);
      if (fs.lstatSync(curPath).isDirectory()) {
        // recurse
        deleteFolderRecursive(curPath);
      } else {
        // delete file
        fs.unlinkSync(curPath);
      }
    });
  }
}

function startProcess() {
  const sheets = xlsParser.parseXls2Json(inputPath);
  let count = 0;
  let chunks = [];
  let index = 0;
  for (let sheet of sheets) {
    if (sheet[0].ContactId !== undefined) {
      for (let row of sheet) {
        count++;
        chunks.push(row);

        if (count === chunkSize) {
          count = 0;
          index++;
          writeToCSV(
            csvHeader,
            chunks,
            `${outputFolder}/users-chunk-${index}.csv`
          );
          chunks = [];
        }
      }
    }
  }

  if (!!chunks.length) {
    count = 0;
    index++;
    writeToCSV(csvHeader, chunks, `${outputFolder}/users-chunk-${index}.csv`);
    chunks = [];
  }
}

function clearFiles() {
  if (fs.existsSync(outputFolder)) {
    deleteAllFiles(outputFolder);
  }
}

(async function () {
  console.log("******** EXCEL TO CSV JOB is started **************");
  clearFiles();
  startProcess();
  console.log("******** EXCEL TO CSV JOB is completed *************");
})();
