const JSZip = require("jszip");
const fs = require("fs");
const path = require("path");

const zip = new JSZip();

// Add files and folders to the zip
zip.file("src/skeleton.xlsm/XMLsource/[Content_Types].xml", fs.readFileSync("src/skeleton.xlsm/XMLsource/[Content_Types].xml"));

const addFolderToZip = (folderPath, zipFolder) => {
  const files = fs.readdirSync(folderPath);
  files.forEach((file) => {
    const fullPath = path.join(folderPath, file);
    const stat = fs.statSync(fullPath);
    if (stat.isDirectory()) {
      const subFolder = zipFolder.folder(file);
      addFolderToZip(fullPath, subFolder);
    } else {
      zipFolder.file(file, fs.readFileSync(fullPath));
    }
  });
};

// Add folders to the zip
let xlFolder = "src/skeleton.xlsm/XMLsource/xl";
addFolderToZip(xlFolder, zip.folder("xl"));

xlFolder = "src/skeleton.xlsm/XMLsource/docProps";
addFolderToZip(xlFolder, zip.folder("docProps"));

xlFolder = "src/skeleton.xlsm/XMLsource/_rels";
addFolderToZip(xlFolder, zip.folder("_rels"));

// Generate the zip file
zip.generateAsync({ type: "nodebuffer" }).then((content) => {
  fs.writeFileSync("Excel_Skeleton.zip", content);
  console.log("Zip file created: Excel_Skeleton.zip");
});