"use client";
import React, { useState } from "react";
import * as XLSX from "xlsx";

const ExcelMerger = () => {
  const [files, setFiles] = useState([]);

  const handleFileChange = (event) => {
    setFiles(event.target.files);
  };

  const handleMergeAndDownload = async () => {
    const allData = [];

    // Function to read each file and return the data
    const readFile = (file, index) => {
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (event) => {
          const data = new Uint8Array(event.target.result);
          const workbook = XLSX.read(data, { type: "array" });

          workbook.SheetNames.forEach((sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
              blankrows: false,
            });
            if (index !== 0) {
              const list = jsonData.slice(1);
              allData.push(...list);
            } else {
              allData.push(...jsonData);
            }
          });

          resolve();
        };
        reader.readAsArrayBuffer(file);
      });
    };

    // Read all files and wait until all are processed
    await Promise.all(
      Array.from(files).map((file, index) => readFile(file, index))
    );
    console.log("allData", allData);
    // Create a new workbook and add the combined data
    const combinedWorksheet = XLSX.utils.aoa_to_sheet(allData);
    const newWorkbook = XLSX.utils.book_new();
    XLSX.utils.book_append_sheet(newWorkbook, combinedWorksheet, "MergedSheet");

    // Write the workbook to a file and trigger download
    const wbout = XLSX.write(newWorkbook, { bookType: "xlsx", type: "array" });
    const blob = new Blob([wbout], { type: "application/octet-stream" });
    const url = URL.createObjectURL(blob);
    const a = document.createElement("a");
    a.href = url;
    a.download = "merged_output.xlsx";
    document.body.appendChild(a);
    a.click();
    document.body.removeChild(a);
  };

  return (
    <div>
      <h1>选择文件夹</h1>
      <input
        type="file"
        multiple
        onChange={handleFileChange}
        ref={(input) => {
          if (input) {
            input.setAttribute("webkitdirectory", "true");
            input.setAttribute("directory", "true");
          }
        }}
      />
      <button onClick={handleMergeAndDownload}>合并并下载Excel文件</button>
    </div>
  );
};

export default ExcelMerger;
