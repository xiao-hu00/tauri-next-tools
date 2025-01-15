"use client";
import React, { useState } from "react";
import * as XLSX from "xlsx";

const ExcelMerger = () => {
  const [files, setFiles] = useState<FileList | null>(null);

  interface FileChangeEvent extends React.ChangeEvent<HTMLInputElement> {
    target: HTMLInputElement & EventTarget;
  }

  const handleFileChange = (event: FileChangeEvent): void => {
    if (event.target.files) {
      setFiles(event.target.files);
    }
  };

  const handleMergeAndDownload = async () => {
    const allData: unknown[][] = [];

    // Function to read each file and return the data
    const readFile = (file: File, index: number) => {
      return new Promise((resolve) => {
        const reader = new FileReader();
        reader.onload = (event) => {
          if (!event.target || !event.target.result) {
            resolve(undefined);
            return;
          }
          const result = event.target.result;
          if (typeof result === "string") {
            resolve(undefined);
            return;
          }
          const data = new Uint8Array(result);
          const workbook = XLSX.read(data, { type: "array" });

          workbook.SheetNames.forEach((sheetName) => {
            const worksheet = workbook.Sheets[sheetName];
            const jsonData = XLSX.utils.sheet_to_json(worksheet, {
              header: 1,
              blankrows: false,
            });
            if (index !== 0) {
              const list = jsonData.slice(1);
              allData.concat(list);
            } else {
              allData.concat(jsonData);
            }
          });

          resolve(undefined);
        };
        reader.readAsArrayBuffer(file);
      });
    };

    // Read all files and wait until all are processed
    await Promise.all(
      files ? Array.from(files).map((file, index) => readFile(file, index)) : []
    );
    // console.log("allData", allData);
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
