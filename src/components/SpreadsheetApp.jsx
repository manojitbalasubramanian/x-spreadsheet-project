import React, { useEffect, useRef, useState } from "react";
import { stox, xtos } from "../utils/spreadsheetUtils";

const SpreadsheetApp = () => {
  const spreadsheetRef = useRef(null);
  const [spreadsheet, setSpreadsheet] = useState(null);

  useEffect(() => {
    const xspreadsheet = window.x_spreadsheet || window.xspreadsheet;
    if (spreadsheetRef.current?.children.length > 0) {
    spreadsheetRef.current.innerHTML = "";
    }
    const s = xspreadsheet(spreadsheetRef.current, { showToolbar: true });
    setSpreadsheet(s);
  }, []);

  const handleFile = (e) => {  // helps to upload .xls or .xlsx file
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const wb = window.XLSX.read(data, { type: "array" });
      const jsonData = stox(wb);   // stox is imported from utils
      if (spreadsheet) {
        spreadsheet.loadData(jsonData);
        } else {
        console.warn("Spreadsheet not initialized yet.");
      }
    };

    reader.readAsArrayBuffer(file);
  };

  const handleExport = () => {  // export and save as .xlsx 
    const json = spreadsheet.getData();
    const wb = xtos(json);  // xtos is imported from utils
    window.XLSX.writeFile(wb, "exported.xlsx");
  };

  const process = () => {   // Assuming you want to copy the value from column C to column D
    if (!spreadsheet) return;

    const data = spreadsheet.getData();

    const updatedData = data.map(sheet => {
        const newRows = { ...sheet.rows };

        Object.keys(newRows).forEach(rowIndex => {
        const row = newRows[rowIndex];
        if (row && row.cells && row.cells[2]) {  // index 2 corresponds to column C
            const cellA = row.cells[2].text;
            row.cells[3] = { ...row.cells[3], text: cellA }; // index 3 corresponds to column D
        }
        });

        return { ...sheet, rows: newRows };
    });

    spreadsheet.loadData(updatedData);  // rendering the choosen file data
    };

  return (
    <div>
      <h1>Spreadsheet Project</h1>

      <input
        type="file"
        accept=".xlsx, .xls"
        onChange={handleFile}
      />
      <button
        onClick={handleExport}>
        Export
      </button>
      <button onClick={process} >process</button>

      <div
        ref={spreadsheetRef}
        style={{ height: "100px" }}
      />
    </div>
  );
};

export default SpreadsheetApp;
