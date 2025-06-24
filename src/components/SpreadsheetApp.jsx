import React, { useEffect, useRef, useState } from "react";
import { stox, xtos } from "../utils/spreadsheetUtils";

const SpreadsheetApp = () => {
  const spreadsheetRef = useRef(null);
  const [spreadsheet, setSpreadsheet] = useState(null);

  useEffect(() => {
    const xspreadsheet = window.x_spreadsheet || window.xspreadsheet;
    const s = xspreadsheet(spreadsheetRef.current, { showToolbar: true });
    setSpreadsheet(s);
  }, []);

  const handleFile = (e) => {
    const file = e.target.files[0];
    const reader = new FileReader();

    reader.onload = (evt) => {
      const data = new Uint8Array(evt.target.result);
      const wb = window.XLSX.read(data, { type: "array" });
      const jsonData = stox(wb);
      spreadsheet.loadData(jsonData);
    };

    reader.readAsArrayBuffer(file);
  };

  const handleExport = () => {
    const json = spreadsheet.getData();
    const wb = xtos(json);
    window.XLSX.writeFile(wb, "exported.xlsx");
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
      <button>process</button>

      <div
        ref={spreadsheetRef}
        style={{ height: "500px" }}
      />
    </div>
  );
};

export default SpreadsheetApp;
