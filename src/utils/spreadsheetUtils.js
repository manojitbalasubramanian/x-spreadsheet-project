// https://github.com/SheetJS/sheetjs/blob/github/demos/xspreadsheet/xlsxspread.js
// clone from the above repo.
export function stox(wb) {
  const out = [];
  wb.SheetNames.forEach(name => {
    const o = { name, rows: {} };
    const ws = wb.Sheets[name];
    if (!ws || !ws["!ref"]) return;

    const range = window.XLSX.utils.decode_range(ws["!ref"]);
    range.s = { r: 0, c: 0 };

    const aoa = window.XLSX.utils.sheet_to_json(ws, {
      raw: false,
      header: 1,
      range,
    });

    aoa.forEach((r, i) => {
      const cells = {};
      r.forEach((c, j) => {
        cells[j] = { text: c };
        const cellRef = window.XLSX.utils.encode_cell({ r: i, c: j });
        if (ws[cellRef] && ws[cellRef].f) {
          cells[j].text = "=" + ws[cellRef].f;
        }
      });
      o.rows[i] = { cells };
    });

    o.merges = [];
    (ws["!merges"] || []).forEach((merge, i) => {
      if (!o.rows[merge.s.r]) o.rows[merge.s.r] = { cells: {} };
      if (!o.rows[merge.s.r].cells[merge.s.c]) o.rows[merge.s.r].cells[merge.s.c] = {};
      o.rows[merge.s.r].cells[merge.s.c].merge = [
        merge.e.r - merge.s.r,
        merge.e.c - merge.s.c,
      ];
      o.merges[i] = window.XLSX.utils.encode_range(merge);
    });

    out.push(o);
  });

  return out;
}

export function xtos(sdata) {
  const out = window.XLSX.utils.book_new();
  sdata.forEach(xws => {
    const ws = {};
    const rowobj = xws.rows;
    const minCoord = { r: 0, c: 0 };
    const maxCoord = { r: 0, c: 0 };

    for (let ri = 0; ri < rowobj.len; ++ri) {
      const row = rowobj[ri];
      if (!row) continue;

      Object.keys(row.cells).forEach(k => {
        const idx = +k;
        if (isNaN(idx)) return;

        const cellRef = window.XLSX.utils.encode_cell({ r: ri, c: idx });
        maxCoord.r = Math.max(maxCoord.r, ri);
        maxCoord.c = Math.max(maxCoord.c, idx);

        let cellText = row.cells[k].text;
        let type = "s";

        if (!cellText) {
          cellText = "";
          type = "z";
        } else if (!isNaN(Number(cellText))) {
          cellText = Number(cellText);
          type = "n";
        } else if (["true", "false"].includes(cellText.toLowerCase())) {
          cellText = Boolean(cellText);
          type = "b";
        }

        ws[cellRef] = { v: cellText, t: type };
        if (type === "s" && cellText[0] === "=") {
          ws[cellRef].f = cellText.slice(1);
        }

        if (row.cells[k].merge) {
          ws["!merges"] = ws["!merges"] || [];
          ws["!merges"].push({
            s: { r: ri, c: idx },
            e: {
              r: ri + row.cells[k].merge[0],
              c: idx + row.cells[k].merge[1],
            },
          });
        }
      });
    }

    ws["!ref"] = window.XLSX.utils.encode_range({ s: minCoord, e: maxCoord });
    window.XLSX.utils.book_append_sheet(out, ws, xws.name);
  });

  return out;
}
