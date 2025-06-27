document.getElementById('upload').addEventListener('change', function (e) {
  const file = e.target.files[0];
  const reader = new FileReader();

  reader.onload = function (e) {
    const data = new Uint8Array(e.target.result);
    const workbook = XLSX.read(data, { type: 'array' });

    const sheetName = workbook.SheetNames[0];
    const sheet = workbook.Sheets[sheetName];
    const json = XLSX.utils.sheet_to_json(sheet, { header: 1 });

    const lastRow = json[json.length - 1];
    const lastRowValue = lastRow ? lastRow.findLast(v => v !== 0 && v !== null && v !== "") : "없음";

    let lastColValue = "없음";
    for (let row = json.length - 1; row >= 0; row--) {
      if (json[row] && json[row].length > 0) {
        const lastNonEmpty = [...json[row]].reverse().find(v => v !== 0 && v !== null && v !== "");
        if (lastNonEmpty !== undefined) {
          lastColValue = lastNonEmpty;
          break;
        }
      }
    }

    document.getElementById("lastRow").textContent = lastRowValue ?? "없음";
    document.getElementById("lastCol").textContent = lastColValue ?? "없음";
  };

  reader.readAsArrayBuffer(file);
});
