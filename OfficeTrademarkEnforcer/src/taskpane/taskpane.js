/* global Office, Word, Excel, PowerPoint */

const trademarkMap = { LTE: "LTE®", LTEM: "LTEM®", "Flash Cigar": "Flash Cigar®", Dilse: "Dilse™" };

Office.onReady((info) => {
  if (info.host) {
    document.getElementById("app-body").style.display = "flex";
    document.getElementById("run").onclick = runBranding;
  }
});

async function runBranding() {
  const host = Office.context.host;
  if (host === Office.HostType.Word) await applyWordBranding();
  else if (host === Office.HostType.Excel) await applyExcelBranding();
  else if (host === Office.HostType.PowerPoint) await applyPowerPointBranding();
}

async function applyWordBranding() {
  await Word.run(async (context) => {
    const body = context.document.body;
    for (const [key, val] of Object.entries(trademarkMap)) {
      const results = body.search(key, { matchCase: true });
      results.load("items");
      await context.sync();
      results.items.forEach((item) => item.insertText(val, "Replace"));
    }
  });
}

async function applyExcelBranding() {
  await Excel.run(async (context) => {
    const range = context.workbook.worksheets.getActiveWorksheet().getUsedRange();
    range.load("values");
    await context.sync();
    let values = range.values.map((row) =>
      row.map((cell) => {
        if (typeof cell === "string") {
          Object.keys(trademarkMap).forEach(
            (k) => (cell = cell.replace(new RegExp(`\\b${k}\\b`, "g"), trademarkMap[k]))
          );
        }
        return cell;
      })
    );
    range.values = values;
    await context.sync();
  });
}
