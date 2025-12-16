/* global document, Office, Word, Excel, PowerPoint */

// Centralized Trademark List
const trademarkMap = {
  LTE: "LTE®",
  LTEM: "LTEM®",
  "Flash Cigar": "Flash Cigar®",
  Dilse: "Dilse™",
};

Office.onReady((info) => {
  if (info.host) {
    document.getElementById("sideload-msg").style.display = "none";
    document.getElementById("app-body").style.display = "flex";

    // Bind the "Run" button from your HTML to our branding function
    document.getElementById("run").onclick = runBrandingTool;
  }
});

export async function runBrandingTool() {
  const host = Office.context.host;

  if (host === Office.HostType.Word) {
    await applyWordBranding();
  } else if (host === Office.HostType.Excel) {
    await applyExcelBranding();
  } else if (host === Office.HostType.PowerPoint) {
    await applyPowerPointBranding();
  }
}

/** WORD: Scans text in the document body **/
async function applyWordBranding() {
  await Word.run(async (context) => {
    const body = context.document.body;
    for (const [key, val] of Object.entries(trademarkMap)) {
      const searchResults = body.search(key, { matchCase: true });
      searchResults.load("items");
      await context.sync();
      searchResults.items.forEach((item) => item.insertText(val, "Replace"));
    }
    await context.sync();
  });
}

/** EXCEL: Scans all used cells in the active sheet **/
async function applyExcelBranding() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const usedRange = sheet.getUsedRange();
    usedRange.load("values");
    await context.sync();

    let values = usedRange.values;
    for (let r = 0; r < values.length; r++) {
      for (let c = 0; c < values[r].length; c++) {
        if (typeof values[r][c] === "string") {
          Object.keys(trademarkMap).forEach((key) => {
            const regex = new RegExp(`\\b${key}\\b`, "g");
            values[r][c] = values[r][c].replace(regex, trademarkMap[key]);
          });
        }
      }
    }
    usedRange.values = values;
    await context.sync();
  });
}

/** POWERPOINT: Scans all shapes and text frames across all slides **/
async function applyPowerPointBranding() {
  await PowerPoint.run(async (context) => {
    const slides = context.presentation.slides;
    slides.load("items");
    await context.sync();

    for (let slide of slides.items) {
      const shapes = slide.shapes;
      shapes.load("items");
      await context.sync();
      for (let shape of shapes.items) {
        if (shape.hasTextFrame) {
          const textRange = shape.textFrame.textRange;
          textRange.load("text");
          await context.sync();
          let newText = textRange.text;
          Object.keys(trademarkMap).forEach((key) => {
            newText = newText.replace(new RegExp(`\\b${key}\\b`, "g"), trademarkMap[key]);
          });
          textRange.insertText(newText, "Replace");
        }
      }
    }
    await context.sync();
  });
}
