Office.onReady(() => {
  if (Office.context) {
    insertText("Add-in loaded successfully");
  }
});

async function insertText(text) {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();
    const range = sheet.getRange("A1");

    range.values = [[text]];
    await context.sync();
  });
}