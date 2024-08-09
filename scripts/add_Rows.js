async function add_Rows() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    let table = sheet.tables.getItemAt(0);
    table.load("rows");

    await context.sync(); // Ensure rows are loaded before accessing them

    for (let i = 0; i < 5; i++) {
      // Add a new row
      table.rows.add();

      // Sync the context to ensure the new row is added
      await context.sync();

      // Reload the rows after adding a new row
      table.load("rows");
      await context.sync();

      // Get the index of the newly added row
      let rowCount = table.rows.items.length;
      let newRowRange = table.rows.getItemAt(rowCount - 1).getRange();

      // Clear the contents of the newly added row
      newRowRange.clear(Excel.ClearApplyTo.contents);

      // Sync the context to apply the changes
      await context.sync();
    }
    await context.sync();
  });
}

  /** Default helper for invoking an action and handling errors. */
async function tryCatch(callback) {
    try {
      await callback();
    } catch (error) {
      // Note: In a production add-in, you'd want to notify the user through your add-in's UI.
      console.error(error);
    }
  }