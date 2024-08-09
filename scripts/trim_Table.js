async function trim_Table() {
  await Excel.run(async (context) => {
    const sheet = context.workbook.worksheets.getActiveWorksheet();

    let table = sheet.tables.getItemAt(0);

    table.rows.load("values");

    return context.sync().then(function() {
      // Iterate through the rows in reverse order
      for (var i = table.rows.items.length - 1; i >= 0; i--) {
        var row = table.rows.items[i];
        var isEmpty = row.values[0].every(function(cell) {
          return cell === null || cell === "";
        });

        // If the row is empty, remove it
        if (isEmpty) {
          table.rows.getItemAt(i).delete();
        }
      }
      // Sync the context to apply the changes
      return context.sync();
    });
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