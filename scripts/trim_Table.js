
async function trim_Table() {
    await Excel.run(async (context) => {
      const sheet = context.workbook.worksheets.getActiveWorksheet();
  
      sheet.protection.unprotect();
  
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