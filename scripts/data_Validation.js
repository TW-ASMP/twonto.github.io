async function data_Validation() {
  await Excel.run(async (context) => {
    let sheet = context.workbook.worksheets.getItem("_Metadata_");
    let inputTable = sheet.tables.getItem("table_Colume_Property_Table").load("values");
    let assetTable = sheet.tables.getItem("asset_Table").load("values");

    let tableObjectName = inputTable.columns.getItem("Table Object Name").load("values");
    let tableColumns = inputTable.columns.getItem("Table Column Name").load("values");
    let validationType = inputTable.columns.getItem("Data Validation Type").load("values");
    let validationRange = inputTable.columns.getItem("Value List Range").load("values");

    await context.sync();

    //Variables to hold filtered values
    let affected = [];
    let validation = [];

    assetTable.load('values');
    await context.sync();

    //Filters for numbered validation on the asst table
    for (let i = 0; i < tableObjectName.values.length; i++) {
      if (tableObjectName.values[i][0] == "asset_Table" && validationType.values[i][0] == "number") {
        affected.push(tableColumns.values[i][0]);
        validation.push(validationRange.values[i][0]);
      }
    }

    for (let i = 0; i < affected.length; i++) {
      let affectedRange = assetTable.columns.getItem(affected[i]).getDataBodyRange().load();
      await context.sync();

      //Clearing and applying new rule
      affectedRange.dataValidation.clear();
      affectedRange.dataValidation.rule = {
        decimal: {
          operator: Excel.DataValidationOperator.between, // Choose the validation type, e.g., "between"
          formula1: 0, // Lower bound of the decimal range
          formula2: 10000.0 // Upper bound of the decimal range
        }
      };
    }

    //Variables to hold filtered values
    affected = [];
    validation = [];

    assetTable.load('values');
    await context.sync();

    //Filters for closed lists and on the asset table
    for (let i = 0; i < tableObjectName.values.length; i++) {
      if (tableObjectName.values[i][0] == "asset_Table" && validationType.values[i][0] == "closed list validation") {
        affected.push(tableColumns.values[i][0]);
        validation.push(validationRange.values[i][0]);
      }
    }

    //Applying validation for closed lists
    for(let i = 0; i< affected.length; i++){
      let affectedRange = assetTable.columns.getItem(affected[i]).getDataBodyRange().load();
      let validationRange = sheet.tables.getItem(validation[i]).load('name');
      await context.sync();

      //Clearing and applying new rule
      affectedRange.dataValidation.clear();
      affectedRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "=INDIRECT(\"" + validationRange.name + "\")"
        }
      };
      affectedRange.dataValidation.errorAlert = {
        message: "Only valididated values",
        showAlert: true, // The default is 'true'.
        style: Excel.DataValidationAlertStyle.stop,
        title: "Invalid data"
      };
    }

    //Variables to hold filtered values
    affected = [];
    validation = [];
    
    //Filters for open lists and on the asst table
    for (let i = 0; i < tableObjectName.values.length; i++) {
      if (tableObjectName.values[i][0] == "asset_Table" && validationType.values[i][0] == "open list validation") {
        affected.push(tableColumns.values[i][0]);
        validation.push(validationRange.values[i][0]);
        }
      }

    //Applying validation for open lists
    for (let i = 0; i < affected.length; i++) {
      let affectedRange = assetTable.columns.getItem(affected[i]).getDataBodyRange().load();
      let validationRange = sheet.tables.getItem(validation[i]).load('name');
      await context.sync();

      //Clearing and applying new rule
      affectedRange.dataValidation.clear();
      affectedRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "=INDIRECT(\"" + validationRange.name + "\")"
          }
          };
      affectedRange.dataValidation.errorAlert = {
        message: "Not valididated value",
        showAlert: true, // The default is 'true'.
        style: Excel.DataValidationAlertStyle.warning,
        title: "Unvalidated Value"
        };

      affectedRange.conditionalFormats.clearAll();
      let formatting = affectedRange.conditionalFormats.add(Excel.ConditionalFormatType.custom);
      let cellAddress = getColLetter(affectedRange.columnIndex + 1);

      formatting.custom.rule.formula = "=AND(ISERROR(MATCH(" + cellAddress + "5, INDIRECT(\"" + validationRange.name + "\"), 0)), NOT(ISBLANK(" + cellAddress + "5)))";
      formatting.custom.format.fill.color ="yellow";
    }
  await context.sync();
  });
}

// Function to get the column letter from the column number
function getColLetter(columnNumber: number): string {
  let temp: number = columnNumber;
  let letter: string = '';
  while (temp > 0) {
    let mod: number = (temp - 1) % 26;
    letter = String.fromCharCode(65 + mod) + letter;
    temp = Math.floor((temp - mod) / 26);
  }
  return letter;
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

