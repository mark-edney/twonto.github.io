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
      let affectedValidation = sheet.tables.getItem(validation[i]).load('name');
      await context.sync();

      //Clearing and applying new rule
      affectedRange.dataValidation.clear();
      affectedRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "=INDIRECT(\"" + affectedValidation.name + "\")"
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
    
    //Filters for open lists and on the asset table
    for (let i = 0; i < tableObjectName.values.length; i++) {
      if (tableObjectName.values[i][0] == "asset_Table" && validationType.values[i][0] == "open list validation") {
        affected.push(tableColumns.values[i][0]);
        validation.push(validationRange.values[i][0]);
        }
      }
    
    //Applying validation for open lists
    for (let i = 0; i < affected.length; i++) {
      let affectedRange = assetTable.columns.getItem(affected[i]).getDataBodyRange().load();
      let affectedValidation = sheet.tables.getItem(validation[i]).load('name');
      await context.sync();

      //Clearing and applying new rule
      affectedRange.dataValidation.clear();
      affectedRange.dataValidation.rule = {
        list: {
          inCellDropDown: true,
          source: "=INDIRECT(\"" + affectedValidation.name + "\")"
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

      formatting.custom.rule.formula = "=AND(ISERROR(MATCH(" + cellAddress + "5, INDIRECT(\"" + affectedValidation.name + "\"), 0)), NOT(ISBLANK(" + cellAddress + "5)))";
      formatting.custom.format.fill.color ="yellow";

      //Variables to hold filtered values
      affected = [];
      validation = [];

      let instrumentationTable = sheet.tables.getItem("instrument_Table").load("values");
      await context.sync();

      //Filters for open lists and on the instrumentaiton table
      for (let i = 0; i < tableObjectName.values.length; i++) {
        if (tableObjectName.values[i][0] == "instrument_Table" && validationType.values[i][0] == "open list validation") {
          affected.push(tableColumns.values[i][0]);
          validation.push(validationRange.values[i][0]);
        }
      }

      //Applying validation for open lists
      for (let i = 0; i < affected.length; i++) {
        let affectedRange = instrumentationTable.columns.getItem(affected[i]).getDataBodyRange().load();
        await context.sync();

        //Clearing and applying new rule
        affectedRange.dataValidation.clear();
        affectedRange.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "=INDIRECT(\"" + validation[i] + "\")"
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

        formatting.custom.rule.formula = "=AND(ISERROR(MATCH($C14, INDIRECT(\"" + validation[i] + "\"), 0)), NOT(ISBLANK($C14)))";
        formatting.custom.format.fill.color = "yellow";

        //Instrumentation Parameter UOM, seperate code
        let affectedRange = instrumentationTable.columns.getItem("Parameter UOM (open value list)").getDataBodyRange().load();
        await context.sync();

        //Clearing and applying new rule
        affectedRange.dataValidation.clear();
        affectedRange.dataValidation.rule = {
          list: {
            inCellDropDown: true,
            source: "=OFFSET(XLOOKUP($C14,INDIRECT(\"paramter_Type[[#All], [Parameter Type]]\"), INDIRECT(\"paramter_Type[[#All], [UOM1]]\")),0,0,1,COUNTA(XLOOKUP($C14, INDIRECT(\"paramter_Type[[#All], [Parameter Type]]\"),INDIRECT(\"paramter_Type[#All]\")))-1)"
          }
        };
        affectedRange.dataValidation.errorAlert = {
          message: "Not valididated value",
          showAlert: true, // The default is 'true'.
          style: Excel.DataValidationAlertStyle.warning,
          title: "Unvalidated Value"
        };

      }
    }
  await context.sync();
  });
}