Office.onReady(info => {
    if (info.host === Office.HostType.Excel) {
        document.getElementById("bind-events").onclick = bindEvents;
    }
});

function bindEvents() {
    Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        sheet.onChanged.add(onWorksheetChange);
        sheet.onSelectionChanged.add(onSelectionChange);
        sheet.onActivated.add(onWorksheetActivate);
        await context.sync();
        console.log("Events bound successfully");
    }).catch(error => {
        console.error(error);
    });
}

async function onWorksheetChange(event) {
    await Excel.run(async (context) => {
        const sheet = context.workbook.worksheets.getActiveWorksheet();
        const range = sheet.getRange(event.address);
        range.load("values");
        await context.sync();
        
        const cellValue = range.values[0][0];
        console.log("Cell changed: ", event.address, cellValue);
        document.getElementById("output").innerText = `Cell ${event.address} changed to: ${cellValue}`;
    }).catch(error => {
        console.error(error);
    });
}

function onSelectionChange(event) {
    console.log("Selection changed: ", event);
    // Add your custom logic here
}

function onWorksheetActivate(event) {
    console.log("Worksheet activated: ", event);
    // Add your custom logic here
}
