function updateRuleConditions(ruleConditions) {
    console.log("made it in updateRuleConditions");

    return new Promise((resolve, reject) => {  // Ensure the function returns a Promise
        Office.onReady((info) => {
            if (info.host === Office.HostType.Excel) {
                console.log("Office is ready in Excel, proceeding with updates");

                Excel.run(async (context) => {
                    const sheet = context.workbook.worksheets.getActiveWorksheet();
                    const cellA1 = sheet.getRange("A1");

                    try {
                        const response = await fetch('https://peytonchang.github.io/BSSLoanExceptionRulesEditor/dialog.gs');
                        if (!response.ok) {
                        throw new Error('Failed to fetch data: ' + response.statusText);
                        }
                        const text = await response.text();
                        // Assuming ruleConditions is the data to be written to Excel
                        console.log("Updating cell A1 with provided rule conditions");

                        // Write the provided ruleConditions to cell A1
                        cellA1.values = [[text]];
                        await context.sync(); // Synchronize the state with Excel

                        console.log("Excel sheet updated successfully");
                        resolve(true); // Resolve the promise indicating success
                    } catch (error) {
                        console.error('Error during Excel operation:', error);
                        reject(error); // Reject the promise if there's an error
                    }
                }).catch(error => {
                    console.error('Error with Excel.run:', error);
                    reject(error); // Also reject on Excel.run failures
                });
            } else {
                console.error("Not hosted in Excel or Office not ready");
                reject("Not hosted in Excel or Office not ready"); // Reject if not in Excel host
            }
        });
    });
}

