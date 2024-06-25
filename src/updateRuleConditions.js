function loadOfficeJsAndInitialize() {
    return new Promise((resolve, reject) => {
        // Create a new script element
        const script = document.createElement('script');
        script.src = "https://appsforoffice.microsoft.com/lib/1/hosted/office.js";
        script.onload = () => {
            // Wait for Office to initialize
            Office.onReady((info) => {
                if (info.host) {
                    console.log('Office is now ready!');
                    resolve(info);
                } else {
                    reject('Office is not ready or not loaded correctly.');
                }
            });
        };
        script.onerror = () => {
            reject('Failed to load Office.js');
        };
        // Append the script to the document's head
        document.head.appendChild(script);
    });
}

function updateRuleConditions(ruleConditions) {
    console.log("Attempting to run updateRuleConditions");

    return new Promise((resolve, reject) => {
        loadOfficeJsAndInitialize().then(() => {
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
                            console.log("Updating cell A1 with fetched data");

                            cellA1.values = [[text]];
                            await context.sync();
                            console.log("Excel sheet updated successfully");
                            resolve(true);
                        } catch (error) {
                            console.error('Error during Excel operation:', error);
                            reject(error);
                        }
                    }).catch(error => {
                        console.error('Error with Excel.run:', error);
                        reject(error);
                    });
                } else {
                    console.error("Not hosted in Excel or Office not ready");
                    reject("Not hosted in Excel or Office not ready");
                }
            });
        }).catch(error => {
            console.error("Failed to load Office.js:", error);
            reject(error);
        });
    });
}