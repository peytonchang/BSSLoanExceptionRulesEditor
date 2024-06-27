(function() {
    Office.onReady(info => {
        if (info.host === Office.HostType.Excel) {
            // Setup event listeners or initializations here
            buttonClick();
        }
    });

    function buttonClick() {
        const getAndDisplay = document.getElementById('getAndDisplayLoanInputData');
        const execute = document.getElementById('executeRules');
        const viewInputData = document.getElementById('viewLoanInputData');
        const viewResultData = document.getElementById('viewResultsData');

        getAndDisplay.addEventListener('click', getAndDisplayLoanInputData);
        execute.addEventListener('click', executeRules);    
        viewInputData.addEventListener('click', viewLoanInputData); 
        viewResultData.addEventListener('click', viewResultsData);
    }

    async function getAndDisplayLoanInputData() {
        try {
            await Excel.run(async (context) => {
                // Get the active worksheet
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                sheet.load('name'); // Optional: Load the sheet name for debugging or logging
                console.log("made it here 1");

                // Get the last column in the used range and load its index
                const lastCol = sheet.getUsedRange().getLastColumn();
                lastCol.load('columnIndex');
                await context.sync();  // Ensure the last column index is loaded
                console.log("made it here 2");
    
                // Get the range for the first row to read the headers and load their values
                const headerRange = sheet.getRange("A1:J1");
                headerRange.load('values');
                await context.sync();  // Ensure the header values are loaded
                console.log("made it here 3");
    
                // Create a dictionary of column headers to their index
                const dicColumn = getColumnDictionary(headerRange.values[0]);
                console.log("made it here 4");
    
                // Get the currently selected range and load its row index
                const activeRange = context.workbook.getSelectedRange();
                activeRange.load('rowIndex');
                await context.sync();  // Ensure the row index of the selected range is loaded
                console.log("made it here 5");
    
                // Calculate the active row index
                const activeRow = activeRange.rowIndex;
                console.log("made it here 6");
    
                // Load necessary cells for "Loan Input Data" and "Lock" from the current row
                const loanInputDataCell = sheet.getCell(activeRow, dicColumn['Loan Input Data']);
                const lockCell = sheet.getCell(activeRow, dicColumn['Lock']);
                loanInputDataCell.load('values');
                lockCell.load('values');
                await context.sync();  // Ensure the cell values are loaded
                console.log("made it here 7");
    
                // Check if the loan data is locked
                if (loanInputDataCell.values[0][0] !== '' && lockCell.values[0][0] === true) {
                    console.log("made it here 8");
                    // Display a message if the data is locked
                    Office.context.ui.displayDialogAsync('The Loan Input Data is locked and cannot be retrieved.');
                } else {
                    console.log("made it here 9");
                    // Call other functions to handle unlocked data
                    await getLoanInputData();
                    await viewLoanInputData();
                }
            });
        } catch (error) {
            console.error("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }
    }

    async function executeRules() {
        try {
            await Excel.run(async (context) => {
                console.log("made it here (executeRules) 1");
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const lastCol = sheet.getUsedRange().getLastColumn();
                lastCol.load('columnIndex');
                await context.sync();
    
                console.log("made it here (executeRules) 2");
                const headerRange = sheet.getRange("A1:J1");
                headerRange.load('values');
                await context.sync();
    
                console.log("made it here (executeRules) 3");
                const dicColumn = getColumnDictionary(headerRange.values[0]);
                const activeRange = context.workbook.getSelectedRange();
                activeRange.load('rowIndex');
                const activeCell = context.workbook.getSelectedRange();
                activeCell.load("address");
                await context.sync();
    
                const activeRow = activeRange.rowIndex;
    
                console.log("made it here (executeRules) 4");
                const environmentCell = sheet.getCell(activeRow, dicColumn['Environment']);
                const ruleProjectCell = sheet.getCell(activeRow, dicColumn['Rule Project']);
                const loanNumberCell = sheet.getCell(activeRow, dicColumn['Loan #']);
                const effectiveDateCell = sheet.getCell(activeRow, dicColumn['Effective Date']);
                const inputDataJsonCell = sheet.getCell(activeRow, dicColumn['Loan Input Data']);
                console.log("inputDataJsonCell: " + inputDataJsonCell);
    
                console.log("made it here (executeRules) 5");
                environmentCell.load('values');
                ruleProjectCell.load('values');
                loanNumberCell.load('values');
                effectiveDateCell.load('values');
                inputDataJsonCell.load('values');
                await context.sync();
    
                console.log("made it here (executeRules) 6");
                let inputDataJson = inputDataJsonCell.values[0][0];
    
                console.log("made it here (executeRules) 7");
                if (!inputDataJson) {
                    console.log("made it here (executeRules) 8");
                    // Asynchronous function to populate data
                    await getLoanInputData();
                    await context.sync();
                    inputDataJsonCell.load('values');
                    await context.sync();
                    inputDataJson = inputDataJsonCell.values[0][0]; // Reload value after update
                }
                console.log("inputDataJson: " + inputDataJson);
    
                console.log("made it here (executeRules) 9");
                if (inputDataJson) {
                    console.log("made it here (executeRules) 10");
                    console.log("inputDataJson:", inputDataJson);
                    let tmpJson = JSON.parse(inputDataJson);
                    const effectiveDate = effectiveDateCell.values[0][0];
    
                    console.log("made it here (executeRules) 11");
                    tmpJson.pricingDate = effectiveDate;
                    tmpJson.loanEligibilityDate = effectiveDate;
                    inputDataJson = JSON.stringify(tmpJson);
                }
    
                console.log("made it here (executeRules) 12");
                console.log("getServicParams call parameter 1 value: " + environmentCell.values[0][0]);
                console.log("getServicParams call parameter 2 value: " + ruleProjectCell.values[0][0]);
                console.log("inputDataJson: " + inputDataJson);
                const serviceParams = getServiceParams(environmentCell.values[0][0], ruleProjectCell.values[0][0]);
                const resultsJSON = await fetchServiceResultsJSON(serviceParams, loanNumberCell.values[0][0], inputDataJson);
                console.log("resultsJSON:", JSON.stringify(resultsJSON, null, 2));
                let formattedResults = '';
                
                console.log("made it here (executeRules) 13");
                if (propertyExists(resultsJSON, 'result')) {
                    console.log("made it here (executeRules) 14");
                    formattedResults = formatServiceResults(resultsJSON);
                }
    
                console.log("made it here (executeRules) 15");
                const timestampCell = sheet.getCell(activeRow, dicColumn['Execute Rules Timestamp']);
                const resultsCell = sheet.getCell(activeRow, dicColumn['Results']);
    
                console.log("made it here (executeRules) 16");
                timestampCell.values = [[new Date().toLocaleString("en-US", { timeZone: "America/New_York" })]];
                resultsCell.values = [[formattedResults]];
    
                await context.sync();
    
                // Assuming viewResultsData is a function to display results
                await viewResultsData();
            });
        } catch (error) {
            console.error("Error in executeRules:", error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info:", JSON.stringify(error.debugInfo));
            }
        }
    }

    function test() {
        Excel.run(function (context) {
            // Get the active worksheet
            var sheet = context.workbook.worksheets.getActiveWorksheet();
        
            // Request to load the 'name' property of the worksheet
            sheet.load('name');
        
            // Sync the context to actually fetch the property
            return context.sync().then(function () {
                console.log("Active worksheet is: " + sheet.name);
                return sheet; // You can now use 'sheet' to interact with the active worksheet
            });
        }).catch(function (error) {
            console.error("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }

    async function viewLoanInputData() {
        console.log("made it here (viewLoanInputData) 2");
        try {
            await Excel.run(async (context) => {
                console.log("made it here (viewLoanInputData) 2");
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const lastCol = sheet.getUsedRange().getLastColumn();
                lastCol.load('columnIndex');
                await context.sync();
    
                console.log("made it here (viewLoanInputData) 3");
                const headerRange = sheet.getRange("A1:J1");
                headerRange.load('values');
                await context.sync();
    
                console.log("made it here (viewLoanInputData) 4");
                const dicColumn = getColumnDictionary(headerRange.values[0]);
                const activeRange = context.workbook.getSelectedRange();
                activeRange.load('rowIndex');
                await context.sync();
    
                console.log("made it here (viewLoanInputData) 5");
                const activeRow = activeRange.rowIndex;
    
                console.log("made it here (viewLoanInputData) 6");
                // Load necessary cells
                // const environmentCell = sheet.getCell(activeRow, dicColumn['Environment'] + 1);
                // const ruleProjectCell = sheet.getCell(activeRow, dicColumn['Rule Project'] + 1);
                // const loanNumberCell = sheet.getCell(activeRow, dicColumn['Loan #'] + 1);
                // const loanInputDataCell = sheet.getCell(activeRow, dicColumn['Loan Input Data'] + 1);
                const environmentCell = sheet.getCell(activeRow, dicColumn['Environment']);
                const ruleProjectCell = sheet.getCell(activeRow, dicColumn['Rule Project']);
                const loanNumberCell = sheet.getCell(activeRow, dicColumn['Loan #']);
                const loanInputDataCell = sheet.getCell(activeRow, dicColumn['Loan Input Data']);

                console.log("made it here (viewLoanInputData) 7");
                environmentCell.load('values');
                ruleProjectCell.load('values');
                loanNumberCell.load('values');
                loanInputDataCell.load('values');
                await context.sync();
    
                console.log("made it here (viewLoanInputData) 8");
                // Retrieve the values
                const environment = environmentCell.values[0][0];
                const ruleProject = ruleProjectCell.values[0][0];
                const loanNumber = loanNumberCell.values[0][0];
                const loanInputData = loanInputDataCell.values[0][0];
    
                console.log("made it here (viewLoanInputData) 9");
                // Display the JSON in a custom UI element or alert, adjust `showJSON` accordingly
                showJSON(loanInputData, `${ruleProject} Input Data Viewer - ${environment} #${loanNumber}`);
            });
        } catch (error) {
            console.error("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }
    }

    async function viewResultsData() {
        await Excel.run(async (context) => {
            console.log("made it here (viewResultsData) 1");
            const sheet = context.workbook.worksheets.getActiveWorksheet();
            const lastCol = sheet.getUsedRange().getLastColumn();
            lastCol.load('columnIndex');
            await context.sync();
    
            console.log("made it here (viewResultsData) 2");
            const headerRange = sheet.getRange("A1:J1");
            headerRange.load('values');
            await context.sync();
    
            console.log("made it here (viewResultsData) 3");
            const dicColumn = getColumnDictionary(headerRange.values[0]);
            const activeCell = context.workbook.getSelectedRange();
            activeCell.load("rowIndex");
            await context.sync();
    
            // Assuming activeCell is a single cell and not a range
            const activeRow = activeCell.rowIndex;
    
            console.log("made it here (viewResultsData) 4");
            // Load necessary cells
            const environmentCell = sheet.getCell(activeRow, dicColumn['Environment']);
            const ruleProjectCell = sheet.getCell(activeRow, dicColumn['Rule Project']);
            const loanNumberCell = sheet.getCell(activeRow, dicColumn['Loan #']);
            const resultsCell = sheet.getCell(activeRow, dicColumn['Results']);
    
            console.log("made it here (viewResultsData) 5");
            environmentCell.load('values');
            ruleProjectCell.load('values');
            loanNumberCell.load('values');
            resultsCell.load('values');
            await context.sync();
    
            console.log("made it here (viewResultsData) 6");
            // Retrieve the values
            const environment = environmentCell.values[0][0];
            const ruleProject = ruleProjectCell.values[0][0];
            const loanNumber = loanNumberCell.values[0][0];
            const resultsData = resultsCell.values[0][0];
    
            console.log("made it here (viewResultsData) 7");
            // Show the JSON results in a custom viewer
            showJSON(resultsData, `${ruleProject} Results Viewer - ${environment} #${loanNumber}`);
        }).catch(error => {
            console.error("Error in viewResultsData:", error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info:", JSON.stringify(error.debugInfo));
            }
        });
    }

    async function getLoanInputData() {
        console.log("made it here (getLoanInputData) 1");
        try {
            await Excel.run(async (context) => {
                console.log("made it here (getLoanInputData) 2");
                const sheet = context.workbook.worksheets.getActiveWorksheet();
                const lastCol = sheet.getUsedRange().getLastColumn();
                lastCol.load('columnIndex');
                await context.sync();
                
                console.log("made it here (getLoanInputData) 3");
                const headerRange = sheet.getRange("A1:J1");
                headerRange.load('values');
                await context.sync();
    
                console.log("made it here (getLoanInputData) 4");
                const dicColumn = getColumnDictionary(headerRange.values[0]);
                console.log("Keys of dicColumn:", Object.keys(dicColumn));
                console.log("Entries of dicColumn:", Object.entries(dicColumn));
                console.log("made it here (getLoanInputData) 4.1");
                console.log("headerRange.values[0]: " + headerRange.values[0]);
                const activeRange = context.workbook.getSelectedRange();
                console.log("made it here (getLoanInputData) 4.2");
                activeRange.load('rowIndex');
                console.log("made it here (getLoanInputData) 4.3");
                await context.sync();
    
                console.log("made it here (getLoanInputData) 5");
                const activeRow = activeRange.rowIndex;
    
                console.log("made it here (getLoanInputData) 6");
                // Load necessary cells
                // const environmentCell = sheet.getCell(activeRow, dicColumn['Environment'] + 1);
                // const ruleProjectCell = sheet.getCell(activeRow, dicColumn['Rule Project'] + 1);
                // const loanNumberCell = sheet.getCell(activeRow, dicColumn['Loan #'] + 1);
                const environmentCell = sheet.getCell(activeRow, dicColumn['Environment']);
                const ruleProjectCell = sheet.getCell(activeRow, dicColumn['Rule Project']);
                const loanNumberCell = sheet.getCell(activeRow, dicColumn['Loan #'])

                console.log("made it here (getLoanInputData) 7");
                environmentCell.load('values');
                ruleProjectCell.load('values');
                loanNumberCell.load('values');
                await context.sync();


                console.log("made it here (getLoanInputData) 8");
                // Retrieve the values
                const environment = environmentCell.values[0][0];
                const ruleProject = ruleProjectCell.values[0][0];
                const loanNumber = loanNumberCell.values[0][0];
                console.log("environment: " + environment);
                console.log("ruleProject: " + ruleProject);
                console.log("loanNumber: " + loanNumber);
                
                console.log("made it here (getLoanInputData) 9");
                const serviceParams = getServiceParams(environment, ruleProject);
                console.log("serviceParams:" + serviceParams);
                if (serviceParams.length > 0) {
                    console.log("made it here (getLoanInputData) 10");
                    const resultsJSON = await fetchServiceInputDataJSON(serviceParams, loanNumber);
                    let inputDataJSON = null;
                    
                    console.log("made it here (getLoanInputData) 11");
                    if (propertyExists(resultsJSON, 'parameters.loanData')) {
                        inputDataJSON = resultsJSON.parameters.loanData.value;
                        
                        console.log("made it here (getLoanInputData) 12");
                        if (ruleProject === 'getLoanExceptions') {
                            console.log("made it here (getLoanInputData) 13");
                            inputDataJSON.fullEligibility = true;
                        }
                    } else if (ruleProject === 'Pricing' && propertyExists(resultsJSON, 'result.inputData')) {
                        console.log("made it here (getLoanInputData) 14");
                        inputDataJSON = resultsJSON.result.inputData;
                    }
    
                    if (inputDataJSON != null) {
                        console.log("made it here (getLoanInputData) 15");
                        condenseJSON(inputDataJSON);
    
                        if (propertyExists(inputDataJSON, 'rateLockDate')) {
                            console.log("made it here (getLoanInputData) 16");
                            const effectiveDate = formatDate(new Date(inputDataJSON.rateLockDate), 'yyyy-MM-dd');
                            // await sheet.getCell(activeRow, dicColumn['Effective Date'] + 1).setValues([[effectiveDate]]);
                            // await sheet.getCell(activeRow, dicColumn['Effective Date']).setValues([[effectiveDate]]);
                            sheet.getCell(activeRow, dicColumn['Effective Date']).values = [[effectiveDate]];
                        }
                        
                        console.log("made it here (getLoanInputData) 17");
                        const currentTimeStamp = formatDate(new Date(), 'MM-dd-yyyy hh:mm:ss a');
                        // await sheet.getCell(activeRow, dicColumn['Input Data Timestamp'] + 1).setValues([[currentTimeStamp]]);
                        // await sheet.getCell(activeRow, dicColumn['Loan Input Data'] + 1).setValues([[JSON.stringify(inputDataJSON)]]);
                        // await sheet.getCell(activeRow, dicColumn['Input Data Timestamp']).setValues([[currentTimeStamp]]);
                        // await sheet.getCell(activeRow, dicColumn['Loan Input Data']).setValues([[JSON.stringify(inputDataJSON)]]);
                        sheet.getCell(activeRow, dicColumn['Input Data Timestamp']).values = [[currentTimeStamp]];
                        sheet.getCell(activeRow, dicColumn['Loan Input Data']).values = [[JSON.stringify(inputDataJSON)]];
                        await context.sync();
                    }
                }
                console.log("made it here (getLoanInputData) 18");
            });
        } catch (error) {
            console.log("made it here (getLoanInputData) 19");
            console.error("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }
    }

    function formatDate(date, format) {
        // Simple date formatting, adjust as needed or use a library like date-fns or moment.js
        let dd = String(date.getDate()).padStart(2, '0');
        let mm = String(date.getMonth() + 1).padStart(2, '0'); //January is 0!
        let yyyy = date.getFullYear();
    
        if (format === 'yyyy-MM-dd') {
            return `${yyyy}-${mm}-${dd}`;
        } else if (format === 'MM-dd-yyyy hh:mm:ss a') {
            let hh = date.getHours();
            let min = String(date.getMinutes()).padStart(2, '0');
            let ss = String(date.getSeconds()).padStart(2, '0');
            let ampm = hh >= 12 ? 'pm' : 'am';
            hh = hh % 12;
            hh = hh ? hh : 12; // the hour '0' should be '12'
            return `${mm}-${dd}-${yyyy} ${hh}:${min}:${ss} ${ampm}`;
        }
        return date.toISOString().slice(0, 10);
    }

    async function fetchServiceInputDataJSON(serviceParams, loanNumber) {
        console.log("made it here (fetchServiceInputDataJSON) 1");
        let inputData = '';
        const accessToken = await getAccessToken(serviceParams);  // Assuming getAccessToken is an async function
        const headers = new Headers({
            "Content-Type": "application/json",
            "Authorization": `Bearer ${accessToken}`
        });
    
        console.log("made it here (fetchServiceInputDataJSON) 2");
        let url = `https://${serviceParams[0]}/lendingservices/externalServices/`;
        let options = {
            method: "GET",
            headers: headers,
            muteHttpExceptions: true  // This option does not exist in the standard fetch API
        };
    
        console.log("made it here (fetchServiceInputDataJSON) 3");
        if (serviceParams[3] === 'getPricing') {
            url += 'getPricingInputData';
            options.body = JSON.stringify({ "bypassLogging": true, "inputData": { "loanNumber": loanNumber } });
            options.method = "POST"; // Changed to POST since we are sending a payload
        } else {
            url += `getRuleEngineParameters?loanNumber=${loanNumber}&beanName=${serviceParams[4]}`;
        }
    
        console.log("made it here (fetchServiceInputDataJSON) 4");
        try {
            console.log("made it here (fetchServiceInputDataJSON) 5");
            console.log("url: " + url);
            const response = await fetch(url, options);
            if (response.status !== 200) {
                console.log("made it here (fetchServiceInputDataJSON) 6");
                console.log(`Validate request returned HTTP status code: ${response.status}`);
            } else {
                console.log("made it here (fetchServiceInputDataJSON) 7");
                const data = await response.json();
                inputData = data;
            }
        } catch (error) {
            console.error('Fetch error:', error);
        }
    
        console.log("made it here (fetchServiceInputDataJSON) 6");
        return inputData;
    }

    function showJSON(json, title) {
        console.log("made it here (showJSON) 1");
        // Escape single quotes in the JSON string
        const safeJson = json.replace(/'/g, "\\'");
    
        console.log("made it here (showJSON) 2");
        // Create a full HTML page content
        const htmlContent = `
            <html>
            <head>
                <base target="_top">
                <link rel="stylesheet" type="text/css" href="https://peytonchang.github.io/BSSLoanExceptionRulesEditor/src/taskpane/RemoteRules/jsonTree.css">
                <script src="https://peytonchang.github.io/BSSLoanExceptionRulesEditor/src/taskpane/RemoteRules/jsonTree.js"></script>
                <script language="javascript">
                function loadJson() {
                    const jsonData = JSON.parse('${safeJson}');
                    const wrapper = document.getElementById("wrapper");
                    const tree = jsonTree.create(jsonData, wrapper);
                    tree.expand(true);
                    const nodesToCollapse = ['applicantData', 'employmentData', 'loanOriginators', 'rateLockData', 'specialFeatures', 'blockTotals', 'feeServiceProviders'];
                    tree.rootNode.childNodes.forEach(node => {
                        if (nodesToCollapse.includes(node.label)) {
                            node.collapse(false);
                        }
                    });
                }
                </script>
            </head>
            <body onload="loadJson();">
                <table border="0" cellpadding="0" cellspacing="0" height="100%" width="100%" align="center">
                    <tr height="100%"><td id="wrapper"></td></tr>
                    <tr><td>&nbsp;</td></tr>
                    <tr><td align="center"><button onclick="closeModal();">Done</button></td></tr>
                </table>
            </body>
            </html>`;
    
        console.log("made it here (showJSON) 3");
        // Display the HTML in a new window or a modal
        const newWindow = window.open("", title, "width=1100,height=700");
        console.log("made it here (showJSON) 4");
        newWindow.document.write(htmlContent);
        console.log("made it here (showJSON) 5");
        newWindow.document.close(); // Close the document to parse the HTML
    }
    
    function closeModal() {
        window.close(); // This will close the modal window
    }
    
    function getServiceParams(environment, ruleProject) {
        console.log("made it here (getServiceParams) 1");
        let serviceParams = [];

        if (environment === 'BlueDLP DEV') {
            serviceParams.push('bluedlp-dev.bluesagedlp.com', 'admin', 'boston');
          } else if (environment === 'BlueDLP QA') {
            serviceParams.push('bluedlp-qa.bluesagedlp.com', 'admin', 'boston');
          } else if (environment === 'BlueDLP UAT') {
            serviceParams.push('bluedlp-uat.bluesagedlp.com', 'admin', 'boston');
          } else if (environment === 'HomeBridge DEV') {
            serviceParams.push('homebridge-dev.bluesageusa.com', '<system>', 'V>5.Ssfq*dDX7s-)');
          } else if (environment === 'HomeBridge QA') {
            serviceParams.push('homebridge-qa.bluesageusa.com', '<system>', 'V>5.Ssfq*dDX7s-)');
          } else if (environment === 'HomeBridge UAT') {
            serviceParams.push('homebridge-uat.bluesageusa.com', '<system>', 'V>5.Ssfq*dDX7s-)');
          } else if (environment === 'Planet DEV') {
            serviceParams.push('planet-dev.bluesageusa.com', '<system>', '@WaN-[Zb3<Nt~JrF');
          } else if (environment === 'Planet QA') {
            serviceParams.push('planet-qa.bluesageusa.com', '<system>', '@WaN-[Zb3<Nt~JrF');
          } else if (environment === 'Planet UAT') {
            serviceParams.push('uatcop.phlcorrespondent.com', '<system>', '@WaN-[Zb3<Nt~JrF');
          } else if (environment === 'Prime DEV') {
            serviceParams.push('prime-dev.bluesageusa.com', '<system>', 'NFvND9RXkL22y_ux');
          } else if (environment === 'Prime QA') {
            serviceParams.push('prime-qa.bluesageusa.com', '<system>', 'NFvND9RXkL22y_ux');
          } else if (environment === 'Prime UAT') {
            serviceParams.push('prime-uat.bluesageusa.com', '<system>', 'NFvND9RXkL22y_ux');
          } else if (environment === 'SpringEQ DEV') {
            serviceParams.push('springeq-dev.bluesageusa.com', '<system>', 'ZHVfEuF{?(3)7,NP');
          } else if (environment === 'SpringEQ QA') {
            serviceParams.push('springeq-qa.bluesageusa.com', '<system>', 'ZHVfEuF{?(3)7,NP');
          } else if (environment === 'SpringEQ UAT') {
            serviceParams.push('springeq-uat.bluesageusa.com', '<system>', 'ZHVfEuF{?(3)7,NP');
          } else if (environment === 'SpringEQ PROD') {
            serviceParams.push('springeq-services.bluesageusa.com', '<system>', 'x(pG6Z%~5cw*jb^2');
          } else if (environment === 'MFM DEV') {
            serviceParams.push('memberfirst-dev.bluesageusa.com', '<system>', '$C>LhSM:)Js?(3ak');
          } else if (environment === 'MFM QA') {
            serviceParams.push('memberfirst-qa.bluesageusa.com', '<system>', '$C>LhSM:)Js?(3ak');
          } else if (environment === 'MFM UAT') {
            serviceParams.push('memberfirst-uat.bluesageusa.com', '<system>', '$C>LhSM:)Js?(3ak');
          } else if (environment === 'Kind DEV') {
            serviceParams.push('kind-dev.bluesageusa.com', '<system>', 'V>5.Ssfq*dDX7s-)');
          } else if (environment === 'Kind QA') {
            serviceParams.push('kind-qa.bluesageusa.com', '<system>', 'V>5.Ssfq*dDX7s-)');
          } else if (environment === 'Kind UAT') {
            serviceParams.push('kind-uat.bluesageusa.com', '<system>', 'V>5.Ssfq*dDX7s-)');
          } else if (environment === 'Kind Retail DEV') {
            serviceParams.push('kindretail-dev.bluesageusa.com', '<system>', 'V>5.Ssfq*dDX7s-)');
          } else if (environment === 'Kind Retail QA') {
            serviceParams.push('kindretail-qa.bluesageusa.com', '<system>', 'V>5.Ssfq*dDX7s-)');
          } else if (environment === 'Kind Retail UAT') {
            serviceParams.push('kindretail-uat.bluesageusa.com', '<system>', 'V>5.Ssfq*dDX7s-)');
          } else if (environment === 'Lendage DEV') {
            serviceParams.push('lendage-dev.bluesageusa.com', '<system>', 'ZHVfEuF{?(3)7,NP');
          } else if (environment === 'Lendage QA') {
            serviceParams.push('lendage-qa.bluesageusa.com', '<system>', 'ZHVfEuF{?(3)7,NP');
          } else if (environment === 'Lendage UAT') {
            serviceParams.push('lendage-uat.bluesageusa.com', '<system>', 'ZHVfEuF{?(3)7,NP');
          } else if (environment === 'Westerra DEV') {
            serviceParams.push('westerra-dev.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'NASB DEV') {
            serviceParams.push('nasb-dev.bluesageusa.com', '<system>', 'M6udjY-3EPRzFMu3');
          } else if (environment === 'NASB QA') {
            serviceParams.push('nasb-qa.bluesageusa.com', '<system>', 'M6udjY-3EPRzFMu3');
          } else if (environment === 'NASB UAT') {
            serviceParams.push('nasb-uat.bluesageusa.com', '<system>', 'M6udjY-3EPRzFMu3');
          } else if (environment === 'MIG DEV') {
            serviceParams.push('mig-dev.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'MIG DEV') {
            serviceParams.push('mig-dev.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'MIG QA') {
            serviceParams.push('mig-qa.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'MIG UAT') {
            serviceParams.push('mig-uat.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'Westerra DEV') {
            serviceParams.push('Westerra-dev.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'Westerra QA') {
            serviceParams.push('Westerra-qa.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'Westerra UAT') {
            serviceParams.push('Westerra-uat.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'DCU DEV') {
            serviceParams.push('dcu-dev.bluesageusa.com', 'admin', 'pw123');
          } else if (environment === 'DCU QA') {
            serviceParams.push('dcu-qa.bluesageusa.com', 'admin', 'pw123');
          } else if (environment === 'DCU UAT') {
            serviceParams.push('dcu-uat.bluesageusa.com', 'admin', 'pw123');
          } else if (environment === 'SCU DEV') {
            serviceParams.push('scu-dev.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'SCU QA') {
            serviceParams.push('scu-qa.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'SCU UAT') {
            serviceParams.push('scu-uat.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'BlueSage DEV') {
            serviceParams.push('bluesage-dev.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'BlueSage QA') {
            serviceParams.push('bluesage-qa.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'BlueSage UAT') {
            serviceParams.push('bluesage-uat.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'FCM DEV') {
            serviceParams.push('fcm-dev.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'FCM QA') {
            serviceParams.push('fcm-qa.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'FCM UAT') {
            serviceParams.push('fcm-uat.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'Chevron DEV') {
            serviceParams.push('cfcu-dev.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'Chevron QA') {
            serviceParams.push('cfcu-qa.bluesageusa.com', 'admin', 'boston');
          } else if (environment === 'Chevron UAT') {
            serviceParams.push('cfcu-uat.bluesageusa.com', 'admin', 'boston');
          }
          console.log("made it here (getServiceParams) 2");
          console.log("serviceParams: " + serviceParams);
          //SELECT url, user_name, password FROM sys_global_external_system WHERE system_name = 'LENDING_SERVICES_REST_API';
          
          if (ruleProject === 'Loan Exceptions') {
            serviceParams.push('getLoanExceptions', 'loanExceptionServiceInputLoanDataDAO', 'loanExceptionRuleEngineService');
          } else if (ruleProject === 'Approval Conditions') {
            serviceParams.push('getLoanApprovalConditions', 'loanApprovalConditionsServiceInputDataDAO', 'loanApprovalConditionsRuleEngineService');
          } else if (ruleProject === 'Document Packages') {
            serviceParams.push('getDocumentPackages', 'documentRulesServiceInputDataDAO', 'documentsRuleEngineService');
          } else if (ruleProject === 'Fees') {
            serviceParams.push('getFees', 'feesServiceInputLoanDataDAO', 'feeRuleEngineService');
          } else if (ruleProject === 'Document Validations (OCR)') {
            serviceParams.push('getDocumentValidations', 'documentValidationServiceInputLoanDataDAO', 'documentValidationRuleEngineService');
          } else if (ruleProject === 'Data Mapping') {
            serviceParams.push('getDataMappings', 'dataMappingServiceInputDataDAO', 'dataMappingRuleEngineService');
          } else if (ruleProject === 'Date Expiration') {
            serviceParams.push('getDateCalculations', 'dateCalculationsServiceInputLoanDataDAO', 'dateCalculationsRuleEngineService');
          } else if (ruleProject === 'Decision Authorization') {
            serviceParams.push('getDecisionAuthorizations', 'decisionAuthorizationServiceInputDataDAO', 'decisionAuthorizationRuleEngineService');
          } else if (ruleProject === 'Pricing') {
            serviceParams.push('getPricing', 'getPricingInputData', 'getPricing');
          }
          console.log("made it here (getServiceParams) 3");
          console.log("serviceParams: " + serviceParams);
          return serviceParams;
    }
    

    function propertyExists(obj, key) {
        // Split the key string into an array of properties
        const properties = key.split(".");
    
        // Iterate through properties, returning false if object is null or property doesn't exist
        let currentObject = obj;
        for (const property of properties) {
            if (!currentObject || !Object.prototype.hasOwnProperty.call(currentObject, property)) {
                return false;
            }
            currentObject = currentObject[property];
        }
    
        return true; // Return true if all properties exist and are not undefined
    }

    function getColumnDictionary(aData) {
        console.log("made it here (getColumnnDictionary) 1");
        const columnDictionary = {};
        // Check if the input is an array and its first element is also an array
        if (!Array.isArray(aData) || !Array.isArray(aData[0])) {
            console.log("made it here (getColumnnDictionary) 2");
            console.error("Invalid header data format. Expected an array of arrays.");
            console.log("aData: " + aData);
            // Check if it's an array of strings (flat array of headers)
            if (Array.isArray(aData) && typeof aData[0] === 'string') {
                console.log("made it here (getColumnnDictionary) 3");
                aData.forEach((header, index) => {
                    const columnName = header.trim();
                    if (columnName !== '') {
                        console.log("made it here (getColumnnDictionary) 4");
                        columnDictionary[columnName] = index;
                    }
                });
            } else {
                console.log("made it here (getColumnnDictionary) 5");
                console.error("Unexpected data format:", aData);
            }
            console.log("made it here (getColumnnDictionary) 6");
            return columnDictionary;
        }
    
        // Process normally if it's an array of arrays
        aData[0].forEach((column, index) => {
            const columnName = column.trim();
            if (columnName !== '') {
                columnDictionary[columnName] = index;
            }
        });
        return columnDictionary;
    }

    function condenseJSON(json) {
        Object.keys(json).forEach(key => {
            if (json[key] !== null && typeof json[key] === "object") {
                condenseJSON(json[key]); // Recursively condense JSON
            } else {
                const type = dataType(json[key]);
    
                // Conditions for removing elements
                if (type === 'Null' ||
                    (type === 'Number' && json[key] === 0) ||
                    (type === 'Boolean' && !json[key]) ||
                    (type === 'String' && key !== 'value' && (json[key] === '' || json[key] === 'NaN'))) {
                    delete json[key];
                }
            }
        });
    }

    function dataType(value) {
        if (value === null) return 'Null';
        if (typeof value === 'number') return isNaN(value) ? 'NaN' : 'Number';
        if (typeof value === 'boolean') return 'Boolean';
        if (typeof value === 'string') return 'String';
        return 'Object';
    }

    async function getAccessToken(serviceParams) {
        console.log("made it here (getAccessToken) 1");
        const url = `https://${serviceParams[0]}/lendingservices/api/login`;
        const options = {
            method: 'POST',
            headers: {
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ "username": serviceParams[1], "password": serviceParams[2] })
        };
    
        try {
            const response = await fetch(url, options);
            console.log("made it here (getAccessToken) 2");
            if (!response.ok) {
                console.log("made it here (getAccessToken) 3");
                console.error('Login request returned HTTP status code:', response.status);
                return '';
            }
            console.log("made it here (getAccessToken) 4");
            const responseJSON = await response.json();
            return responseJSON.access_token;
        } catch (error) {
            console.error('Error fetching access token:', error);
            return '';
        }
    }

    function formatServiceResults(resultsJSON) {
        console.log("made it here (formatServiceResults) 1");
        let retVal = '';
      
        if (propertyExists(resultsJSON, 'result.serviceResults.businessRules')) {
            console.log("made it here (formatServiceResults) 2");
            resultsJSON.result.serviceResults.businessRules.forEach(exceptionRule => {
            if (!exceptionRule.valid) {
                console.log("made it here (formatServiceResults) 3");
                retVal += `${retVal ? ',' : ''}"${exceptionRule.id} - ${exceptionRule.message.replace(/"/g, '\'')}":`;
                if (propertyExists(exceptionRule, 'facts')) {
                    console.log("made it here (formatServiceResults) 4");
                    retVal += '{';
                    exceptionRule.facts.forEach((fact, index) => {
                    retVal += `${index ? ',' : ''}"${fact.name}":"${fact.value}"`;
                    });
      
                    if (exceptionRule.validatorClassName != null) {
                        console.log("made it here (formatServiceResults) 5");
                    retVal += `,"validatorClassName":"${exceptionRule.validatorClassName}"`;
                    }
                    retVal += '}';
              } else {
                    retVal += '""';
              }
            }
        });
        } else if (propertyExists(resultsJSON, 'result.serviceResults.approvalConditionItems')) {
            console.log("made it here (formatServiceResults) 6");
            resultsJSON.result.serviceResults.approvalConditionItems.forEach(conditionItem => {
                if (conditionItem.selected) {
                    console.log("made it here (formatServiceResults) 7");
                    retVal += `${retVal ? ',' : ''}"${conditionItem.id}${conditionItem.borrowerId ? ' [' + conditionItem.borrowerId + ']' : ''} - ${conditionItem.description.replace(/"/g, '\'')}":{`;
                    conditionItem.mergedText = conditionItem.mergedText.replace(/"/g, '\'').replace(/\n/g,'<br>');
                    retVal += `"mergedText":"${conditionItem.mergedText}", "responsibleParty":"${conditionItem.responsibleParty}", "required":"${conditionItem.required}"`;
      
                    if (conditionItem.preSelectionClassName != null) {
                        console.log("made it here (formatServiceResults) 8");
                    retVal += `,"preSelectionClassName":"${conditionItem.preSelectionClassName}"`;
                }
                retVal += '}';
            }
          });
        } else if (propertyExists(resultsJSON, 'result.serviceResults.documentServiceResults')) {
            console.log("made it here (formatServiceResults) 9");
            resultsJSON.result.serviceResults.documentServiceResults.forEach(docResult => {
            const documentPackage = docResult.documentPackage;
            retVal += `${retVal ? ',' : ''}"${documentPackage.id} - ${documentPackage.description}${documentPackage.applicable ? '' : ' (Not Available)'}":{`;
            retVal += `"applicable":"${documentPackage.applicable}"`;
      
            if (propertyExists(documentPackage, 'applicabilityExceptions')) {
                console.log("made it here (formatServiceResults) 10");
                retVal += `,"applicabilityExceptions":"`;
                documentPackage.applicabilityExceptions.forEach((exception, index) => {
                retVal += `${index ? '<br>' : ''}${exception}`;
                });
                retVal += '"';
            }
      
            console.log("made it here (formatServiceResults) 11");
            retVal += `,"disclosureType":"${documentPackage.disclosureType}"`;
            if (documentPackage.applicabilityEvaluatorClassName != null) {
                retVal += `,"applicabilityEvaluatorClassName":"${documentPackage.applicabilityEvaluatorClassName}"`;
            }
            retVal += '}';
            });
        } else if (propertyExists(resultsJSON, 'result.serviceResults.feeResults')) {
            condenseJSON(resultsJSON);
            retVal = JSON.stringify(resultsJSON.result.serviceResults);
        } else if (propertyExists(resultsJSON, 'result.serviceResults.documentValidationRules')) {
            condenseJSON(resultsJSON);
            resultsJSON.result.serviceResults.documentValidationRules.forEach(documentValidation => {
                retVal += `${retVal ? ',' : ''}"${documentValidation.ruleId} - ${documentValidation.description}":${JSON.stringify(documentValidation.ruleItems)}`;
            });
        } else if (propertyExists(resultsJSON, 'result.serviceResults.dataMappingItems')) {
            condenseJSON(resultsJSON);
            resultsJSON.result.serviceResults.dataMappingItems.forEach(dataMappingItem => {
                retVal += `${retVal ? ',' : ''}"${dataMappingItem.key}":"${dataMappingItem.value}"`;
            });
        } else if (propertyExists(resultsJSON, 'result.serviceResults.decisionAuthorizationItems')){
            condenseJSON(resultsJSON);
            retVal = JSON.stringify(resultsJSON.result.serviceResults);
        } else if (propertyExists(resultsJSON, 'result.serviceResults.pricingResults')) {
            condenseJSON(resultsJSON);
            retVal = JSON.stringify(resultsJSON.result.serviceResults);
        } else if (propertyExists(resultsJSON, 'result.serviceResults.dateItems')) {
            condenseJSON(resultsJSON);
            resultsJSON.result.serviceResults.dateItems.forEach(dateItem => {
                const bExpirationData = (dateItem.identifier + '').indexOf('EXPIRATION_DATE') > -1;
                retVal += `${retVal ? ',' : ''}"${bExpirationData ? '[' : ''}${dateItem.identifier}${bExpirationData ? ']' : ''}":"${(dateItem.value + '').substring(0, 10)}"`;
            });
        }
      
        return `${retVal.startsWith('{') ? '' : '{'}${retVal}${retVal.startsWith('{') ? '' : '}'}`;
      }

    
      
      async function fetchServiceResultsJSON(serviceParams, loanNumber, inputDataJson) {

        let resultsJSON;
        const accessToken = await getAccessToken(serviceParams);  // Assuming this is an async function
      
        const url = `https://${serviceParams[0]}/lendingservices/externalServices/${serviceParams[3]}`;
        const options = { 
            method: 'post',
            headers: {
                'Content-Type': 'application/json',
                'Authorization': `Bearer ${accessToken}`
                },
            body: JSON.stringify({
                bypassLogging: true,
                ruleEngineBeanName: serviceParams[5],
                inputData: JSON.parse(inputDataJson)
                })
            };
      
        console.log(url);
        console.log(options);
      
        try {
            const response = await fetch(url, options);
          if (response.ok) {  // Checks if the status code is 2xx
            resultsJSON = await response.json();  // Parses the JSON response body
          } else {
                console.log(`Validate request returned HTTP status code: ${response.status}`);
          }
        } catch (error) {
            console.error(`Error making request: ${error}`);
        }
      
        return resultsJSON;
      }
      

})();
