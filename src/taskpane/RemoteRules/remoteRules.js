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
                    getLoanInputData();
                    viewLoanInputData();
                }
            });
        } catch (error) {
            console.error("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        }
    }

    function executeRules() {

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

    function viewResultsData() {

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
                const environmentCell = sheet.getCell(activeRow, dicColumn['Environment'] + 1);
                const ruleProjectCell = sheet.getCell(activeRow, dicColumn['Rule Project'] + 1);
                const loanNumberCell = sheet.getCell(activeRow, dicColumn['Loan #'] + 1);
    
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
                            await sheet.getCell(activeRow, dicColumn['Effective Date'] + 1).setValues([[effectiveDate]]);
                        }
                        
                        console.log("made it here (getLoanInputData) 17");
                        const currentTimeStamp = formatDate(new Date(), 'MM-dd-yyyy hh:mm:ss a');
                        await sheet.getCell(activeRow, dicColumn['Input Data Timestamp'] + 1).setValues([[currentTimeStamp]]);
                        await sheet.getCell(activeRow, dicColumn['Loan Input Data'] + 1).setValues([[JSON.stringify(inputDataJSON)]]);
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
            const response = await fetch(url, options);
            if (response.status !== 200) {
                console.log(`Validate request returned HTTP status code: ${response.status}`);
            } else {
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

})();
