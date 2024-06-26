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
        // Assume we have a way to fetch the active worksheet and its data
        const oWks = await getActiveWorksheet();
        const aData = await oWks.getRangeData(1, 1, 1, await oWks.getLastColumn());
    
        const dicColumn = getColumnDictionary(aData);
        const iRow = await oWks.getActiveCellRow();
        
        const loanInputDataCell = await oWks.getCell(iRow, dicColumn['Loan Input Data'] + 1);
        const lockCell = await oWks.getCell(iRow, dicColumn['Lock'] + 1);
    
        if (loanInputDataCell !== '' && lockCell === true) {
            alert('Loan Input Data Locked', 'Cannot get Loan Input Data.'); // Adjust this to match your UI handling
        } else {
            getLoanInputData();
            viewLoanInputData();
        }
    }

    function executeRules() {

    }

    async function viewLoanInputData() {
        const oWks = await getActiveWorksheet();
        const aData = await oWks.getRangeData(1, 1, 1, await oWks.getLastColumn());
        const dicColumn = getColumnDictionary(aData);
        const activeRow = await oWks.getActiveCellRow();
    
        const environment = await oWks.getCell(activeRow, dicColumn['Environment'] + 1);
        const ruleProject = await oWks.getCell(activeRow, dicColumn['Rule Project'] + 1);
        const loanNumber = await oWks.getCell(activeRow, dicColumn['Loan #'] + 1);
        const loanInputData = await oWks.getCell(activeRow, dicColumn['Loan Input Data'] + 1);
    
        showJSON(loanInputData, `${ruleProject} Input Data Viewer - ${environment} #${loanNumber}`);
    }

    function viewResultsData() {

    }

    async function getLoanInputData() {
        const oWks = await getActiveWorksheet();
        const aData = await oWks.getRangeData(1, 1, 1, await oWks.getLastColumn());
        const dicColumn = getColumnDictionary(aData);
        const iRow = await oWks.getActiveCellRow();
        const environment = await oWks.getCell(iRow, dicColumn['Environment'] + 1);
        const ruleProject = await oWks.getCell(iRow, dicColumn['Rule Project'] + 1);
        const loanNumber = await oWks.getCell(iRow, dicColumn['Loan #'] + 1);
    
        const serviceParams = getServiceParams(environment, ruleProject);
    
        if (serviceParams.length > 0) {
            const resultsJSON = await fetchServiceInputDataJSON(serviceParams, loanNumber);
            let inputDataJSON = null;
    
            if (propertyExists(resultsJSON, 'parameters.loanData')) {
                inputDataJSON = resultsJSON.parameters.loanData.value;
    
                if (ruleProject === 'getLoanExceptions') {
                    inputDataJSON.fullEligibility = true;
                }
            } else if (ruleProject === 'Pricing' && propertyExists(resultsJSON, 'result.inputData')) {
                inputDataJSON = resultsJSON.result.inputData;
            }
    
            if (inputDataJSON != null) {
                condenseJSON(inputDataJSON);
    
                if (propertyExists(inputDataJSON, 'rateLockDate')) {
                    const effectiveDate = formatDate(new Date(inputDataJSON.rateLockDate), 'yyyy-MM-dd');
                    await oWks.setCell(iRow, dicColumn['Effective Date'] + 1, effectiveDate);
                }
    
                const currentTimeStamp = formatDate(new Date(), 'MM-dd-yyyy hh:mm:ss a');
                await oWks.setCell(iRow, dicColumn['Input Data Timestamp'] + 1, currentTimeStamp);
                await oWks.setCell(iRow, dicColumn['Loan Input Data'] + 1, JSON.stringify(inputDataJSON));
            }
        }
    }

    async function fetchServiceInputDataJSON(serviceParams, loanNumber) {
        let inputData = '';
        const accessToken = await getAccessToken(serviceParams);  // Assuming getAccessToken is an async function
        const headers = new Headers({
            "Content-Type": "application/json",
            "Authorization": `Bearer ${accessToken}`
        });
    
        let url = `https://${serviceParams[0]}/lendingservices/externalServices/`;
        let options = {
            method: "GET",
            headers: headers,
            muteHttpExceptions: true  // This option does not exist in the standard fetch API
        };
    
        if (serviceParams[3] === 'getPricing') {
            url += 'getPricingInputData';
            options.body = JSON.stringify({ "bypassLogging": true, "inputData": { "loanNumber": loanNumber } });
            options.method = "POST"; // Changed to POST since we are sending a payload
        } else {
            url += `getRuleEngineParameters?loanNumber=${loanNumber}&beanName=${serviceParams[4]}`;
        }
    
        try {
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
    
        return inputData;
    }

    function showJSON(json, title) {
        // Escape single quotes in the JSON string
        const safeJson = json.replace(/'/g, "\\'");
    
        // Create a full HTML page content
        const htmlContent = `
            <html>
            <head>
                <base target="_top">
                <link rel="stylesheet" type="text/css" href="jsonTree.css">
                <script src="jsonTree.js"></script>
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
    
        // Display the HTML in a new window or a modal
        const newWindow = window.open("", title, "width=1100,height=700");
        newWindow.document.write(htmlContent);
        newWindow.document.close(); // Close the document to parse the HTML
    }
    
    function closeModal() {
        window.close(); // This will close the modal window
    }
    
    function getServiceParams(environment, ruleProject) {
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
        const columnDictionary = {};  // Use an object to store key-value pairs
        if (!aData || !aData[0]) return columnDictionary;  // Return empty object if no data
    
        aData[0].forEach((column, index) => {
            const columnName = column.trim();  // Assuming column can be directly trimmed
            if (columnName !== '') {
                columnDictionary[columnName] = index;
                // console.log(`Col: #${index} - ${columnName}`); // Optionally log the column names and indexes
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


})();
