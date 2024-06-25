(function () {
    let currentDialog = null;
  
    Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
            checkLoginState();
        }
    });

    function checkLoginState() {
        const loggedIn = localStorage.getItem('loggedIn') === 'true';
        console.log("loggedIn val: " + loggedIn)
        if (!loggedIn) {
            // Ensure the loginButton exists before adding an event listener
            const loginButton = document.getElementById('loginButton');
            if (loginButton) {
                loginButton.addEventListener('click', login);
            }
        } else {
            if (!window.location.pathname.endsWith('home.html')) {
                window.location.href = 'home.html';
            } else {
                // Ensure the open-dialog-btn exists before adding an event listener
                const openDialogButton = document.getElementById('open-dialog-btn');
                const openRulesConditions = document.getElementById('rulesConditions');
                const openGoogleButton = document.getElementById('openUrl');

                openDialogButton.addEventListener('click', openDialog);
                openRulesConditions.addEventListener('click', rulesConditionsWindow);    
                openGoogleButton.addEventListener('click', urlWindow); 
            }
        }
    }
  
    function login() {
        const enteredPassword = document.getElementById('passwordInput').value;
        const universalPassword = "bluesage123";

        if (enteredPassword === universalPassword) {
            localStorage.setItem('loggedIn', 'true');
            window.location.href = 'home.html';
            checkLoginState();
        } else {
            document.getElementById('errorMessage').style.display = 'block';
        }
    }

    function openDialog() {

        const dialogUrl = 'https://peytonchang.github.io/BSSLoanExceptionRulesEditor/src/dialog.html'; // Adjust as necessary
        Office.context.ui.displayDialogAsync(dialogUrl, { height: 50, width: 50 }, (result) => {
            if (result.status === Office.AsyncResultStatus.Failed) {
                console.error('Failed to open dialog: ' + result.error.message);
            } else {
                currentDialog = result.value;
                console.log('Dialog opened successfully.');
                currentDialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessageFromDialog);
                currentDialog.addEventHandler(Office.EventType.DialogEventReceived, handleDialogEvent);
            }
        });
    }
  
    function handleDialogEvent(event) {
        if (event.type === "dialogClosed") {
            currentDialog = null;
            console.log('Dialog closed.');
        }
    }

    function rulesConditionsWindow() {
        if (!window.location.pathname.endsWith('RulesConditions.html')) {
            window.location.href = 'RuleConditions.html';
        }
    }

    function urlWindow() {
        if (!window.location.pathname.endsWith('dialog.html')) {
            window.location.href = 'https://peytonchang.github.io/BSSLoanExceptionRulesEditor/src/dialog.html';
        }
    }

  })();