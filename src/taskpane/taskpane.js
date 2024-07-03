(function () {
    let currentDialog = null;
  
    Office.onReady((info) => {
        if (info.host === Office.HostType.Excel) {
            checkLoginState();
        }
    });

    function checkLoginState() {
        const loggedIn = localStorage.getItem('loggedIn') === 'true';
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
                console.log("made it here 1");
                // Ensure the open-dialog-btn exists before adding an event listener
                const openDialogButton = document.getElementById('open-dialog-btn');
                const openRulesConditions = document.getElementById('rulesConditions');
                const openGoogleButton = document.getElementById('openUrl');
                // const openRemoteRules = document.getElementById('remoteRules');
                const dropdown = document.getElementById('dropdown-button');

                openDialogButton.addEventListener('click', openDialog);
                openRulesConditions.addEventListener('click', rulesConditionsWindow);    
                openGoogleButton.addEventListener('click', urlWindow); 
                // openRemoteRules.addEventListener('click', openRemoteRulesUI);
                dropdown.addEventListener('click', toggleDropdown);
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

    function openRemoteRulesUI() {
        if (!window.location.pathname.endsWith('remoteRulesHome.html')) {
            window.location.href = 'https://peytonchang.github.io/BSSLoanExceptionRulesEditor/src/taskpane/RemoteRules/remoteRulesHome.html';
        }
    }

    function toggleDropdown() {
        console.log("made it here 2");
        document.getElementById("dropdown-content").classList.toggle("show");
    }

    window.onclick = function(event) {
        console.log("made it here 3");
        if (!event.target.matches('.dropbtn')) {
            var dropdowns = document.getElementsByClassName("dropdown-content");
            for (var i = 0; i < dropdowns.length; i++) {
                var openDropdown = dropdowns[i];
                if (openDropdown.classList.contains('show')) {
                    openDropdown.classList.remove('show');
                }
            }
        }
    }

  })();