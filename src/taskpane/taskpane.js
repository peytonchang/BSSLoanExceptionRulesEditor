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

    async function openDialog() {
        const url = 'https://bluesage-dev.bluesageusa.com/droolsrules/RuleEditor-Ex.html'; // Ensure this URL is correct

        fetch(url)
            .then(response => {
                if (!response.ok) {
                    throw new Error(`HTTP error! Status: ${response.status}`);
                }
                return response.text();
            })
            .then(htmlContent => {
                const contentContainer = document.getElementById('content-container');
                contentContainer.innerHTML = htmlContent; // Replace the content
                // Optional: If you want to hide other elements, manage their visibility here
            })
            .catch(error => {
                console.error('Failed to load rule conditions:', error);
                document.getElementById('content-container').innerHTML = 'Failed to load content. Please try again later.';
            });
            try {
                const response = await fetch(url);
                if (!response.ok) {
                    throw new Error(`HTTP error! status: ${response.status}`);
                }
                const data = await response.text();
                return data; // This returns the content as a string
            } catch (error) {
                console.error('Fetch error:', error);
                return null; // Return null or handle the error as needed
            }
        
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

    function toggleDropdown() {
        const dropdownContent = document.getElementById("dropdown-content");
        if (dropdownContent) {
            if (!dropdownContent.classList.contains("show")) {
                dropdownContent.style.maxHeight = dropdownContent.scrollHeight + "px";
            } else {
                dropdownContent.style.maxHeight = null;
            }
            dropdownContent.classList.toggle("show");
        }
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