// Working Excel Integration
const msalConfig = {
    auth: {
        clientId: CONFIG.CLIENT_ID,
        authority: 'https://login.microsoftonline.com/common',
        redirectUri: getRedirectUri()
    }
};

// Function to determine correct redirect URI based on environment
function getRedirectUri() {
    const origin = window.location.origin;
    
    // For local development, ensure we use the standard localhost URI that should be registered in Azure AD
    if (origin.includes('localhost') || origin.includes('127.0.0.1')) {
        return 'http://localhost:8888';
    }
    
    // For Netlify deployments and production, use the current origin
    return origin;
}

const WORKBOOK_ID = CONFIG.WORKBOOK_ID;
const msalInstance = new msal.PublicClientApplication(msalConfig);
let account = null;

// DOM Elements
const formSection = document.getElementById('formSection');
const dashboardSection = document.getElementById('dashboardSection');
const showFormBtn = document.getElementById('showForm');
const showDashboardBtn = document.getElementById('showDashboard');
const assetForm = document.getElementById('assetForm');
const searchInput = document.getElementById('searchInput');
const assetTableBody = document.getElementById('assetTableBody');
const messageDiv = document.getElementById('message');
const signInBtn = document.getElementById('signInBtn');
const userInfo = document.getElementById('userInfo');

// Initialize
document.addEventListener('DOMContentLoaded', async function() {
    document.getElementById('date').valueAsDate = new Date();
    
    // Debug: Log the current configuration
    console.log('MSAL Configuration:', {
        clientId: msalConfig.auth.clientId,
        redirectUri: msalConfig.auth.redirectUri,
        authority: msalConfig.auth.authority,
        currentOrigin: window.location.origin
    });
    
    // Initialize MSAL and handle any redirect response
    await msalInstance.initialize();
    
    // Handle redirect response (for when user returns from Microsoft login)
    try {
        const response = await msalInstance.handleRedirectPromise();
        if (response) {
            account = response.account;
            console.log('Authentication successful via redirect:', response.account.username);
        } else {
            account = msalInstance.getActiveAccount();
            if (account) {
                console.log('Found cached account:', account.username);
            }
        }
    } catch (error) {
        console.error('Error handling redirect response:', error);
    }
    
    updateUI();
    
    showFormBtn.addEventListener('click', showForm);
    showDashboardBtn.addEventListener('click', showDashboard);
    assetForm.addEventListener('submit', handleFormSubmit);
    searchInput.addEventListener('input', filterTable);
    signInBtn.addEventListener('click', signIn);
    
    loadAssets();
});

// Authentication
async function signIn() {
    try {
        const response = await msalInstance.loginPopup({
            scopes: ['https://graph.microsoft.com/Files.ReadWrite.All'],
            prompt: 'select_account'  // Force account selection to avoid cached issues
        });
        account = response.account;
        updateUI();
        showMessage('Signed in successfully!', 'success');
        loadAssets();
    } catch (error) {
        console.error('Sign-in failed:', error);
        
        // Provide more specific error messages
        let errorMessage = 'Sign-in failed. ';
        if (error.errorCode === 'user_cancelled') {
            errorMessage += 'Sign-in was cancelled.';
        } else if (error.errorCode === 'consent_required') {
            errorMessage += 'Admin consent required for this application.';
        } else if (error.errorCode === 'invalid_request') {
            errorMessage += 'Invalid configuration. Please check the app registration.';
        } else {
            errorMessage += `Error: ${error.errorMessage || error.message}`;
        }
        
        showMessage(errorMessage, 'error');
    }
}

function updateUI() {
    if (account) {
        signInBtn.textContent = 'Sign Out';
        signInBtn.onclick = signOut;
        userInfo.textContent = `Signed in as: ${account.name}`;
    } else {
        signInBtn.textContent = 'Sign in to Microsoft';
        signInBtn.onclick = signIn;
        userInfo.textContent = '';
    }
}

async function signOut() {
    try {
        await msalInstance.logoutPopup();
        account = null;
        updateUI();
        showMessage('Signed out successfully!', 'success');
        displaySampleData();
    } catch (error) {
        console.error('Sign-out failed:', error);
    }
}

// Get access token
async function getAccessToken() {
    if (!account) throw new Error('Not signed in');
    
    try {
        const response = await msalInstance.acquireTokenSilent({
            scopes: ['https://graph.microsoft.com/Files.ReadWrite.All'],
            account: account
        });
        return response.accessToken;
    } catch (error) {
        console.log('Silent token acquisition failed, trying popup:', error);
        try {
            const response = await msalInstance.acquireTokenPopup({
                scopes: ['https://graph.microsoft.com/Files.ReadWrite.All'],
                account: account
            });
            return response.accessToken;
        } catch (popupError) {
            console.error('Token acquisition failed:', popupError);
            throw new Error('Failed to acquire access token. Please sign in again.');
        }
    }
}

// Navigation
function showForm() {
    formSection.classList.add('active');
    dashboardSection.classList.remove('active');
    showFormBtn.classList.add('active');
    showDashboardBtn.classList.remove('active');
}

function showDashboard() {
    dashboardSection.classList.add('active');
    formSection.classList.remove('active');
    showDashboardBtn.classList.add('active');
    showFormBtn.classList.remove('active');
    loadAssets();
}

// Form submission
async function handleFormSubmit(e) {
    e.preventDefault();
    
    if (!validateForm()) return;
    
    if (!account) {
        showMessage('Please sign in to Microsoft first', 'error');
        return;
    }
    
    const formData = new FormData(assetForm);
    const data = {};
    
    for (let [key, value] of formData.entries()) {
        if (key === 'items') {
            data[key] = data[key] ? data[key] + ', ' + value : value;
        } else {
            data[key] = value;
        }
    }
    
    try {
        await saveToExcel(data);
        showMessage('Asset saved to Excel successfully!', 'success');
        assetForm.reset();
        document.getElementById('date').valueAsDate = new Date();
        loadAssets();
    } catch (error) {
        console.error('Error:', error);
        showMessage('Error saving to Excel: ' + error.message, 'error');
    }
}

// Save to Excel
async function saveToExcel(data) {
    const accessToken = await getAccessToken();
    
    // Get the next row number
    let nextRow = 2;
    try {
        const rangeResponse = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${WORKBOOK_ID}/workbook/worksheets/Sheet1/usedRange`, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        
        if (rangeResponse.ok) {
            const rangeData = await rangeResponse.json();
            nextRow = rangeData.rowCount + 1;
        }
    } catch (error) {
        console.log('Using default row 2');
    }
    
    // Prepare data row
    const rowData = [
        new Date().toISOString(),
        data.date,
        data.initials,
        data.surname,
        data.contact,
        data.school,
        data.purpose,
        data.serialNumber || '',
        data.brand,
        data.fault,
        data.items || '',
        data.password || '',
        data.notes || '',
        data.sender,
        data.receiver
    ];
    
    // Add headers if this is the first row
    if (nextRow === 2) {
        const headers = ['Timestamp', 'Date', 'Initials', 'Surname', 'Contact', 'School', 'Purpose', 'Serial Number', 'Brand & Model', 'Fault Description', 'Items Included', 'Password', 'Notes', 'Sender', 'Receiver'];
        
        await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${WORKBOOK_ID}/workbook/worksheets/Sheet1/range(address='A1:O1')`, {
            method: 'PATCH',
            headers: {
                'Authorization': `Bearer ${accessToken}`,
                'Content-Type': 'application/json'
            },
            body: JSON.stringify({ values: [headers] })
        });
    }
    
    // Add the data row
    const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${WORKBOOK_ID}/workbook/worksheets/Sheet1/range(address='A${nextRow}:O${nextRow}')`, {
        method: 'PATCH',
        headers: {
            'Authorization': `Bearer ${accessToken}`,
            'Content-Type': 'application/json'
        },
        body: JSON.stringify({ values: [rowData] })
    });
    
    if (!response.ok) {
        const error = await response.json();
        throw new Error(error.error?.message || 'Failed to save to Excel');
    }
}

// Load assets
async function loadAssets() {
    if (!account) {
        displaySampleData();
        return;
    }
    
    try {
        const accessToken = await getAccessToken();
        const response = await fetch(`https://graph.microsoft.com/v1.0/me/drive/items/${WORKBOOK_ID}/workbook/worksheets/Sheet1/usedRange`, {
            headers: { 'Authorization': `Bearer ${accessToken}` }
        });
        
        if (response.ok) {
            const result = await response.json();
            const values = result.values;
            
            if (values && values.length > 1) {
                const headers = values[0];
                const rows = values.slice(1);
                
                const assets = rows.map(row => {
                    const asset = {};
                    headers.forEach((header, index) => {
                        asset[header.toLowerCase().replace(/\s+/g, '')] = row[index] || '';
                    });
                    return asset;
                });
                
                displayAssets(assets);
                return;
            }
        }
    } catch (error) {
        console.error('Error loading from Excel:', error);
    }
    
    displaySampleData();
}

// Display functions
function displayAssets(assets) {
    assetTableBody.innerHTML = '';
    
    assets.forEach(asset => {
        const row = document.createElement('tr');
        row.classList.add('status-in-progress');
        
        row.innerHTML = `
            <td>${formatDate(asset.date)}</td>
            <td>${asset.initials} ${asset.surname}</td>
            <td>${asset.contact}</td>
            <td>${asset.school}</td>
            <td>${asset.purpose}</td>
            <td>${asset['serialnumber'] || asset.serialNumber || ''}</td>
            <td>${asset['brand&model'] || asset.brand}</td>
            <td>${asset['faultdescription'] || asset.fault}</td>
            <td>${asset['itemsincluded'] || asset.items || 'None'}</td>
            <td>${asset.sender || ''}</td>
            <td>${asset.receiver || ''}</td>
            <td><span class="status">In Progress</span></td>
        `;
        
        assetTableBody.appendChild(row);
    });
}

function displaySampleData() {
    const sampleData = [
        {
            date: '2024-01-15',
            initials: 'J',
            surname: 'Smith',
            contact: '123-456-7890',
            school: 'Central High School',
            purpose: 'Repair',
            serialNumber: 'DL123456789',
            brand: 'Dell Latitude 5520',
            fault: 'Screen flickering issue',
            items: 'Laptop Bag, Charger',
            sender: 'Tech Support',
            receiver: 'J. Smith'
        }
    ];
    
    displayAssets(sampleData);
}

function formatDate(dateString) {
    if (!dateString) return '';
    const date = new Date(dateString);
    return date.toLocaleDateString();
}

function filterTable() {
    const searchTerm = searchInput.value.toLowerCase();
    const rows = assetTableBody.querySelectorAll('tr');
    
    rows.forEach(row => {
        const text = row.textContent.toLowerCase();
        row.style.display = text.includes(searchTerm) ? '' : 'none';
    });
}

function validateForm() {
    const required = ['date', 'initials', 'surname', 'contact', 'school', 'serialNumber', 'brand', 'fault', 'sender', 'receiver'];
    const purpose = document.querySelector('input[name="purpose"]:checked');
    
    for (let field of required) {
        const element = document.getElementById(field);
        if (!element.value.trim()) {
            showMessage(`Please fill in ${field.replace(/([A-Z])/g, ' $1').toLowerCase()}`, 'error');
            element.focus();
            return false;
        }
    }
    
    if (!purpose) {
        showMessage('Please select a purpose', 'error');
        return false;
    }
    
    return true;
}

function showMessage(text, type) {
    messageDiv.textContent = text;
    messageDiv.className = `message ${type}`;
    messageDiv.classList.add('show');
    
    setTimeout(() => {
        messageDiv.classList.remove('show');
    }, 3000);
}