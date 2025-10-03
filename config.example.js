// Configuration example - Copy this to config.js and add your values
const CONFIG = {
    CLIENT_ID: '286edccc-cb3b-4841-887c-5b445b17c9b9', // Your Azure App Registration Client ID
    WORKBOOK_ID: 'be7c3aa5-eb17-4042-a442-1df0d5cf5c62' // Your Excel Workbook ID from OneDrive URL
};

// Export for use in other files
if (typeof module !== 'undefined' && module.exports) {
    module.exports = CONFIG;
}