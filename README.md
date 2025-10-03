# ICT Asset Management Portal

A responsive web application for managing ICT Asset Log Book data using HTML, CSS, JavaScript, and Microsoft Excel as a database.

## Quick Setup

1. **Copy config file**: `cp config.example.js config.js`
2. **Edit config.js** with your Azure Client ID and Excel Workbook ID
3. **Open working-excel.html** in your browser
4. **Sign in and start using!**

## Features

- **Responsive Design**: Works on desktop and mobile devices
- **Data Entry Form**: Clean form matching physical log book requirements
- **Asset Dashboard**: View and search all assets with status tracking
- **Real-time Status**: Visual indicators for In Progress vs Completed items
- **Search & Filter**: Dynamic filtering by School, Surname, or Model
- **Microsoft Excel Integration**: Uses Excel Online with Microsoft Graph API
- **Microsoft Authentication**: Secure sign-in with Microsoft accounts

## Setup Instructions

### 1. Azure App Registration

1. Go to [Azure Portal](https://portal.azure.com)
2. Navigate to "Azure Active Directory" > "App registrations"
3. Click "New registration"
4. Name: "ICT Asset Portal"
5. Supported account types: "Accounts in any organizational directory and personal Microsoft accounts"
6. Redirect URI: Your website URL (e.g., `https://yoursite.com`)
7. Click "Register"
8. Copy the "Application (client) ID"

### 2. Excel Workbook Setup

1. Create a new Excel file in OneDrive or SharePoint
2. Note the file ID from the URL (the long string after `/edit/`)
3. Optionally add an Office Script using `office-script.js`

### 3. Frontend Configuration

1. Copy `config.example.js` to `config.js`
2. Edit `config.js` and replace:
   - `YOUR_AZURE_CLIENT_ID_HERE` with your Azure App Client ID
   - `YOUR_EXCEL_WORKBOOK_ID_HERE` with your Excel file ID
3. Save the file

### 3. Hosting

Upload all files (`index.html`, `style.css`, `script.js`) to your web hosting service or open `index.html` directly in a browser for local testing.

## File Structure

```
ICT Asset/
├── index.html              # Basic HTML structure
├── excel-integration.html  # Enhanced HTML with Microsoft auth
├── style.css              # Responsive CSS styling
├── script.js              # Basic JavaScript (Google Sheets version)
├── excel-script.js        # Microsoft Graph API JavaScript
├── office-script.js       # Office Script for Excel Online
└── README.md              # This file
```

## Form Fields

- Date (auto-filled with current date)
- Initials & Surname
- Contact Number
- School Name
- Purpose (Repair/Loan)
- Brand & Model
- Fault Description/Details
- Items Included (checkboxes)
- Password
- Notes

## Status Logic

- **In Progress** (Yellow): Collection fields are empty
- **Completed** (Green): Collection fields have data

## Browser Compatibility

- Chrome 60+
- Firefox 55+
- Safari 12+
- Edge 79+

## Security Notes

- Microsoft Graph API handles authentication and authorization
- All data is stored in your Excel file in OneDrive/SharePoint
- Uses OAuth 2.0 with Microsoft identity platform
- Secure token-based authentication
- No sensitive data is stored in the frontend code

## Customization

- Modify colors in `style.css` for branding
- Add additional form fields by updating both HTML and Apps Script
- Extend status logic in JavaScript as needed
- Add data validation rules in the Apps Script

## Usage

### For Excel Integration:
1. Use `excel-integration.html` as your main file
2. Users must sign in with Microsoft account
3. Data is automatically saved to your Excel file

### For Basic Version:
1. Use `index.html` for offline/demo mode
2. Shows sample data without authentication

## Troubleshooting

1. **Authentication issues**: Verify Azure App Registration settings
2. **Excel access denied**: Check file permissions in OneDrive/SharePoint
3. **CORS errors**: Ensure redirect URI matches your hosting URL
4. **Data not saving**: Verify Excel workbook ID is correct
5. **Mobile display issues**: Check viewport meta tag and CSS media queries