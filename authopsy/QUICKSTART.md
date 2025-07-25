# Quick Start Guide - Testing the Outlook Add-in

## Prerequisites
1. **Node.js 14+** - Download from [nodejs.org](https://nodejs.org/)
2. **Outlook** - Web, Desktop, or New Outlook
3. **Administrator privileges** (for development certificates)

## Installation Steps

### 1. Install Dependencies
```powershell
npm install
```

### 2. Start Development Server
```powershell
npm run dev-server
```
This will:
- Start a local HTTPS server on port 3000
- Generate development certificates
- Build the add-in files

### 3. Sideload the Add-in

#### Option A: Outlook on the Web
1. Go to [outlook.office.com](https://outlook.office.com)
2. Click the gear icon (Settings) in top right
3. Select "View all Outlook settings"
4. Go to "General" > "Manage add-ins"
5. Click "Add a custom add-in" > "Add from file"
6. Upload the `manifest.xml` file from your project
7. Click "Install"

#### Option B: Outlook Desktop (Windows/Mac)
1. Open Outlook Desktop
2. Go to "Insert" tab in ribbon
3. Click "Get Add-ins" or "Store"
4. Click "My add-ins" 
5. Click "Add a custom add-in" > "Add from file"
6. Select the `manifest.xml` file
7. Click "Install"

#### Option C: New Outlook
1. Open New Outlook
2. Click "Apps" in the toolbar
3. Select "More apps" > "My add-ins"
4. Click "Add a custom add-in" > "Add from file"
5. Upload `manifest.xml`
6. Install the add-in

### 4. Test the Add-in
1. **Open any email** in Outlook
2. **Look for the "Check Authentication" button** in the ribbon
3. **Click the button** to open the task pane
4. **View the authentication results** for SPF, DKIM, and DMARC

## Troubleshooting

### Common Issues:

1. **"Add-in won't load"**
   - Ensure dev server is running (`npm run dev-server`)
   - Check if `https://localhost:3000` is accessible
   - Accept the self-signed certificate if prompted

2. **"Certificate errors"**
   ```powershell
   npx office-addin-dev-certs install
   ```

3. **"Button not appearing"**
   - Try refreshing Outlook
   - Check if add-in is enabled in settings
   - Restart Outlook

4. **"Headers not loading"**
   - This is normal for some emails
   - The add-in will show mock data for demonstration
   - Try different emails from various senders

### Debug Mode
Add `?debug=true` to see detailed information:
- Raw header values
- Parsing results  
- API availability status

## Production Deployment

### For Organization:
1. Host the built files on your web server
2. Update manifest URLs to your domain
3. Deploy via Microsoft 365 admin center

### For Public Distribution:
1. Submit to Microsoft AppSource
2. Follow Microsoft's validation process
3. Users can install from the store

## Security Notes
- The add-in only reads email headers
- No email content is stored or transmitted
- All processing happens locally
- HTTPS is required for all connections
