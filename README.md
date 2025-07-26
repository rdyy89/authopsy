# Authopsy - Outlook Add-in

An Outlook add-in that displays DMARC, DKIM, and SPF authentication results for emails.

## 🚀 Quick Start (Local Development)

### Prerequisites
- Python 3.x installed on your system
- Outlook (Desktop or Web)

### Setup Steps

1. **Start the local server:**
   ```bash
   # On Windows:
   serve.bat
   
   # On Mac/Linux:
   ./serve.sh
   ```
   This will start a web server at `http://localhost:8000`

2. **Install the add-in in Outlook:**
   - In Outlook, go to **File** > **Options** > **Add-ins**
   - Click **Manage Office Add-ins**
   - Click **Custom Add-ins** > **Add from file**
   - Select `manifest-local.xml` from this folder
   - Click **Install**

3. **Test the add-in:**
   - Open any email in Outlook
   - Look for the "Authopsy" button in the ribbon
   - Click it to see the authentication results

## 🌐 Production Deployment

For production use, you need to host the files on a web server. Here are the recommended options:

### Option 1: GitHub Pages (Free & Easy)
1. Push this code to a GitHub repository
2. Enable GitHub Pages in repository settings
3. Update `manifest.xml` with your GitHub Pages URL
4. Install using the updated manifest

### Option 2: Other Hosting Services
- **Netlify**: Drag and drop deployment
- **Vercel**: GitHub integration
- **Azure Static Web Apps**: Microsoft's hosting service

## 📁 File Structure

```
authopsy/
├── manifest.xml              # Production manifest (needs hosting URL)
├── manifest-local.xml        # Local development manifest
├── serve.bat                 # Start server on Windows
├── serve.sh                  # Start server on Mac/Linux
├── assets/
│   ├── tick.png             # Success icon
│   └── cross.png            # Failure icon
└── src/
    ├── taskpane/
    │   ├── taskpane.html    # Main UI
    │   ├── taskpane.js      # Core functionality
    │   └── taskpane.css     # Styling
    └── commands/
        ├── commands.html    # Command functions
        └── commands.js      # Command handlers
```

## 🔧 Features

- **DMARC**: Domain-based Message Authentication, Reporting & Conformance
- **DKIM**: DomainKeys Identified Mail
- **SPF**: Sender Policy Framework
- **Loading States**: Visual feedback while processing
- **Error Handling**: Graceful fallbacks when headers aren't available
- **Accessibility**: High contrast and screen reader support

## 🐛 Troubleshooting

### "X-Frame-Options" Error
This happens when trying to load files from GitHub raw URLs. Use the local development setup above or host on a proper web server.

### Add-in Not Appearing
1. Check that the server is running (`http://localhost:8000` should be accessible)
2. Verify the manifest is installed correctly
3. Try restarting Outlook
4. Check the browser console for errors

### No Authentication Results
This is normal for some emails. The add-in will show demo data when real authentication headers aren't available.

## 🔒 Security

The add-in only reads email headers and doesn't modify or send any data. All processing happens locally in your browser.

## 📝 Development

To modify the add-in:
1. Edit files in the `src/` directory
2. Refresh Outlook or reload the add-in
3. Changes are reflected immediately (no rebuild needed)

## 📄 License

MIT License - Feel free to use and modify as needed.
