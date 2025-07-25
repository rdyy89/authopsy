# Email Authentication Checker - Outlook Add-in

An Outlook add-in that provides real-time SPF, DKIM, and DMARC authentication checking for emails across all Outlook platforms.

## Features

- ✅ **SPF Validation**: Checks Sender Policy Framework compliance
- 🔐 **DKIM Verification**: Validates DomainKeys Identified Mail signatures  
- 🛡️ **DMARC Analysis**: Analyzes Domain-based Message Authentication policies
- 📊 **Security Score**: Provides an overall security rating (0-100)
- 🌐 **Cross-Platform**: Works with Outlook Web, New Outlook, and Outlook Classic
- 📱 **Responsive**: Optimized for different screen sizes

## Installation

### Prerequisites
- Node.js 14 or higher
- npm or yarn package manager

### Development Setup

1. **Clone the repository**
   ```bash
   git clone <your-repo-url>
   cd authopsy
   ```

2. **Install dependencies**
   ```bash
   npm install
   ```

3. **Start development server**
   ```bash
   npm run dev-server
   ```

4. **Sideload the add-in**
   ```bash
   npm run sideload
   ```

### Building for Production

```bash
npm run build
```

## Usage

1. **Open an email** in Outlook (Web, Desktop, or Mobile)
2. **Click the "Check Authentication" button** in the ribbon
3. **View the results** in the task pane showing:
   - SPF status and details
   - DKIM verification results
   - DMARC compliance analysis
   - Overall security score with recommendations

## How It Works

The add-in analyzes email headers to extract authentication information:

- **SPF**: Checks `Authentication-Results` and `Received-SPF` headers
- **DKIM**: Examines `DKIM-Signature` and authentication results
- **DMARC**: Analyzes DMARC policy compliance from headers

### Security Score Calculation

- **SPF**: 30 points (Pass=30, Soft Fail=15, Neutral=10)
- **DKIM**: 35 points (Pass=35, Neutral=10)  
- **DMARC**: 35 points (Pass=35, Quarantine=15, None=5)

## Browser Compatibility

- ✅ Outlook on Windows (2016, 2019, 2021, Microsoft 365)
- ✅ Outlook on Mac
- ✅ Outlook on the Web
- ✅ New Outlook for Windows
- ✅ Outlook Mobile (iOS/Android)

## File Structure

```
authopsy/
├── manifest.xml           # Add-in manifest
├── src/
│   ├── taskpane/
│   │   ├── taskpane.html  # Main UI
│   │   ├── taskpane.css   # Styles
│   │   └── taskpane.js    # Core logic
│   └── assets/            # Icons and images
├── webpack.config.js      # Build configuration
└── package.json          # Dependencies
```

## Development

### Available Scripts

- `npm run build` - Build for production
- `npm run build:dev` - Build for development
- `npm run dev-server` - Start development server
- `npm run start` - Start add-in debugging
- `npm run stop` - Stop add-in debugging
- `npm run validate` - Validate manifest
- `npm run sideload` - Sideload add-in

### Testing

To test the add-in:

1. Start the development server: `npm run dev-server`
2. Sideload the manifest in Outlook
3. Open any email and click the "Check Authentication" button

### Customization

You can customize the add-in by modifying:

- **UI**: Edit `src/taskpane/taskpane.html` and `taskpane.css`
- **Logic**: Update authentication checking in `src/taskpane/taskpane.js`
- **Manifest**: Configure add-in properties in `manifest.xml`

## Deployment

### Option 1: Microsoft AppSource
1. Build the production version
2. Package the add-in
3. Submit to Microsoft AppSource

### Option 2: Organization Deployment
1. Deploy to your organization's catalog
2. Configure manifest URLs to point to your hosting location
3. Install for users via admin center

### Option 3: Self-Hosting
1. Host the built files on HTTPS server
2. Update manifest URLs
3. Distribute manifest file to users

## Security Considerations

- All communication is over HTTPS
- No email content is stored or transmitted
- Only email headers are analyzed
- Complies with Office Add-in security requirements

## Troubleshooting

### Common Issues

1. **Add-in not loading**: Check HTTPS certificates and manifest URLs
2. **Headers not accessible**: Ensure proper Office.js API permissions
3. **Authentication results missing**: Some emails may not have complete headers

### Debug Mode

Enable debugging:
```bash
npm run start
```

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly across platforms
5. Submit a pull request

## License

MIT License - see LICENSE file for details

## Support

For support and questions:
- Create an issue in the GitHub repository
- Check the Office Add-ins documentation
- Review the Office.js API reference

## Changelog

### v1.0.0
- Initial release
- SPF, DKIM, DMARC checking
- Cross-platform compatibility
- Security scoring system
