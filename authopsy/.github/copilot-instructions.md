<!-- Use this file to provide workspace-specific custom instructions to Copilot. For more details, visit https://code.visualstudio.com/docs/copilot/copilot-customization#_use-a-githubcopilotinstructionsmd-file -->

# Email Authentication Checker - Outlook Add-in

This is an Outlook add-in project that displays SPF, DKIM, and DMARC authentication results for emails. The add-in works across Outlook Web, New Outlook, and Outlook Classic.

## Key Features
- Real-time email authentication checking
- SPF (Sender Policy Framework) validation
- DKIM (DomainKeys Identified Mail) verification
- DMARC (Domain-based Message Authentication) compliance
- Security score calculation
- Cross-platform compatibility

## Development Guidelines
- Use Office.js APIs for Outlook integration
- Follow Microsoft Office Add-in development best practices
- Ensure compatibility with Internet Explorer 11 for Outlook Classic
- Use semantic HTML and accessible design patterns
- Implement proper error handling for email header parsing
- Maintain responsive design for different screen sizes

## Technical Stack
- JavaScript (ES5 compatible for IE11)
- Office.js API
- Webpack for bundling
- Babel for transpilation
- HTML/CSS for UI

## Authentication Logic
- Parse email headers for authentication results
- Extract SPF results from Authentication-Results and Received-SPF headers
- Check DKIM signatures and validation results
- Analyze DMARC policy compliance
- Calculate weighted security score based on all three protocols

When working with this codebase, prioritize:
1. Cross-platform compatibility
2. Proper email header parsing
3. User-friendly error messages
4. Clear visual indicators for authentication status
5. Accessibility compliance
