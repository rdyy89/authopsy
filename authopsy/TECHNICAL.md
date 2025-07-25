# Email Authentication Technical Guide

## Overview
This Outlook add-in analyzes email headers to determine the authentication status of incoming messages using three primary protocols: SPF, DKIM, and DMARC.

## Authentication Protocols

### SPF (Sender Policy Framework)
SPF is an email authentication method that detects forging sender addresses during email delivery.

**How it works:**
1. The receiving server extracts the sender's domain from the email
2. It queries the domain's DNS records for an SPF record
3. The SPF record lists authorized IP addresses/servers for that domain
4. The receiving server checks if the email came from an authorized server

**SPF Results:**
- `pass`: Email came from authorized server
- `fail`: Email came from unauthorized server  
- `softfail`: Email came from server not explicitly authorized
- `neutral`: Domain owner has not specified policy
- `none`: No SPF record found
- `temperror`: Temporary error during lookup
- `permerror`: Permanent error in SPF record

### DKIM (DomainKeys Identified Mail)
DKIM provides email authentication through digital signatures.

**How it works:**
1. Sending server adds a digital signature to email headers
2. The signature is created using a private key
3. Receiving server retrieves the public key from DNS
4. Server verifies the signature matches the email content

**DKIM Results:**
- `pass`: Digital signature is valid
- `fail`: Digital signature verification failed
- `neutral`: Signature present but not verified
- `none`: No DKIM signature found
- `temperror`: Temporary error during verification
- `permerror`: Permanent error in signature

### DMARC (Domain-based Message Authentication)
DMARC builds on SPF and DKIM to provide policy-based authentication.

**How it works:**
1. Domain owner publishes a DMARC policy in DNS
2. Policy specifies what to do with emails that fail SPF/DKIM
3. Receiving server checks SPF and DKIM alignment
4. Server applies the policy (none, quarantine, reject)

**DMARC Results:**
- `pass`: Email passes DMARC policy
- `fail`: Email fails DMARC policy
- `quarantine`: Email should be quarantined
- `reject`: Email should be rejected
- `none`: No DMARC policy found

## Header Analysis

### Primary Headers Analyzed
- `Authentication-Results`: Contains consolidated auth results
- `Received-SPF`: SPF check results
- `DKIM-Signature`: DKIM signature information
- `ARC-Authentication-Results`: Authentication Results Chain

### Example Headers

**Authentication-Results Header:**
```
Authentication-Results: mx.google.com;
    spf=pass (google.com: domain of sender@example.com designates 209.85.128.180 as permitted sender) smtp.mailfrom=example.com;
    dkim=pass (test mode) header.i=@example.com;
    dmarc=pass (p=QUARANTINE sp=QUARANTINE dis=NONE) header.from=example.com
```

**Received-SPF Header:**
```
Received-SPF: pass (google.com: domain of sender@example.com designates 209.85.128.180 as permitted sender) client-ip=209.85.128.180;
```

**DKIM-Signature Header:**
```
DKIM-Signature: v=1; a=rsa-sha256; c=relaxed/relaxed;
    d=example.com; s=20161025;
    h=from:to:subject:date;
    bh=base64hash;
    b=signature
```

## Security Scoring Algorithm

The add-in calculates a security score (0-100) based on authentication results:

### Scoring Breakdown
- **SPF (30 points maximum)**
  - Pass: 30 points
  - Soft Fail: 15 points  
  - Neutral: 10 points
  - Fail/None: 0 points

- **DKIM (35 points maximum)**
  - Pass: 35 points
  - Neutral: 10 points
  - Fail/None: 0 points

- **DMARC (35 points maximum)**
  - Pass: 35 points
  - Quarantine: 15 points
  - None: 5 points
  - Fail/Reject: 0 points

### Score Interpretation
- **90-100**: Excellent security - Strong authentication
- **70-89**: Good security - Minor issues
- **50-69**: Moderate security - Some failures
- **0-49**: Poor security - Potential spoofing risk

## Implementation Details

### Cross-Platform Compatibility
The add-in is designed to work across:
- **Outlook Web App**: Full functionality
- **Outlook Desktop (Windows/Mac)**: Full functionality
- **New Outlook**: Full functionality  
- **Outlook Mobile**: Limited to available APIs

### API Usage
- Uses Office.js Mailbox API 1.8+ features when available
- Graceful fallback for older API versions
- Mock data for demonstration when headers unavailable

### Browser Support
- Internet Explorer 11 (for Outlook 2016 Windows)
- Modern browsers (Chrome, Firefox, Safari, Edge)
- Mobile WebView containers

## Limitations

### Technical Limitations
1. **Header Access**: Some email providers may filter authentication headers
2. **API Availability**: Older Outlook versions have limited header access
3. **Network Restrictions**: Corporate firewalls may affect DNS lookups

### Authentication Limitations
1. **Forwarded Emails**: May break SPF alignment
2. **Mailing Lists**: Can affect DKIM signatures
3. **Email Gateways**: May add/modify headers

## Security Considerations

### Data Privacy
- No email content is stored or transmitted
- Only email headers are analyzed locally
- No external API calls for authentication data

### Threat Detection
- Identifies potential email spoofing attempts
- Highlights suspicious authentication patterns
- Provides actionable security recommendations

## Troubleshooting

### Common Issues
1. **"Unable to retrieve headers"**: Limited API access in older Outlook
2. **"Mock data" warning**: Headers not accessible, using demo data
3. **Inconsistent results**: Different email providers use different headers

### Debug Information
Enable debug mode by adding `?debug=true` to taskpane URL to see:
- Raw header values
- Parsing results
- API availability status

## Future Enhancements

### Planned Features
- **ARC (Authenticated Received Chain)** analysis
- **BIMI (Brand Indicators for Message Identification)** support
- **Detailed policy examination**
- **Historical tracking** of sender reputation
- **Custom scoring algorithms**
- **Bulk email analysis**
