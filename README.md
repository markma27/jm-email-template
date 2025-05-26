# JM Email Template Launcher

A Chrome extension that allows users to choose email templates from a SharePoint Online list and launch them in Outlook Web.

## Features

- Fetch email templates from SharePoint Online list
- Filter templates by category
- Preview template content before sending
- Launch emails directly in Outlook Web
- Fallback to `mailto:` if Outlook Web is not available
- Modern UI with Aptos font

## Installation

1. Clone or download this repository
2. Install dependencies for icon generation (optional, only if you want to regenerate icons):
   ```bash
   npm install canvas
   node generate_icons.js
   ```
3. Open Chrome and navigate to `chrome://extensions/`
4. Enable "Developer mode" in the top right
5. Click "Load unpacked" and select this directory

## Usage

1. Click the extension icon in Chrome's toolbar
2. Sign in to SharePoint if prompted
3. Select a template from the dropdown (optionally filter by category)
4. Preview the template content
5. Click "Launch Outlook Email" to open a new email with the template

## SharePoint List Requirements

The extension expects a SharePoint list named "JM Email Template Library" with the following columns:
- Title (text)
- Email Category (choice/text)
- Email Body (HTML-formatted text)

## Configuration

The SharePoint site URL is configured in `popup.js`. Update `SHAREPOINT_CONFIG` if needed:

```javascript
const SHAREPOINT_CONFIG = {
    siteUrl: 'https://jaquillardminns.sharepoint.com/sites/JaquillardMinns411',
    listTitle: 'JM Email Template Library'
};
```

## Security

- Uses SharePoint authentication (no credentials stored)
- HTML content is sanitized before display
- CSP prevents inline scripts and external resources
- All communication uses HTTPS

## Development

The extension is built with vanilla JavaScript and uses:
- SharePoint REST API for data
- Modern CSS with Flexbox
- Chrome Extension Manifest V3

## Files

- `manifest.json` - Extension configuration
- `popup.html` - Extension popup UI
- `popup.css` - Styles for the popup
- `popup.js` - Main extension logic
- `icon16.png` - 16x16 extension icon
- `icon48.png` - 48x48 extension icon
- `icon128.png` - 128x128 extension icon
- `generate_icons.js` - Icon generation script

## Troubleshooting

1. **Templates not loading**
   - Ensure you're signed into SharePoint
   - Check SharePoint list permissions
   - Verify list name and structure

2. **Email not launching**
   - Check if popups are blocked
   - Ensure you're signed into Outlook Web
   - Try the `mailto:` fallback

3. **Preview not showing**
   - Check if template has HTML content
   - Verify HTML format in SharePoint

## License

This project is proprietary and confidential.

## Support

For support, please contact your system administrator or IT support team. 