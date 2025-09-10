# TensAI Office Add-in

This is a unified Office Add-in that works across Word, Excel, PowerPoint, and Outlook, integrating with TensAI services at www.dev2.tens-ai.com.

## Features

- **Multi-Host Support**: Works seamlessly across Word, Excel, PowerPoint, and Outlook
- **Excel Integration**: Advanced data manipulation and analysis in Excel workbooks
- **Word Integration**: Document processing and content analysis in Word
- **PowerPoint Integration**: Presentation enhancement and content generation
- **Outlook Integration**: Email processing and productivity features
- **Data Import/Export**: Import and export data between Office applications and TensAI
- **AI Analysis**: Run AI-powered analysis on your Office documents
- **Real-time Collaboration**: Connect with TensAI dashboard for real-time insights

## Project Structure

```
TensAI Office Addin/
├── manifest.xml          # Office Add-in manifest file
├── index.html           # Main add-in page
├── taskpane.html        # Task pane interface
├── commands.html        # Add-in commands
├── package.json         # Node.js dependencies and scripts
├── assets/              # Icons and images
│   ├── icon-16.png
│   ├── icon-32.png
│   ├── icon-64.png
│   ├── icon-80.png
│   ├── icon-128.png
│   └── logo-filled.png
└── README.md           # This file
```

## Prerequisites

- Node.js (v14 or higher)
- npm
- Office 365 or Office 2016/2019/2021
- Access to www.dev2.tens-ai.com

## Installation

1. **Install dependencies:**
   ```bash
   npm install
   ```

2. **Validate the manifest:**
   ```bash
   npm run validate
   ```

## Development

### Local Testing

To test the add-in locally, you'll need to:

1. **Host the files on www.dev2.tens-ai.com:**
   - Upload all HTML files to your web server
   - Upload the assets folder with all icons
   - Ensure HTTPS is enabled

2. **Sideload the add-in:**
   ```bash
   npm run sideload
   ```

### Available Scripts

- `npm run validate` - Validate the manifest file
- `npm run start` - Start local development server
- `npm run stop` - Stop local development server
- `npm run sideload` - Sideload the add-in to Office

## Deployment

### Option 1: Centralized Deployment (Recommended for Organizations)

1. Upload the manifest.xml and all referenced files to your web server
2. Use Microsoft 365 Admin Center to deploy the add-in to your organization
3. Users will see the add-in automatically in their Office applications

### Option 2: SharePoint App Catalog

1. Upload the add-in files to a SharePoint App Catalog
2. Deploy through SharePoint for on-premises environments

### Option 3: AppSource (Public Distribution)

1. Package the add-in with all required files
2. Submit to AppSource for public distribution
3. Follow Microsoft's validation and approval process

## Configuration

### Required Files on Web Server

The following files must be hosted on www.dev2.tens-ai.com:

- `index.html` - Main add-in page
- `taskpane.html` - Task pane interface
- `commands.html` - Add-in commands
- `assets/icon-16.png` - 16x16 icon
- `assets/icon-32.png` - 32x32 icon
- `assets/icon-80.png` - 80x80 icon

### HTTPS Requirement

All URLs in the manifest must use HTTPS. Ensure your web server has a valid SSL certificate.

## Usage

### Installation
1. **Install the add-in** in any Office application (Word, Excel, PowerPoint, or Outlook)
2. **Open the task pane** by clicking the "Open TensAI" button in the Home ribbon (or appropriate tab)

### Features by Application

#### Excel
- **Load Sample Data**: Populate worksheets with sample data
- **Get Worksheet Info**: View current worksheet information
- **Data Import/Export**: Import/export data between Excel and TensAI
- **AI Analysis**: Run analysis on spreadsheet data

#### Word
- **Document Processing**: Analyze and enhance Word documents
- **Content Generation**: Generate content using TensAI
- **Data Import/Export**: Import/export document data

#### PowerPoint
- **Presentation Enhancement**: Improve presentations with AI
- **Content Generation**: Generate slides and content
- **Data Import/Export**: Import/export presentation data

#### Outlook
- **Email Processing**: Analyze and enhance emails
- **Productivity Features**: Improve email workflow
- **Data Import/Export**: Import/export email data

### Common Features
- **Open TensAI Dashboard**: Access the web interface
- **AI Analysis**: Run AI-powered analysis on your content
- **Real-time Integration**: Connect with TensAI services

## API Integration

The add-in is configured to work with TensAI services at:
- **Base URL**: https://www.dev2.tens-ai.com
- **API Endpoint**: https://www.dev2.tens-ai.com/api

## Troubleshooting

### Common Issues

1. **Add-in not loading**: Check that all URLs in manifest.xml are accessible
2. **Icons not displaying**: Ensure icon files are uploaded to the correct paths
3. **HTTPS errors**: Verify SSL certificate is valid and properly configured

### Validation Errors

If you encounter validation errors:

1. **Missing Support URL**: Add a valid support URL in manifest.xml
2. **Invalid AppDomain**: Ensure domains are valid and accessible
3. **Unreachable URLs**: Verify all referenced files are hosted and accessible

## Support

For support and questions:
- **Website**: https://www.dev2.tens-ai.com
- **Support**: https://www.dev2.tens-ai.com/support

## License

MIT License - see LICENSE file for details.

## Contributing

1. Fork the repository
2. Create a feature branch
3. Make your changes
4. Test thoroughly
5. Submit a pull request

## Version History

- **v1.0.0** - Initial release with basic TensAI integration
# TensAI-addin
