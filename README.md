# Purchase Order Generator

A Python-based desktop application that automates the creation of purchase orders by extracting information from supplier quotation PDFs using Google's Gemini AI and generating formatted Excel files.

## 🚀 Features

- **Universal AI Scanning**: Transform ANY supplier quotation PDF into a professional purchase order using advanced AI - works with quotations from any company, any format
- **Modern Grouped Interface**: Clean, organized layout with labeled sections for Project Information, Personnel Information, and Attachments
- **Real-time API Status**: Visual indicator showing API connection status (green/red dot) in the header
- **Smart Form Validation**: Automatic PO number format validation with helpful error messages
- **Interactive Date Picker**: Calendar widget with "Today" button for quick date selection
- **International Phone Support**: Country code dropdown with phone number input
- **API Key Management**: Built-in settings dialog with connection testing and clipboard paste functionality
- **Progress Tracking**: Real-time generation progress with elapsed time display
- **Auto-Save Preferences**: Optional form data persistence with checkbox control
- **Smart File Handling**: Browse dialog for PDF selection with intelligent file naming
- **Custom Success Dialog**: Post-generation options to save, open, or continue working

## 📋 Prerequisites

- Python 3.8 or higher
- Google Gemini API key
- Windows/macOS/Linux

## 🛠️ Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/ikmalrasli/po-generator.git
   cd po-generator
   ```

2. **Install required packages**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up Google API Key**
   - Launch the application with `python main.py`
   - Click the "⚙ Settings" button in the top-right corner
   - Enter your Google Gemini API key or click "Get one here" to obtain one
   - Use "Test Connection" to verify your key works
   - Click "OK" to save

4. **Set up Excel template**
   - ⚠️ **Important**: The provided `po_template_sample.xlsx` in the `templates/` folder is a **SAMPLE TEMPLATE**
   - **Customize the sample template**:
     - Update company name, address, and contact information
     - Add your company logo and branding
     - Modify header/footer to match your company's PO format
     - Adjust styling, colors, and fonts as needed
   - **Rename and activate**: After customization, rename `po_template_sample.xlsx` to `po_template.xlsx`
   - Ensure your template maintains the same cell structure and formatting used by the application

## 🎯 Usage

### Running the Application

```bash
python main.py
```

### Complete User Workflow

1. **API Setup (First Time Only)**:
   - Look at the API status indicator in the header (red dot = "API Key Missing")
   - Click "⚙ Settings" button to open the API key dialog
   - Paste your Google Gemini API key or click the link to get one
   - Test the connection to verify it works
   - Save the key - the status indicator will turn green ("API Active")

2. **Fill in Project Information**:
   - **PO Number**: Enter in format P-######-###M (e.g., P-250719-001M)
   - **Project Name**: Enter the project name
   - **PO Issue Date**: Use the date picker or click "Today" for current date

3. **Fill in Personnel Information**:
   - **Purchaser Name**: Enter the purchaser's full name
   - **Purchaser Telephone**: Select country code (+60 default) and enter phone number
   - **Manager Name**: Enter the approving manager's name

4. **Attach Quotation PDF**:
   - Click "Browse" to select ANY supplier's quotation PDF (from any company, any format)
   - The AI will automatically scan and extract all relevant information
   - The file path will appear in the text field
   - The "Generate Purchase Order" button will become enabled when both API key and PDF are set

5. **Configure Auto-Save (Optional)**:
   - Check "Remember details for next time (auto-save)" to save your form data
   - When enabled, your information will be automatically restored on next launch
   - Changes are saved in real-time as you type

6. **Generate Purchase Order**:
   - Click "Generate Purchase Order" (button shows "Generating..." during processing)
   - Window title displays elapsed time during generation
   - Progress is shown with real-time timer

7. **Handle Generated File**:
   - **Success Dialog** appears with generation time
   - Choose "Save As..." to save the file with custom name/location
   - Click "Open File" to view the generated Excel immediately
   - Or click "OK" to keep the file in temp folder

8. **Additional Actions**:
   - **Clear Form**: Resets all fields while preserving auto-save setting
   - **Save As...**: Save the generated PO to your preferred location
   - **Open File**: Open the most recently saved file

### Interface Elements

- **Header**: Shows app title and real-time API status with settings access
- **Grouped Sections**: Organized input fields in logical categories
- **Smart Buttons**: Context-aware button states (disabled/enabled based on requirements)
- **Validation**: Real-time input validation with helpful error messages
- **Progress Feedback**: Visual feedback during processing with timing information

## 📁 Project Structure

```
purchase-order-generator/
├── main.py                 # Application entry point
├── config/
│   ├── settings.py         # Configuration and paths
│   └── user_settings.json  # Auto-saved user preferences
├── core/
│   ├── excel_generator.py  # Excel file generation logic
│   ├── pdf_processor.py    # AI-powered PDF processing
│   └── utils.py           # Helper functions and validations
├── gui/
│   ├── app.py             # Main GUI application with dialogs
│   └── components.py      # Reusable UI components
├── templates/
│   ├── po_template_sample.xlsx  # ⚠️ SAMPLE TEMPLATE - CUSTOMIZE AND RENAME
│   └── po_template.xlsx  # Your active template (rename from sample)
├── temp/                  # Temporary files (auto-created)
├── jsons/                 # Extracted JSON data (auto-created)
└── requirements.txt       # Python dependencies
```

## 🏗️ Customization

### Template Customization

⚠️ **The provided template is a sample** - you MUST customize it with your company's branding:

1. **Customize the Sample Template**:
   - Open `templates/po_template_sample.xlsx`
   - **Company Branding**: Add your company logo, update name, address, contact info
   - **Visual Design**: Adjust colors, fonts, styling to match your company's PO format
   - **Header/Footer**: Modify to include your company's standard PO layout
   - **Save**: After customization, **rename** `po_template_sample.xlsx` to `po_template.xlsx`

2. **Key Template Requirements** (maintain these cell references):
   - Header fields in specific cells (B9, B10, B11, etc.)
   - Item table starting at row 31
   - Total calculation formulas in column I
   - Signature fields for approvals

### Field Mapping

The application maps data to these Excel cells (adjust if your template differs):

- **Company Info**: B9 (name), B10-B11 (address)
- **Contact Details**: C13-C16 (PIC information)
- **PO Header**: E8 (PO#), H8 (date), H9-H10 (project)
- **Items Table**: Rows 31+ (item details)
- **Totals**: Automatic calculation in column I

## 🔧 Building Executable

To create a standalone executable:

```bash
# Using the provided build script
python build.py

# Create installer (after executable is built)
python build_installer.py
```

The executable will be created in the `dist/` folder, and the installer in the root directory.

## 🐛 Troubleshooting

### Common Issues

1. **API Key Issues**:
   - Check the API status indicator in the header
   - Use "Test Connection" in the API settings dialog
   - Ensure internet connection for API calls
   - Verify key validity at https://aistudio.google.com/api-keys

2. **Form Validation Errors**:
   - PO Number must follow format: P-######-###M
   - All required fields must be filled before generation
   - Quotation PDF must exist and be accessible

3. **PDF Processing Fails**:
   - Ensure PDF is not password protected
   - Check that PDF contains readable text (not scanned images)
   - Verify file size is reasonable (< 10MB recommended)

4. **Excel Generation Errors**:
   - Verify your custom template maintains required cell structure
   - Check that Excel is not open when generating files
   - Ensure write permissions in the target directory

5. **Build Issues**:
   - Run `python build.py` first before `python build_installer.py`
   - Ensure NSIS is installed for installer creation
   - Check that all dependencies are in requirements.txt

## 📝 License

This project is licensed under the MIT License - see the LICENSE file for details.

## 🤝 Contributing

1. Fork the repository
2. Create a feature branch (`git checkout -b feature/AmazingFeature`)
3. Commit your changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## 📞 Support

If you encounter any issues or have questions:

1. Check the troubleshooting section above
2. Search existing GitHub issues
3. Create a new issue with detailed information

## ⚠️ Important Notes

- **Template Requirement**: The success of this application depends on using your company's proper Excel template. The provided sample is for reference only.
- **API Costs**: Google Gemini API usage may incur costs. Monitor your API usage.
- **Data Privacy**: PDFs are sent to Google's servers for processing. Ensure compliance with your organization's data policies.
- **Auto-Save**: User preferences are stored locally and persist between application launches.
- **Backup**: Always keep backups of your custom template and generated files.
