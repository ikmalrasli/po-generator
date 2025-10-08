# Purchase Order Generator

A Python-based desktop application that automates the creation of purchase orders by extracting information from supplier quotation PDFs using Google's Gemini AI and generating formatted Excel files.

## 🚀 Features

- **AI-Powered PDF Processing**: Automatically extracts company details, items, and terms from quotation PDFs using Google Gemini AI
- **Excel Template Generation**: Creates professionally formatted purchase orders based on customizable templates
- **User-Friendly GUI**: Simple interface built with Tkinter for easy data entry and file management
- **Auto-Save Preferences**: Remember frequently used details with the auto-save feature
- **Export Flexibility**: Save generated POs to any location with custom filenames

## 📋 Prerequisites

- Python 3.8 or higher
- Google Gemini API key
- Windows/macOS/Linux

## 🛠️ Installation

1. **Clone the repository**
   ```bash
   git clone https://github.com/yourusername/purchase-order-generator.git
   cd purchase-order-generator
   ```

2. **Install required packages**
   ```bash
   pip install -r requirements.txt
   ```

3. **Set up environment variables**
   - Create a `.env` file in the root directory
   - Add your Google Gemini API key:
     ```
     GOOGLE_API_KEY=your_api_key_here
     ```

4. **Set up Excel template**
   - ⚠️ **Important**: The provided `po_template.xlsx` in the `data/templates/` folder is a **SAMPLE TEMPLATE**
   - Replace it with your own company's purchase order template
   - Ensure your template maintains the same cell structure and formatting used by the application

## 🎯 Usage

### Running the Application

```bash
python main.py
```

### Basic Workflow

1. **Fill in PO Details**:
   - PO Number (format: P-######-###M)
   - Project Name
   - Purchaser Information
   - Manager Details

2. **Select Quotation PDF**:
   - Browse and select the supplier's quotation PDF
   - The AI will automatically extract relevant information

3. **Generate & Save**:
   - Click "Generate Purchase Order"
   - Use "Save As..." to choose the save location and filename

### Auto-Save Feature

- Check "Remember details for next time" to automatically save your inputs
- When enabled, your details will be pre-filled on next startup
- Changes are saved automatically as you type

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
│   ├── app.py             # Main GUI application
│   └── components.py      # Reusable UI components
├── data/
│   └── templates/
│       └── po_template.xlsx  # ⚠️ SAMPLE TEMPLATE - REPLACE WITH YOUR OWN
├── temp/                  # Temporary files (auto-created)
├── jsons/                 # Extracted JSON data (auto-created)
└── requirements.txt       # Python dependencies
```

## 🏗️ Customization

### Template Customization

⚠️ **The provided template is a sample** - you MUST replace it with your company's actual purchase order template:

1. **Replace the Sample Template**:
   - Copy your company's Excel purchase order template to `data/templates/po_template.xlsx`
   - Ensure it maintains the same cell references used in the code

2. **Key Template Requirements**:
   - Header fields in specific cells (B9, B10, B11, etc.)
   - Item table starting at row 31
   - Total calculation formulas
   - Signature fields

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

# Or using PyInstaller directly
pyinstaller build.spec
```

The executable will be created in the `dist/` folder.

## 🐛 Troubleshooting

### Common Issues

1. **API Key Error**:
   - Ensure `.env` file exists with valid `GOOGLE_API_KEY`
   - Check internet connection for API calls

2. **Template Not Found**:
   - Verify `po_template.xlsx` exists in `data/templates/`
   - Remember: Replace the sample template with your own

3. **PDF Processing Fails**:
   - Ensure PDF is not password protected
   - Check that PDF contains readable text (not scanned images)

4. **Excel Generation Errors**:
   - Verify your custom template maintains required cell structure
   - Check that Excel is not open when generating files

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
- **Backup**: Always keep backups of your custom template and generated files.
