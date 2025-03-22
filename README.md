# PowerPoint Automation with Excel VBA

## Overview
This VBA project automates the generation of PowerPoint presentations using data from an Excel sheet. Each company listed in the dataset receives a customized PowerPoint report, with placeholders replaced by actual data. The second slide (used as a template) is removed before saving each company's final presentation.

## Features
✅ Generates PowerPoint presentations for each company in the dataset  
✅ Populates slides with relevant financial data  
✅ Automatically deletes the placeholder slide (Slide 2) before saving  
✅ Uses a pre-defined PowerPoint template for consistency  
✅ Runs entirely within Excel using VBA  

## Installation
### Prerequisites
- Microsoft Excel (with Macro-enabled Workbook support `.xlsm`)
- Microsoft PowerPoint
- Basic knowledge of VBA (optional, but useful for modifications)

### Setup Steps
1. **Clone or Download the Repository**
   ```sh
   git clone https://github.com/yourusername/powerpoint-automation-vba.git
   cd powerpoint-automation-vba
   ```
2. **Prepare Your Files**
   - Open `Sample_Data.xlsm` in Excel.
   - Ensure your data is formatted correctly.
   - Place `Sample_Presentation.pptx` in the same directory as your Excel file.
3. **Enable Macros in Excel**
   - Go to `File > Options > Trust Center > Trust Center Settings > Macro Settings`
   - Enable "Trust access to VBA project object model" (Optional but recommended for full functionality)

## Usage
1. **Open `Sample_Data.xlsm`** in Excel.
2. **Run the VBA Macro:**
   - Open the `Developer` tab in Excel.
   - Click `Visual Basic` to open the VBA editor.
   - Locate and run `GenerateCompanyPresentations` macro.
3. **Check Output:**
   - The script generates separate PowerPoint files in the same directory.
   - Each file follows the naming convention `{CompanyName}_Report.pptx`.
   - Slide 2 (placeholder) is removed before saving.

## VBA Code Highlights
```vba
' Delete Slide 2 for every company's presentation before saving
If pptPres.Slides.Count >= 2 Then
    pptPres.Slides(2).Delete
End If
pptPres.Save
pptPres.Close
```

## File Structure
```
📂 powerpoint-automation-vba
├── 📄 Sample_Data.xlsm        # Excel file containing company data
├── 📄 Sample_Presentation.pptx # PowerPoint template
├── 📄 README.md               # Project documentation
└── 📄 Export_Data_to_Ppt_Code.bas            # VBA script file (optional export)
```

## Troubleshooting
- **Issue: PowerPoint does not open**
  - Ensure PowerPoint is installed and not running in background processes.
- **Issue: Placeholders not replaced**
  - Check that placeholder tags (e.g., `[Company Name]`) match exactly in both Excel and PowerPoint.
- **Issue: Slide 2 is not deleted**
  - Ensure your template has at least 2 slides.

## Contributing
Feel free to fork this project, submit issues, or make pull requests! 🚀

## License
This project is open-source and available under the MIT License.

