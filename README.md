# Tag Description Updater

A user-friendly tool to batch update tag descriptions in Ignition JSON exports using descriptions from an Excel spreadsheet.

## Overview

This application streamlines the process of updating tag descriptions across multiple data blocks in Ignition. Instead of manually editing each tag in Ignition, you can:

1. **Prepare descriptions** in an Excel file with a structured format
2. **Export tags** from Ignition as JSON
3. **Run the updater** to automatically match and update descriptions
4. **Import** the updated JSON back into Ignition

## Features

- 📋 Simple graphical interface (GUI) - no command line needed
- 📊 Supports multiple Excel sheets (processes all except SCADA_SIGNAL)
- 🏷️ Automatically writes Excel sheet names as `Group` parameter to UDT instances
- 🔍 Optional filter to process only alarm tags (`Is Alarm = True`)
- 📁 Real-time logging of all updates
- 🚀 Automatic output folder opening after completion
- ✅ Handles nested tag structures (UDT instances with Configuration folders)

## Requirements

### Option 1: Using Python

- Python 3.7 or higher
- openpyxl (for Excel file handling)

Install dependencies with:

```bash
pip install -r requirements.txt
```

Then run:

```bash
python TagDescriptionUpdater.py
```

### Option 2: Using Executable (.exe) 

A standalone executable version is available that runs without Python installed. Simply download and run:

```
TagDescriptionUpdater.exe
```

This is the easiest option for end users who don't have Python installed.

## File Formats

### Excel File Format

Create an Excel file with the following columns:

| Column | Header | Required | Example |
|--------|--------|----------|---------|
| A | Data Block | Yes | `Motor_Control` |
| B | Tag Name | Yes | `Motor_Speed` |
| C | Type | No | `Integer` |
| D | Unit | No | `RPM` |
| E | Min Value | No | `0` |
| F | Is Alarm | Optional | `TRUE` or `FALSE` |
| G | Other Info | No | Any additional data |
| H | Description | Yes | `Actual speed of main motor` |

**Notes:**
- First row is treated as header and skipped
- Only tags with a description (Column H) will be processed
- The filter option only processes rows where `Is Alarm = TRUE` (Column F)
- You can have multiple sheets; the tool processes all except `SCADA_SIGNAL`
- Tag Name and Data Block must match exactly with the JSON file

### JSON File Format

The JSON file should be an Ignition export with the following structure:

```json
{
  "name": "COMMON_ALARM",
  "parameters": {
    "DataBlock": {
      "dataType": "String",
      "value": "AC"
    },
    "DriverDeviceConnection": {
      "dataType": "String",
      "value": "[Siemens Enhanced]"
    },
    "Group": {
      "dataType": "String",
      "value": "Group2"
    }
  },
  "tagType": "UdtInstance",
  "tags": [
    {
      "name": "FolderName",
      "tagType": "Folder",
      "tags": [
        {
          "name": "AlarmTag",
          "tagType": "UdtInstance",
          "parameters": {
            "DataBlock": {
              "dataType": "String",
              "value": "Motor_Control"
            },
            "Group": {
              "dataType": "String",
              "value": "SheetName"
            }
          },
          "tags": [
            {
              "name": "Configuration",
              "tagType": "Folder",
              "tags": [
                {
                  "name": "Alarm_Description",
                  "tagType": "AtomicTag",
                  "value": ""
                },
                {
                  "name": "Label",
                  "tagType": "AtomicTag",
                  "value": ""
                }
              ]
            }
          ]
        }
      ]
    }
  ]
}
```

**Note:** The `Group` field in the parameters will be automatically populated with the Excel sheet name during the update process.

## How to Use

### Step 1: Prepare Your Excel File

1. Create an Excel spreadsheet with tag descriptions
2. Ensure columns are arranged as specified above
3. Add descriptions in Column H
4. Save the file (e.g., `tag_descriptions.xlsx`)

### Step 2: Export JSON from Ignition

1. In Ignition Designer, export your tag structure as JSON
2. Save the file (e.g., `tags_export.json`)

### Step 3: Run the Updater

1. Run the application:
   ```bash
   python TagDescriptionUpdater.py
   ```

2. **Select Excel File** - Click the blue button and choose your Excel file

3. **(Optional) Toggle Filter** - Click the green filter button to enable/disable filtering for alarm tags only

4. **Select JSON File** - Click the blue button and choose your Ignition JSON export

5. **Click "Update Descriptions"** - The tool will:
   - Load descriptions from Excel
   - Match tags by Data Block and Tag Name
   - Update descriptions in JSON
   - Ask where to save the updated file

### Step 4: Import Back to Ignition

1. Open your Ignition Designer
2. Import the updated JSON file
3. Your tags now have updated descriptions!

## Matching Logic

The tool matches tags using a two-part key: `DataBlock|TagName`

- **For AtomicTag**: Uses the parent folder name as Data Block
- **For UdtInstance**: Uses the `DataBlock` parameter if available, otherwise uses parent folder name

### Group Field Assignment

The Excel sheet name is automatically written to the `Group` parameter in UDT instances:
- Each Excel sheet (except `SCADA_SIGNAL`) is processed separately
- The sheet name becomes the value of the `Group` parameter in the updated JSON
- This allows you to track which source sheet each tag came from in the output

## Example Workflow

1. Export your Ignition tags → `tags_export.json`
2. Create Excel with descriptions → `descriptions.xlsx`
3. Run: `python TagDescriptionUpdater.py`
4. Select both files
5. Click "Update Descriptions"
6. Save output as → `tags_updated.json`
7. Import `tags_updated.json` back into Ignition

## Troubleshooting

### No tags updated
- ✓ Check that Data Block and Tag Name match exactly (case-sensitive)
- ✓ Verify descriptions are in Column H
- ✓ Ensure JSON has correct structure with matching tag names

### Some tags not updating
- ✓ If filtering is enabled, ensure tags have `Is Alarm = TRUE` in Excel
- ✓ Check that the JSON structure has Configuration/Alarm_Description folders

### Excel file won't load
- ✓ Ensure file is `.xlsx` format (not `.xls` or `.csv`)
- ✓ Check that file is not corrupted
- ✓ Close the file if it's open in Excel
