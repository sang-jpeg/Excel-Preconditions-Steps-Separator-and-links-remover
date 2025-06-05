# Excel Test Case Processor - Documentation

## Overview
This is a comprehensive Excel processing tool designed to clean, format, and restructure test cases written in Gherkin format. The tool automatically separates combined content, formats test steps, and produces clean, well-organized test documentation.

## Core Features

### 1. Column Splitting and Separation
- **Primary Function**: Splits a single column (default: Column C) containing combined Preconditions, Steps, and Expected Results into three separate columns
- **Smart Detection**: Uses customizable keywords to identify where each section begins
- **Default Keywords**:
  - Preconditions: "Precondition, Pre-condition, Prerequisites, Pre-requisites"
  - Steps: "Steps, Test Steps, Procedure, Actions, Test Procedure"
  - Expected Results: "Expected Result, Expected Results, Expected Outcome, Expected"

### 2. Automatic Formatting Fixes
- **Joined Text Detection**: Automatically detects and separates text that was incorrectly joined together
  - Example: `"User is logged inUser has access"` → `"User is logged in\nUser has access"`
- **Pattern Recognition**: Identifies common patterns like:
  - User statements
  - Device/System/App conditions
  - Numbered items (1., 1.1, 2., etc.)
  - Sentence boundaries (lowercase→uppercase transitions)

### 3. References Section Handling
- **Smart Removal**: Automatically removes "References (for internal use only)" sections
- **Conditional Removal**: Removes References sections that contain links (like Confluence URLs)
- **Preservation**: Keeps References sections that don't contain links

### 4. Gherkin Formatting
- **Keyword Capitalization**: Automatically capitalizes Given/When/Then/And/But keywords
- **Line Breaking**: Each Gherkin statement starts on a new line
- **Number Detection**: Formats numbered lists (1., 1.1, a., i., etc.) with proper line breaks

### 5. Data Appending
- **Users Applied Feature**: Appends data from another column (default: E) to Preconditions
- **Format**: Added as "Users applied: [data]" separated by three underscores (___)
- **Column Removal**: Automatically removes the source column after appending

### 6. Link Cleaning
- **Image Link Processing**: Cleans QMetry image links
  - From: `!URL|width=200,height=183!`
  - To: `URL`
- **URL Removal**: Removes non-QMetry URLs (like Atlassian/JIRA links)
- **Smart Detection**: Only processes cells that contain links, preserving formatting in others

### 7. Interactive Preview
- **Expandable Rows**: Click indicators to expand/collapse long content
- **Visual Feedback**: Hover effects and clear indicators for rows with more content
- **Bulk Controls**: "Expand All" and "Collapse All" buttons
- **Responsive Design**: Works on desktop, tablet, and mobile devices

## User Interface

### File Upload Section
- **Choose File Button**: Disabled after file selection to prevent multiple uploads
- **Clear File (×)**: Red button to remove selected file and reset
- **Drag & Drop**: Supports dragging Excel files directly onto the upload area

### Settings Panel
1. **Keyword Configuration**: Customizable keywords for each section
2. **Column Selection**: Choose which column to split (default: C)
3. **Numbered Steps Detection**: Toggle for detecting numbered items
4. **Append Data**: Optional feature to append data from another column
5. **Link Cleaning**: Toggle for cleaning image links with column selector
6. **Gherkin Formatting**: Toggle for formatting Gherkin keywords

### Preview and Download
- **Preview Changes Button**: Processes file and shows results
- **Interactive Table**: Expandable cells for reviewing changes
- **Process Another File**: Clears everything and starts fresh
- **Download Processed File**: Saves the cleaned Excel file

## Processing Order

1. **Fix Formatting Issues** - Detects and fixes joined text
2. **Separate Content** - Splits into Preconditions/Steps/Expected Results
3. **Remove References** - Cleans internal reference sections
4. **Append User Data** - Adds "Users applied" information
5. **Format Gherkin** - Capitalizes keywords and adds line breaks
6. **Clean Links** - Processes image links and removes unwanted URLs

## Technical Details

### Supported File Types
- `.xlsx` (Excel 2007+)
- `.xls` (Legacy Excel)

### Browser Compatibility
- Chrome (recommended)
- Firefox
- Safari
- Edge
- Mobile browsers

### Libraries Used
- **XLSX.js**: For reading and writing Excel files
- **Pure JavaScript**: No server-side processing required
- **HTML5 APIs**: File reading and drag-and-drop support

## Common Use Cases

### Example 1: Basic Test Case Cleaning
**Input**: Single column with mixed content
```
Preconditions: User is logged inUser has premium account
Steps: 1.Click profile2.Select settings
Expected Result: Settings page loads
```

**Output**: Three separate columns
- Preconditions: 
  ```
  User is logged in
  User has premium account
  ```
- Steps:
  ```
  1. Click profile
  2. Select settings
  ```
- Expected Results:
  ```
  Settings page loads
  ```

### Example 2: Gherkin Formatting
**Input**: 
```
given user is on homepage when they click login then login form appears
```

**Output**:
```
Given user is on homepage
When they click login
Then login form appears
```

## Best Practices

1. **Backup Your Data**: Always keep a copy of your original Excel file
2. **Review Preview**: Check the preview before downloading to ensure correct processing
3. **Customize Keywords**: Adjust keywords to match your team's terminology
4. **Test Small First**: Try with a few rows before processing large files

## Troubleshooting

### Issue: Choose File button stays disabled
- **Solution**: Reload the page


### Issue: Formatting not detected
- **Solution**: Add specific keywords to the detection patterns
- **Check**: Ensure content follows expected patterns

### Issue: Preview not showing
- **Verify**: File format is supported (.xlsx or .xls)

## Future Enhancements
- Multi-column link cleaning
- Custom formatting rules
- Batch file processing
- Export to different formats
- Integration with test management tools

---

This tool streamlines test case management by automating repetitive formatting tasks, ensuring consistency across test documentation, and saving significant manual effort.
