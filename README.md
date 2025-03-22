# Office Macros

- Repository contains useful VBA macros for Microsoft Word documents.
- Macros simplify tasks in everyday document processing.
- They have been tested and used successfully for various personal use cases.
- Macros are valid as per the date of the most recent commit.

## Table of Contents
- [Word Document Macros](#word-document-macros)
  - [SetFixedTableWidth Macro](#setfixedtablewidth-macro)
  - [RemoveExcessLineSpacing Macro](#removeexcesslinespacing-macro)
- [How to Implement](#how-to-implement)
- [Notes](#notes)

## Word Document Macros

### SetFixedTableWidth Macro
- Sets a fixed table width of 15.98 centimetres for every table in the active document.
- Converts centimetres to points using Word's built-in conversion function.
- Disables the AutoFit feature to prevent automatic resizing.
- Ensures a consistent, uniform table appearance throughout the document.

### RemoveExcessLineSpacing Macro
- Reduces any group of blank lines (more than two in a row) to just one.
- Maintains existing spacing if there are two or fewer consecutive blank paragraphs.
- Helps clean up document formatting and maintain consistent spacing.

## How to Implement

1. **Open Microsoft Word**

2. **Access the VBA Editor:**
   - Press **Option + F11** or go to **Tools > Macro > Visual Basic Editor**.

3. **Insert a New Module:**
   - In the VBA Editor, right-click on your document or template in the left-hand panel.
   - Select **Insert > Module**.

4. **Copy and Paste the Macros:**
   - Copy the desired macro code from this README and paste it into the module window.

5. **Run the Macro:**
   - Close the VBA Editor.
   - In Word, go to **Tools > Macro > Macros**, select the macro you want to run, and click **Run**.

## Notes

- Intended for personal use; modifications may be made to suit specific requirements.
- Save a backup of your document before running any macros, especially if not familiar with VBA.
- For further customisation, refer to [Microsoft’s official VBA documentation](https://learn.microsoft.com/en-us/office/vba/library-reference/concepts/getting-started-with-vba-in-office?).
