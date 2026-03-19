# General Utilities Module

This module contains a comprehensive collection of utility functions designed to simplify common tasks related to data handling, file operations, statistical comparisons, Excel manipulations, and geospatial processing. Below is an overview of the functions available in the module.

## ✅ How to Use

To use this module in your Python project, import the functions you need. Example:

```python
from general_utilities import (
    addPrefixesToColumnNames,
    convertToDatetime,
    compareValueRangesMathematically,
    saveToExcel
)

# Example usage:
dataframe = pandas.read_csv("example.csv")
dataframe = addPrefixesToColumnNames(dataframe, column_names=["name", "age"], prefixes="demo_")
date = convertToDatetime("2024-04-20")
scores = compareValueRangesMathematically([1, 2, 3], [1.1, 2.1, 3.2])
saveToExcel(dataframe, "output.xlsx", "Processed Data")
```

You can also import the entire module if needed:

```python
import general_utilities
```

Or use an alias for convenience:

```python
import general_utilities as GU

dataframe = GU.addPrefixesToColumnNames(dataframe, column_names=["id", "score"], prefixes="meta_")
```

---

## 📊 DataFrame Operations

### General Column Manipulation

* `addEmptyColumnsToDataframe(dataframe, columns, dtype=None, overwrite=False)`: Adds new empty columns with optional data types.
* `addPrefixesToColumnNames(dataframe, column_names=None, prefixes="df")`: Adds specified prefixes to given column names.
* `addSuffixesToColumnNames(dataframe, column_names=None, suffixes="df")`: Adds specified suffixes to given column names.
* `reduceColumns(dataframe, columns_to_keep=None, columns_to_drop=None)`: Reduces the DataFrame to desired columns.
* `reorderColumnsOfDataFrame(dataframe, starter_columns, column_before_starters=None)`: Reorders columns placing certain ones at the start or after a specified column.
* `replaceColumnNameOfDataFrame(dataframe, old_substrings, new_substrings)`: Renames columns by replacing substrings.
* `categorizeColumnsByType(dataframe, keys=[], ignoring_columns=[])`: Categorizes columns into string or numeric types excluding keys and ignored columns.

### Missing Values & Cleaning

* `highlightDataFrameMissingValues(dataframe)`: Highlights missing values in yellow.
* `highlightMissingValue(column)`: Helper function that returns CSS styling for missing values.
* `convertAllStringNumericsToNumeric(dataframe, ignoring_columns=[])`: Converts all string-based numerics to actual numeric values.
* `convertDataFrameStringNumericToNumeric(dataframe, numeric_columns=None, ignoring_columns=None)`: Converts specific columns from string to numeric.
* `convertStringNumericToNumeric(value)`: Converts a single string numeric value to numeric type.
* `removeDuplicateData(dataframe, ignoring_columns=[])`: Removes duplicate rows and duplicate columns.
* `removeEmptyRows(dataframe, columns_to_check)`: Removes rows that have all NaNs in specified columns.

### Column Expansion & Splitting

* `expandColumns(dataframe, source_columns, destination_columns, string_separators, remove_source_columns=False)`: Splits columns into multiple based on delimiters.
* `splitDataFrameHorizontally(dataframe, common_columns, columns_to_split)`: Splits horizontally into multiple DataFrames.
* `splitDataFrameVertically(dataframe, grouping_columns)`: Splits vertically based on groupings.
* `splitDataFrameVerticallyIntoExcelFiles(dataframe, grouping_columns, save_folder, file_name_prefix, data_value_as_file_name)`: Saves vertical splits as separate Excel files.

### Column & Value Helpers

* `replaceSubstringsInDataFrame(dataframe, column_names, old_substrings, new_substrings)`: Replaces multiple substrings with new ones in specified columns.
* `roundCoordinates(dataframe, coordinate_columns, precision_digits)`: Rounds coordinates to specified precision.
* `uniqueValuesCount(values)`: Counts unique non-null values.

---

## 📈 Statistical Utilities

### Metrics

* `calculateR2(observed, predicted)`: Calculates R-squared.
* `calculateRMSD(observed, predicted)`: Calculates Root Mean Square Deviation.
* `calculateDegreeOfAgreement(observed, predicted)`: Computes Willmott’s degree of agreement.

### Comparisons

* `compareValueRangesMathematically(first_list, second_list)`: Compares two sets using R², RMSD, and agreement.
* `compareByBiasCorrection(observed, predicted)`: Applies bias correction to predictions.
* `compareKDE(observed, predicted)`: Compares Kernel Density Estimations.
* `compareRangesDifferenceByQuantiles(observed, predicted, quantiles)`: Quantile comparison.
* `compareRangesDifferenceByKLDivergence(first_list, second_list)`: KL Divergence comparison.
* `compareRangesDifferenceByKSTest(first_list, second_list)`: KS Test comparison.
* `compareRangesDifferenceByMannWhitney(first_list, second_list)`: Mann-Whitney U test comparison.
* `getNormalRangesDifference(first_list, second_list)`: Difference by range width.
* `getVariantRangesDifference(first_list, second_list, acceptable_percent)`: Combines multiple statistical differences.

---

## 🗓️ Date and Time

* `convertToDatetime(value, source_format=None)`: Converts a value to datetime.
* `changeDateTimeFormatInDataFrame(dataframe, column_names, new_formats)`: Changes date format of specified columns.
* `insertDateByTimestampIntoDataFrame(dataframe, timestamp_column, date_column_name)`: Inserts date column from timestamp.
* `insertYearByTimestampIntoDataFrame(dataframe, timestamp_column, year_column_name)`: Extracts year from timestamp.
* `addNewDateColumnByDateRangesToDataFrame(dataframe, column_name, date_ranges, new_date_column_name, new_date_format)`: Creates new date columns based on date range.

---

## 🧽 Geospatial Utilities

* `convertShapeFileDataToDataFrame(file_path, projection_system)`: Converts shapefile to DataFrame.
* `extractCSVDataIntoDataFrame(file_path, file_proj, dest_proj)`: Extracts CSV and transforms projection.
* `getDominantProjectionSystem(shape_files_path)`: Finds most common CRS in shape files.
* `getDominantProjectionSystemOfCSVFiles(csv_files_path)`: Finds common projection in CSV files.
* `addLatAndLongColumnsToDataframe(dataframe, location_column, lat_column, lon_column, remove_location)`: Splits location string into lat/lon columns.

---

## 📂 File and Sheet Handling

### Excel

* `makeExcelFileColumnsWidthAutoFit(workbook)`: Adjusts the width of all columns in all sheets of a workbook to fit their content.
* `makeExcelFileRowsHeightAutoFit(workbook)`: Adjusts the height of all rows in all sheets of a workbook to fit their content.
* `makeExcelFileColumnsWidthAutoFit(workbook)`: Adjusts the width of all columns in all sheets of a workbook to fit their content.
* `addDataFrameAsSheetToExcel(dataframe, title, file_path)`: Adds a DataFrame as a new sheet.
* `saveToExcel(dataframe, save_path, file_subject)`: Saves to Excel with formatting.
* `saveDataFramesInExcel(dataframes, sheet_names, save_path, file_subject, freeze_position)`: Saves multiple DataFrames to one Excel.
* `makeExcelFileAutoFitWithFrozenPane(file_path, file_subject)`: Applies autofit and freeze to all sheets.
* `makeColumnsWidthAutoFit(worksheet)`: Adjusts column width.
* `makeRowsHeightAutoFit(worksheet)`: Adjusts row height.
* `freezePaneInExcelFile(workbook, freeze_position)`: Freezes pane at given cell.
* `removeSheetsFromExcelFile(file_path, sheet_names)`: Removes specified sheets.

### Reading & Preparing

* `extractDataFrame(file_path, sheet_names=None, ignored_sheets=None, headers_row_index=0, first_data_row=1, include_file_path=False, required_columns=None, reformat_names=True)`: Reads Excel or CSV files with flexible sheet/column selection and formatting.
* `prepareDataFrame(dataframe, file_path, headers_row_index=0, first_data_row=1, include_file_path=True, reformat_names=True)`: Cleans and standardizes DataFrame headers, adds file paths, and reformats column names.

---

## 📁 File Discovery

* `findCSVFilesBySubstring(folder_path, file_name_substring, recursive=True)`: Locates CSVs by name pattern.
* `findCSVFilesInFolder(folder_path, subdirectories, check_projection_system)`: Locates CSVs and checks projection.
* `findShapeFilesInFolder(folder_path, subdirectories)`: Locates shapefiles and projection systems.
* `findXLSXFilesBySubstring(folder_path, file_name_substring, recursive=True)`: Locates XLSX files by name pattern.

---

## 💾 Data Persistence & Bulk Loading

### Checkpoint Management

* `getCheckpointFileName(base_directory, name)`: Generates sanitized checkpoint file path.
* `hasCheckpoint(base_directory, checkpoint_name)`: Checks if checkpoint file exists.
* `loadCheckpoint(base_directory, checkpoint_name)`: Loads checkpoint data from pickle file.
* `saveCheckpoint(base_directory, checkpoint_name, data)`: Saves data to checkpoint file.

### Bulk File Loading

* `loadCSVFilesIntoDataFrames(folder_path, recursive=True, required_columns=None)`: Loads all CSV files from folder into list of DataFrames.
* `loadExcelFilesIntoDataFrames(folder_path, recursive=True, required_columns=None, reformat_names=True)`: Loads all Excel files from folder into list of DataFrames.

### File Handling

* `sanitizeFilename(name)`: Removes special characters from filename strings.
* `saveToCSV(dataframe, save_path, file_subject="")`: Saves DataFrame to CSV file with file-lock checking.

---

## 🔄 DataFrame Merging

* `mergeDataFramesHorizontallyOnCommonColumns(dataframes, data_frames_names)`: Merge multiple DataFrames on common columns.
* `mergeDataFramesHorizontallyOnSpecificColumns(dataframes, data_frames_names, merging_columns)`: Merge multiple DataFrames using specific columns.
* `mergeDataFramesVertically(dataframes, type_names=None, type_column=None, insert_index=0)`: Vertical merge/concatenation with optional type label column.
* `mergeSheetsHorizontallyOnSpecificColumns(file_path, merging_columns, selected_sheets=None)`: Merge Excel sheets horizontally on specific columns.
* `mergeSheetsVertically(file_path, selected_sheets=None, column_name_for_sheet_titles=None, sheet_titles=None)`: Merge Excel sheets vertically into single DataFrame.

---

## 📀 Utilities

* `generateCombinations(items, max_count)`: Generates combinations of items.
* `fillDataFrameByAnotherDataFrame(source_dataframe, destination_dataframe, source_columns, destination_columns)`: Fills columns in one DataFrame using another.
* `isNumber(value)`: Checks if value is a number.
* `isFileOpen(file_path)`: Checks if file is locked.
* `checkFileIsClosedBeforeSave(save_path)`: Shows warning until file is closed.
* `evaluateAndSplitLocation(location)`: Splits location string into latitude and longitude.
* `normalizeDataFrame(dataframe, keys, ignoring_columns, variance_check)`: Normalizes DataFrame by removing redundant columns.
* `makeAverageOnDataframe(dataframe, keys, check_numerics, fill_missing_values)`: Averages values across groups.
