# General Utilities Module

This module contains a comprehensive collection of utility functions designed to simplify common tasks related to data handling, file operations, statistical comparisons, Excel manipulations, and geospatial processing. Below is a detailed overview of all available functions organized by category.

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

* `addEmptyColumnsToDataframe(dataframe: DataFrame, columns: list[str], dtype: str = None, overwrite: bool = False)`: Adds new empty columns with optional data types.
* `addLatAndLongColumnsToDataframe(dataframe: DataFrame, location_column: str = 'location', lat_column: str = 'latitude', lon_column: str = 'longitude', remove_location: bool = True)`: Splits location string into separate latitude and longitude columns.
* `addPrefixesToColumnNames(dataframe: DataFrame, column_names: list[str] = None, prefixes: list[str] | str = "df")`: Adds specified prefixes to given column names.
* `addSuffixesToColumnNames(dataframe: DataFrame, column_names: list[str] = None, suffixes: list[str] | str = "df")`: Adds specified suffixes to given column names.
* `categorizeColumnsByType(dataframe: DataFrame, keys: list = [], ignoring_columns: list = [])`: Categorizes columns into string or numeric types.
* `reduceColumns(dataframe: DataFrame, columns_to_keep: list[str] = None, columns_to_drop: list[str] = None)`: Keeps or drops specified columns.
* `reorderColumnsOfDataFrame(dataframe: DataFrame, starter_columns: list[str], column_before_starters: str = None)`: Reorders columns with specific ones at start.
* `replaceColumnNameOfDataFrame(dataframe: DataFrame, old_substrings: list[str], new_substrings: list[str])`: Renames columns by replacing substrings.

### Missing Values & Cleaning

* `highlightDataFrameMissingValues(dataframe: DataFrame)`: Returns styled DataFrame highlighting missing values.
* `highlightMissingValue(column: Series)`: Helper that returns CSS styling for missing values.
* `convertAllStringNumericsToNumeric(dataframe: DataFrame, ignoring_columns: list[str] = [])`: Converts all string-based numerics to numeric type.
* `convertDataFrameStringNumericToNumeric(dataframe: DataFrame, numeric_columns: list[str] = None, ignoring_columns: list[str] = None)`: Converts specific columns to numeric.
* `convertStringNumericToNumeric(value: str)`: Converts single string numeric value to numeric type.
* `removeDuplicateData(dataframe: DataFrame, ignoring_columns: list = [])`: Removes duplicate rows and columns.
* `removeEmptyRows(dataframe: DataFrame, columns_to_check: list[str])`: Removes rows with all NaN in specified columns.
* `normalizeDataFrame(dataframe: DataFrame, keys: list = [], ignoring_columns: list = [], variance_check: bool = True)`: Normalizes DataFrame by removing redundant and null columns.

### Column Expansion & Splitting

* `expandColumns(dataframe: DataFrame, source_columns: list[str], destination_columns: list[str | list[str]], string_separators: list[list[str] | str], remove_source_columns: bool = False)`: Splits columns into multiple based on delimiters and regex patterns.
* `splitDataFrameHorizontally(dataframe: DataFrame, common_columns: list[str], columns_to_split: list[str])`: Splits DataFrame horizontally (into multiple DataFrames).
* `splitDataFrameVertically(dataframe: DataFrame, grouping_columns: list[str])`: Splits DataFrame vertically based on groupings.
* `splitDataFrameVerticallyIntoExcelFiles(dataframe: DataFrame, grouping_columns: list[str], save_folder: str, file_name_prefix: str = None, data_value_as_file_name: bool = True)`: Exports vertical splits to separate Excel files.

### Row/Value Helpers

* `replaceSubstringsInDataFrame(dataframe: DataFrame, column_names: list[str], old_substrings: list[list[str]], new_substrings: list[list[str]])`: Replaces substrings in specified columns.
* `roundCoordinates(dataframe: DataFrame, coordinate_columns: list[str], precision_digits: list[int])`: Rounds coordinates to specified precision.
* `uniqueValuesCount(values: Series)`: Counts unique non-null values.

---

## 📈 Statistical Utilities

### Comparison Metrics

* `calculateR2(observed: list, predicted: list)`: Calculates R-squared coefficient.
* `calculateRMSD(observed: list, predicted: list)`: Calculates Root Mean Square Deviation.
* `calculateDegreeOfAgreement(observed: list, predicted: list)`: Computes Willmott's degree of agreement (d statistic).

### Statistical Comparisons

* `compareValueRangesMathematically(first_list: list, second_list: list)`: Compares two sets using R², RMSD, and agreement.
* `compareByBiasCorrection(observed: list, predicted: list)`: Applies bias correction to predictions.
* `compareKDE(observed: list, predicted: list)`: Compares Kernel Density Estimations.
* `compareRangesDifferenceByQuantiles(observed: list, predicted: list, quantiles=[0.25, 0.5, 0.75])`: Quantile comparison.
* `compareRangesDifferenceByKLDivergence(first_list: list, second_list: list)`: KL Divergence comparison.
* `compareRangesDifferenceByKSTest(first_list: list, second_list: list)`: KS Test comparison.
* `compareRangesDifferenceByMannWhitney(first_list: list, second_list: list)`: Mann-Whitney U test comparison.
* `getNormalRangesDifference(first_list: list, second_list: list)`: Difference by range width.
* `getVariantRangesDifference(first_list: list, second_list: list, acceptable_percent: int = 10)`: Combines multiple statistical differences.

---

## 🗓️ Date and Time

* `convertToDatetime(value: str, source_format: str = None)`: Converts various formats to datetime.
* `changeDateTimeFormatInDataFrame(dataframe: DataFrame, column_names: list[str], new_formats: list[str])`: Changes date format of specified columns.
* `insertDateByTimestampIntoDataFrame(dataframe: DataFrame, timestamp_column: str = 'timestamp', date_column_name: str = 'date')`: Inserts date column from timestamp.
* `insertYearByTimestampIntoDataFrame(dataframe: DataFrame, timestamp_column: str = 'timestamp', year_column_name: str = 'year')`: Extracts year from timestamp.
* `addNewDateColumnByDateRangesToDataFrame(dataframe: DataFrame, column_name: str, date_ranges: list, new_date_column_name: str, new_date_format: str)`: Creates date columns based on date ranges.

---

## 🧽 Geospatial Utilities

* `convertShapeFileDataToDataFrame(file_path: str, projection_system: str)`: Converts shapefile to DataFrame with projection conversion.
* `extractCSVDataIntoDataFrame(file_path: str, file_projection_system: str, destination_projection_system: str)`: Extracts CSV and transforms projection system.
* `getDominantProjectionSystem(shape_files_path: list)`: Finds most common CRS in shape files.
* `getDominantProjectionSystemOfCSVFiles(csv_files_path: list)`: Finds common projection in CSV files with coordinate detection.

---

## 📂 File and Sheet Handling

### Excel Operations

* `addDataFrameAsSheetToExcel(dataframe: DataFrame, title: str, file_path: str)`: Adds a DataFrame as new sheet to Excel file.
* `makeExcelFileColumnsWidthAutoFit(workbook: Workbook)`: Adjusts width of all columns in all sheets.
* `makeExcelFileRowsHeightAutoFit(workbook: Workbook)`: Adjusts height of all rows in all sheets.
* `makeColumnsWidthAutoFit(worksheet: worksheet)`: Adjusts column width for a worksheet.
* `makeRowsHeightAutoFit(worksheet: worksheet)`: Adjusts row height for a worksheet.
* `makeExcelFileAutoFitWithFrozenPane(file_path: str, file_subject: str)`: Applies auto-fit and freeze panes.
* `saveToExcel(dataframe: DataFrame, save_path: str, file_subject: str = "")`: Saves DataFrame to Excel with formatting.
* `saveDataFramesInExcel(dataframes: list[DataFrame], sheet_names: list[str], save_path: str, file_subject: str = "", freeze_position: tuple = None)`: Saves multiple DataFrames to Excel.
* `freezePaneInExcelFile(workbook: Workbook, freeze_position: tuple = None)`: Freezes pane at given position.
* `removeSheetsFromExcelFile(file_path: str, sheet_names: list[str])`: Removes specified sheets.

### Reading & Preparing Data

* `extractDataFrame(file_path: str, selected_sheets: list[str] = None, ignored_sheets: list[str] = None, headers_row_index: int = 0, first_data_row: int = 1, include_file_path: bool = False, required_columns: list[str] = None, reformat_names: bool = True)`: Reads Excel or CSV with flexible options.
* `prepareDataFrame(dataframe: DataFrame, include_file_path: bool = False, file_path: str = "", headers_row_index: int = 0, first_data_row: int = 1, reformat_names: bool = True)`: Cleans and standardizes DataFrame headers and formats data.

---

## 📁 File Discovery

* `findCSVFilesBySubstring(folder_path: str, file_name_substring: str = None, recursive: bool = True)`: Locates CSV files by pattern.
* `findCSVFilesInFolder(folder_path: str, subdirectories: list = None, check_projection_system: bool = True)`: Locates CSV files and checks projection.
* `findShapeFilesInFolder(folder_path: str, subdirectories: list = None)`: Locates shapefiles with projection detection.
* `findXLSXFilesBySubstring(folder_path: str, file_name_substring: str = None, recursive: bool = True)`: Locates XLSX files by pattern.

---

## 💾 Data Persistence & Bulk Loading

### Checkpoint Management

* `getCheckpointFileName(base_directory: str, name: str)`: Generates sanitized checkpoint file path with .pkl extension.
* `hasCheckpoint(base_directory: str, checkpoint_name: str)`: Checks if checkpoint file exists.
* `loadCheckpoint(base_directory: str, checkpoint_name: str)`: Loads checkpoint data from pickle file.
* `saveCheckpoint(base_directory: str, checkpoint_name: str, data)`: Saves data to checkpoint file.

### Bulk File Loading

* `loadCSVFilesIntoDataFrames(folder_path: str, recursive: bool = True, required_columns: list[str] = None)`: Loads all CSV files into list of DataFrames.
* `loadExcelFilesIntoDataFrames(folder_path: str, recursive: bool = True, required_columns: list[str] = None, reformat_names: bool = True)`: Loads all Excel files into list of DataFrames.

### File Utilities

* `sanitizeFilename(name: str)`: Removes special characters from filename strings.
* `saveToCSV(dataframe: DataFrame, save_path: str, file_subject: str = "")`: Saves DataFrame to CSV with file-lock checking.
* `isFileOpen(file_path: str)`: Checks if file is currently open or locked.
* `checkFileIsClosedBeforeSave(save_path: str)`: Shows warning dialog until file is closed.
* `evaluateAndSplitLocation(location: str)`: Splits location string "lat, lon" into tuple.

---

## 🔄 DataFrame Merging & Combining

* `mergeDataFramesHorizontallyOnCommonColumns(dataframes: list[DataFrame], data_frame_names: list[str])`: Merges DataFrames on common columns.
* `mergeDataFramesHorizontallyOnSpecificColumns(dataframes: list[DataFrame], data_frame_names: list[str], merging_columns: list[str])`: Merges on specific columns.
* `mergeDataFramesVertically(dataframes: list[DataFrame], type_names: list[str] = None, type_column: str = None, insert_index: int = 0)`: Vertical concatenation with optional type column.
* `mergeSheetsHorizontallyOnSpecificColumns(file_path: str, merging_columns: list[str], selected_sheets: list[str] = None)`: Merges Excel sheets horizontally.
* `mergeSheetsVertically(file_path: str, selected_sheets: list[str] = None, column_name_for_sheet_titles: str = None, sheet_titles: list[str] = None)`: Merges Excel sheets vertically.

---

## 🔧 General Utilities

* `generateCombinations(items: list, max_count: int)`: Generates combinations of items.
* `fillDataFrameByAnotherDataFrame(source_dataframe: DataFrame, destination_dataframe: DataFrame, source_columns: list[str], destination_columns: list[str])`: Fills columns from another DataFrame.
* `isNumber(value)`: Checks if value is a number.
* `makeAverageOnDataframe(dataframe: DataFrame, keys: list, check_numerics: bool = False, fill_missing_values: bool = True)`: Groups and aggregates DataFrame.

---

## 🔗 Dependencies

* `pandas` - DataFrame manipulation
* `openpyxl` - Excel operations
* `pyproj` - Coordinate transformations
* `scipy.stats` - Statistical functions
* `sklearn` - Machine learning utilities
* `fiona` - GIS file handling
* `geopandas` - Geospatial data handling
* `tkinter` - GUI components

---

## 📝 General Notes

* Functions handle None and NaN values gracefully
* Date conversion supports multiple formats automatically
* Column names standardized to lowercase with underscores
* Geospatial functions use EPSG codes
* Checkpoint functions use pickle format
* Excel operations include auto-fit and freeze panes
