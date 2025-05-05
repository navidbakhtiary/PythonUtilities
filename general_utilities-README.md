# General Utils - Python Utility Library

## 📦 Overview

`general_utilities.py` is a versatile Python module containing utility functions for:

- 🧮 **Pandas DataFrame operations**
- 📊 **Statistical evaluations**
- 📁 **File handling and Excel interaction**
- 🌍 **Geospatial data processing**

These functions are reusable in a wide variety of data science, data engineering, and GIS tasks.

---

## 🚀 How to Use

```bash
pip install pandas numpy openpyxl geopandas fiona pyproj
```

```python
from general_utils import addPrefixesToColumnNames, calculateR2, convertShapeFileDataToDataFrame
```

---

## 📚 Function Categories

### 📄 DataFrame Utilities

| Function                            | Description                                   |
| ----------------------------------- | --------------------------------------------- |
| `addPrefixesToColumnNames`          | Add a prefix to column names in a DataFrame   |
| `addSuffixesToColumnNames`          | Add a suffix to column names in a DataFrame   |
| `categorizeColumnsByType`           | Split columns into string/numeric types       |
| `convertAllStringNumericsToNumeric` | Convert string-like numbers to numeric values |
| `expandColumns`                     | Split delimited column into multiple columns  |
| `highlightDataFrameMissingValues`   | Style DataFrame to show missing cells         |

### 📈 Math & Statistics Utilities

| Function                     | Description                        |
| ---------------------------- | ---------------------------------- |
| `calculateR2`                | Compute R² score                   |
| `calculateRMSD`              | Compute Root Mean Square Deviation |
| `calculateDegreeOfAgreement` | Compute Willmott’s D index         |

### 📁 File & Excel Utilities

| Function                      | Description                         |
| ----------------------------- | ----------------------------------- |
| `checkFileIsClosedBeforeSave` | Prompt to close file before writing |
| `isFileOpen`                  | Detect if file is locked or open    |

### 🌐 General Data Utilities

| Function                                   | Description                                          |
| ------------------------------------------ | ---------------------------------------------------- |
| `addLatAndLongColumnsToDataframe`          | Parse 'location' string into latitude and longitude  |
| `addNewDateColumnByDateRangesToDataFrame`  | Add a column based on date range grouping            |
| `convertShapeFileDataToDataFrame`          | Convert shapefile points to a DataFrame              |
| `extractDataFrame`                         | Extract CSV or Excel into a DataFrame                |
| `prepareDataFrame`                         | Clean and format DataFrame with standard headers     |
| `findCSVFilesBySubstring`                  | Find all CSVs with a given substring in filename     |
| `findCSVFilesInFolder`                     | Recursively search for CSVs in a folder              |
| `findShapeFilesInFolder`                   | Find and classify shapefiles by projection           |
| `findXLSXFilesBySubstring`                 | Find XLSX files matching a name pattern              |
| `getDominantProjectionSystem`              | Find most common projection among shapefiles         |
| `getDominantProjectionSystemOfCSVFiles`    | Estimate projection system from coordinate ranges    |
| `mergeSheetsHorizontallyOnSpecificColumns` | Merge Excel sheets horizontally by key columns       |
| `mergeSheetsVertically`                    | Stack Excel sheets vertically with sheet names added |

---

## 🛠️ Contributing

Feel free to submit a pull request or issue if you'd like to add more reusable functions or suggest improvements.

---

## 📄 License

MIT License © 2025 navidbakhtiary
