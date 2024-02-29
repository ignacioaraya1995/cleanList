# Property Data Filtering Script

This script processes a dataset of properties from an Excel file, applying multiple filters to exclude certain entries based on predefined criteria and then outputting the filtered list to a new Excel file. The criteria include excluding properties based on specific tags, removing entries with a "Dead Lead" status, and more.

## Features

- Filters properties based on exclusion tags listed in a separate Excel file.
- Removes properties with a "Dead Lead" status.
- Excludes properties with a BUYBOX SCORE of 0.
- Ignores properties with a blank MAILING ADDRESS.
- Removes duplicate entries based on MAILING ADDRESS and MAILING ZIP, keeping only the entry with the highest score.
- Outputs the filtered dataset to a new Excel file.

## Prerequisites

Before running this script, you will need:

- Python (3.6 or later)
- Pandas library
- An Excel file named `data.xlsx` containing your dataset
- An Excel file named `tags.xlsx` containing tags to exclude

## Installation

### Installing Python

To run this script, you need Python installed on your system. If you do not have Python installed, follow these steps:

#### Windows

1. Visit the official Python website at [python.org](https://www.python.org/downloads/).
2. Download the latest version of Python for Windows.
3. Run the installer, ensuring that you check the box to "Add Python to PATH" during installation.

#### macOS

1. You can install Python using Homebrew (a package manager for macOS). If you do not have Homebrew installed, open Terminal and run:
   ```
   /bin/bash -c "$(curl -fsSL https://raw.githubusercontent.com/Homebrew/install/HEAD/install.sh)"
   ```
2. Then, install Python by running:
   ```
   brew install python
   ```

#### Linux

Most Linux distributions come with Python pre-installed. To check if Python is installed and to see the version, open a terminal and run:
```
python3 --version
```
If Python is not installed or you need a different version, you can typically install it using your distribution's package manager. For example, on Ubuntu, you can install Python by running:
```
sudo apt-get update
sudo apt-get install python3
```

### Installing Required Libraries

After installing Python, you can install the required libraries using `pip`, Python's package installer. Run the following command in your terminal or command prompt:

```
pip install pandas openpyxl
```

## Usage

1. Ensure you have the `data.xlsx` and `tags.xlsx` files in the same directory as the script.
2. Open your terminal or command prompt and navigate to the directory containing the script.
3. Run the script with the command:
   ```
   python script_name.py
   ```
   Replace `script_name.py` with the name of your Python script.

4. The script will process the data and output the results to a new Excel file named `filtered_properties.xlsx`. Progress will be printed to the console.

## License

This project is open source and available under the [MIT License](https://opensource.org/licenses/MIT).

---

Replace `script_name.py` with the actual filename of your Python script. You can also customize the License section as needed, depending on your preferences or requirements.
