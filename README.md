# Skills Matrix Automation Bot

This Python script automates the process of updating a Skills Matrix in the rapport3 system. The script leverages Selenium for browser automation, enabling efficient interaction with the system's web interface to input data from an Excel spreadsheet.

## Table of Contents

- [Prerequisites](#prerequisites)
- [Setup](#setup)
- [Usage](#usage)
- [Folder Structure](#folder-structure)
- [Disclaimer](#disclaimer)
- [License](#license)

## Prerequisites

Before running this script, ensure you have the following installed:

- [Python 3.x](https://www.python.org/downloads/)
- [Selenium](https://pypi.org/project/selenium/)
- [Pandas](https://pandas.pydata.org/)
- [Google Chrome](https://www.google.com/chrome/) (ensure the browser is up-to-date)
- [ChromeDriver](https://sites.google.com/a/chromium.org/chromedriver/downloads) (Ensure the path in the script matches your local installation)

## Setup

1. **Clone the repository**:
   ```bash
   git clone https://github.com/your-username/skills-matrix-automation-bot.git
   cd skills-matrix-automation-bot
   ```

2. **Install required Python packages**:
   ```bash
   pip install -r requirements.txt
   ```

3. **Configure the Script**:
   - Ensure the `PATH` variable points to the correct location of your `chromedriver.exe`.
   - Save your session cookies in the specified `user-data-dir` to bypass manual login steps. This script assumes you have already logged in and saved the session.
   - Place your Excel file (`New_Skills_Matrix_Everyone.xlsx`) in the `Skills Matrix_rapport data imput/Skills_Matrix/` directory.

## Usage

Run the script to automate the Skills Matrix update:

```bash
python skills_matrix_automation.py
```

The script will:
- Load the skills data from the Excel file.
- Navigate through the rapport3 system and update each skill entry based on the data provided in the Excel file.

### Important Notes

- **Browser Control**: The script will take control of your Chrome browser. Do not interact with the browser during the script execution to avoid interference.
- **Data Integrity**: Ensure that the data in the Excel file is correctly formatted and that the `Competency_Value` and `Experience_Value` columns contain valid integers between 0 and 5.

## Folder Structure

```bash
.
├── Skills Matrix_rapport data imput/
│   └── Skills_Matrix/
│       └── New_Skills_Matrix_Everyone.xlsx
├── cookies/             # Directory storing user session cookies for rapport3 login
├── chromedriver.exe     # ChromeDriver executable
└── skills_matrix_automation.py  # Main automation script
```

## Disclaimer

This script is intended for educational and personal use. Ensure you have permission to interact with the rapport3 system in this manner. Use it responsibly, especially when handling sensitive information.

## License

This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.
