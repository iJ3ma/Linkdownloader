# Excel Link Downloader

## Overview
The Excel Link Downloader is a Python application designed to extract hyperlinks and URLs from Excel files and websites, download the linked files, and manage the downloads through a user-friendly graphical interface.

## Features
- Extracts links from Excel files (.xlsx).
- Scrapes links from specified websites.
- Downloads files concurrently with configurable settings.
- Provides a log of successful and failed downloads.

## Installation

1. Clone the repository:
   ```
   git clone https://github.com/yourusername/excel-link-downloader.git
   cd excel-link-downloader
   ```

2. Install the required dependencies:
   ```
   pip install -r requirements.txt
   ```

3. Run the application:
   ```
   python -m src.excel_downloader
   ```

## Usage
- Select one or more Excel files to extract links from.
- Add website URLs to scrape for downloadable links.
- Choose a main download folder where files will be saved.
- Adjust settings for maximum concurrent downloads and timeout duration.
- Click "Start Download" to begin the process.

## Contributing
Contributions are welcome! If you have suggestions for improvements or find bugs, please open an issue or submit a pull request.

## License
This project is licensed under the MIT License - see the [LICENSE](LICENSE) file for details.