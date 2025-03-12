# ExcellentScraper

A powerful web scraping tool for extracting article content and exporting it to Excel.

## Features

- Modern dark/light UI using CustomTkinter
- Scrape up to 10 URLs simultaneously
- Intelligent content extraction algorithm that identifies real article content
- Advanced filtering of ads, navigation elements, and non-article content
- Smart title extraction from multiple sources (OpenGraph, Schema.org, Twitter cards)
- Automatic handling of different website layouts and structures
- Save results to Excel spreadsheets
- Merge Excel files to build a comprehensive dataset
- Detailed status logging
- Visual animations and feedback

## Requirements

- Python 3.6+
- Dependencies listed in `requirements.txt`

## Installation

1. Clone this repository or download the source code
2. Create a virtual environment (optional but recommended):
   ```
   python -m venv ExcellentScraper-env
   ```
3. Activate the virtual environment:
   - Windows: `ExcellentScraper-env\Scripts\activate`
   - Unix/MacOS: `source ExcellentScraper-env/bin/activate`
4. Install the required packages:
   ```
   pip install -r requirements.txt
   ```

## Usage

1. Run the application:
   ```
   python ExcellentScraper.py
   ```
2. Enter up to 10 URLs in the input fields
3. Click "Start Scraping" to begin the extraction process
4. The application will create a new Excel file in the `scraped_data` directory
5. To merge Excel files, click the "Merge Excel Files" button

## How It Works

The application uses a sophisticated multi-tiered approach to web scraping:

1. First, it attempts to scrape with BeautifulSoup (faster, lighter) with realistic browser headers
2. If BeautifulSoup fails, it falls back to Selenium (more powerful for dynamic content)
3. The scraper uses multiple strategies to identify article content:
   - Recognizes common article container elements
   - Analyzes content density to find the most relevant text
   - Filters out non-content elements like ads, navigation, footers, etc.
   - Prioritizes proper paragraph text over other elements
   - Handles different website layouts and structures
4. Title extraction follows a priority-based approach:
   - Looks for structured data (Schema.org article headline)
   - Checks for OpenGraph and Twitter card metadata
   - Analyzes heading elements for the most relevant title
   - Cleans titles by removing site names and separators
5. Data is organized in Excel with timestamps, URLs, headings, and content

## Advanced Content Extraction

ExcellentScraper uses several techniques to ensure high-quality content extraction:

- **Content recognition**: Identifies the main article content from various page layouts
- **Element filtering**: Removes sidebars, ads, related content, navigation, and other non-article elements
- **Content density analysis**: Finds the section of the page with the highest concentration of relevant text
- **Paragraph extraction**: Focuses on proper article paragraphs while filtering out non-content text
- **Title prioritization**: Uses multiple sources to find the most accurate article title
- **Adaptable strategy**: Falls back to progressively more aggressive extraction methods if needed

## Excel Output Format

- Column A: Timestamp of when the scraping was completed
- Column B: URL of the scraped article
- Column C: First heading (usually the article title)
- Column D: Full article content
- Additional columns: Additional headings found in the article

## Merging Excel Files

The "Merge Excel Files" functionality allows you to combine multiple scraped datasets:
- If you select an existing file, new data will be appended to it
- This enables building a comprehensive dataset over time

## Troubleshooting

- **Poor quality content extraction**: Some websites use unusual layouts that might confuse the scraper. Try using the Selenium method for these sites by intentionally causing the BeautifulSoup method to fail (e.g., by using an invalid header).
- **Missing content**: For sites with heavy JavaScript rendering, the app will automatically use Selenium, but you may need to increase the wait time for very complex sites.
- **Unwanted content included**: If you notice ads or navigation elements in your extracted content, please report the website structure for future improvements.
- **Character encoding issues**: The app attempts to detect the correct encoding, but some sites might use non-standard encoding. Manual cleanup might be required.
- **Selenium issues**: Ensure you have a compatible version of Chrome installed. The app uses webdriver-manager to download the appropriate driver.
- **Empty content**: Some websites implement anti-scraping measures. The app uses realistic headers to avoid detection, but some sites might still block automated access.

## Future Improvements

- Site-specific extraction rules for popular websites
- Support for paywalled content (with user credentials)
- Content categorization and sentiment analysis
- Image extraction and inclusion in reports
- Support for additional export formats (JSON, CSV, etc.)

## License

This project is open-source and available under the MIT License. 