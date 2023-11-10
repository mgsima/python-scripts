# python-scripts
Collection of small python scripts with different small projects

# Projects

## 1. Auto-blau-cover
This Python script is designed for manipulating and generating Word documents (.docx) using the python-docx library. It is particularly useful for:

- Extracting Text Styles: Retrieves styles such as font size, boldness, and alignment from a specified Word document template.
- Generating Documents from Text File: Creates a Word document with contents sourced from a text file ('tags.txt'), applying predefined formatting to each line.
- Creating JSON Structure from Directories: Generates a JSON file that maps the structure of subdirectories within a given directory path.
- Producing Formatted Word Documents: Based on the structure and contents of subdirectories in a specified path, this script creates two differently styled Word documents.

This suite of functionality is ideal for automating document creation and formatting tasks, leveraging directory data and template styles for efficient document management.

## 2. auto-thumbnails
This Python script is designed for creating contact sheets from image files using the Pillow library. It offers the following functionalities:

- Creating Contact Sheets: Generates contact sheets from a collection of images in a specified directory. Each sheet arranges images in a grid format with customizable columns and rows.
- Handling Multiple Image Formats: Supports a variety of image formats including jpg, png, jpeg, gif, bmp, and tiff.
- Customizable Layout and Text: Allows for setting image width, margin, and spacing. Adds image file names below each image using text wrapping for better readability.
- Output in Specified Directory: Saves the generated contact sheets in a user-specified output directory, naming them sequentially.

This script is particularly useful for organizing and displaying multiple images in a single document, ideal for photographic summaries, thumbnails, or quick reference sheets.

## 3. image-procesor
This Python script is designed for converting image files to PNG format using the Pillow library. Key features include:

- Command-Line Arguments: Accepts two command-line arguments for source and destination directories.
- Automatic Directory Creation: Checks for the existence of the destination directory and creates it if it doesn't exist.
- Batch Image Conversion: Iterates through all image files in the specified source directory, converting each to PNG format.
- Saving Converted Images: Places the converted PNG images into the designated destination directory, retaining the original file names.

This script is particularly useful for bulk conversion of images to PNG format, streamlining the process for large collections of images, such as a 'Pokedex' of images.

## 4. mail-python
This Python script is designed for sending HTML emails using the `smtplib` and `email` libraries. Its main features include:

- **HTML Email Template**: Utilizes an HTML template (`index.html`) for the email body, allowing customization of the email content.
- **Email Composition**: Constructs an email with a sender address, recipient address, subject, and HTML content. The content is dynamically populated using the `Template` class.
- **Secure SMTP Connection**: Establishes a secure connection with Gmail's SMTP server using SSL.
- **User Authentication**: Logs into the sender's Gmail account securely to send the email.
- **Sending Email**: The composed email is sent to the specified recipient.

This script is particularly useful for automating the process of sending customized HTML emails, making it suitable for newsletters, notifications, or personalized communication.

## 5. printer-automatization
This Python script is primarily designed for batch printing PDF documents and configuring printer settings on Windows using libraries like `win32api`, `win32print`, and `win32com`. The script's functionalities include:

- **Printer Configuration Display**: Retrieves and displays various settings and properties of the connected printers using `win32com`.
- **Document Sorting and Selection**: Sorts and selects PDF documents from a specified directory, prioritizing certain files based on naming conventions.
- **Batch Printing PDFs**: Automates the process of printing multiple PDF documents from a given folder. It sets the default printer, configures print settings, and sends files to the printer.
- **Advanced Printer Configuration**: Offers capabilities to modify and apply printer settings such as orientation, color, and paper source programmatically.

This script is especially useful in office settings or for tasks that require automated printing of multiple documents, along with the ability to programmatically control and display printer configurations.

## 6. scrapping-book
This Python script utilizes `selenium` and `BeautifulSoup` for web scraping and data extraction. Key functionalities include:

- **Web Scraping with Selenium and BeautifulSoup**: Automates the browser using Selenium to navigate web pages and extract HTML content. BeautifulSoup is then used to parse and process this HTML data.
- **Text Extraction from Web Pages**: Targets specific `<div>` elements on web pages, extracting and processing their textual content.
- **Iteration Over Multiple Pages**: Capable of iterating through a series of web pages and cumulatively extracting text from each.
- **Saving Extracted Data**: Compiles all the extracted text into a single string and saves it into a text file (`resultado.txt`).

This script is particularly useful for gathering text data from multiple web pages for content aggregation, analysis, or archival purposes.

## 7. python-exercises
Python script collection serves as a practice ground, sharpening core programming skills with a focus on practical applications and problem-solving.
