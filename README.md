# Rate-Scrapper

[![License: CC0-1.0](https://licensebuttons.net/l/zero/1.0/80x15.png)](http://creativecommons.org/publicdomain/zero/1.0/)

This PyQt5 application allows you to scrape currency rates from different web pages using an API. It provides a user-friendly interface for selecting currencies, web pages, report folder, and generating the currency rate report.

<br>

## Graphical User Interface

![Screenshot](https://raw.githubusercontent.com/lgrigor/Rate-Scrapper/main/Documentation/readme_main_frame.PNG)

## Features
* Scrapes currency rates from multiple web pages using an API.
* User-friendly interface with easy navigation and selection options.
* Allows selection of currencies to scrape.
* Supports selection of web pages to scrape.
* Allows selection of the report folder to save the generated report.
* Generates a currency rate report in a specified folder.

<br>

## Installation
* Download the `RateScraper.exe` file from the repository bin folder.
* Double-click the `RateScraper.exe` file to run the application.

<br>

## Usage
The application window will appear with a list of available currencies and web pages to scrape currency rates from. <br>
Select the desired currencies by double-checking the corresponding currencies. <br>
Alternatively, you can manually add currencies by following the syntax guidelines. <br>
In case of any errors, an error message will be displayed along with the appropriate usage instructions. <br>
<br>

<img src="https://raw.githubusercontent.com/lgrigor/Rate-Scrapper/main/Documentation/readme_frame_error_message.PNG" width="700" height="350">

<br>
Select the desired web page(s) by checking the corresponding checkboxes. <br>

<br>
Click the "Select Report Folder" button to choose the folder where you want to save the generated report. <br>
<br>

<img src="https://raw.githubusercontent.com/lgrigor/Rate-Scrapper/main/Documentation/selecting_folders.PNG" width="700" height="350">

<br>
Click the "Generate" button to initiate the scraping process and generate the currency rate report. <br>


<br>
Once the scraping is complete, the application will save the generated report in the selected folder. <br>
The report file is an Excel document that encompasses a comprehensive array of currencies and selected pages. <br>

It adheres to a naming convention denoted as `report_2023_06_01__21_45_00.xlsx` where `report` signifies the nature of the file, and the subsequent sequence of digits represents the year, month, day, hour, minute, and second of its creation. <br>


<br>
<img src="https://raw.githubusercontent.com/lgrigor/Rate-Scrapper/main/Documentation/readme_report_1.PNG" width="700" height="150">

The report file content: <br>
If the service page does not offer currency conversion functionality, the corresponding rate field will display a value of 0.<br>
<br>
<img src="https://raw.githubusercontent.com/lgrigor/Rate-Scrapper/main/Documentation/readme_report_2.PNG" width="700" height="250">



