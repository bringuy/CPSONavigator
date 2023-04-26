## Overview
A web scraper/navigator which goes through the website [CPSO](https://www.cpso.on.ca/) and gets the contact information of all the doctors in the Tricity area (Kitchener, Waterloo, Cambridge) area. 
Data such as the doctor's name, address, and contact information are all stored in an Excel sheet. Uses Python and Selenium Webdriver.

## Installation
1. Clone the repository to your local machine
2. Install Python 3.x and pip.
3. Install the required packages using the following command:
```
pip install -r requirements.txt
```
4. Install the latest version of the Chrome browser.

## Usage
Run the following command in the terminal to execute the script:
```
python main.py
```
Wait for the script to complete. The script may take several minutes to run depending on the number of doctors.
Once the script has completed, an Excel sheet named  **GP_Demographics.xlsx** containing the data will be saved in the project directory.
