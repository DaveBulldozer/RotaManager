# Rota Management System

## Overview
The Rota Management System is designed to streamline shift scheduling and reporting within Google Sheets. This tool introduces a comprehensive menu in your Google Sheets environment to help manage employee rosters, check for schedule conflicts, generate various reports, and manage shift-related data more efficiently.

## Features
- **Trim White Spaces:** Clean up cells by removing leading and trailing spaces.
- **Check for Clashes:** Highlight cells in the schedule where employee shift clashes occur.
- **Generate Shift Report:** Create detailed reports for specific shifts including date, type, location, and hours.
- **Generate Full Report:** Aggregate reports for all employees over a specified period.
- **Generate Location Report:** Summarize employee schedules by location.
- **Query by Total Hours:** Filter shifts based on total hours worked.
- **Sort Reports:** Generate reports sorted alphabetically, by least hours, or most hours worked.
- **Name and Total Hours Report:** List each employee along with their accumulated working hours.

## Setup
1. **Open your Google Sheets where you want to install the script.**
2. **Go to Extensions > Apps Script.**
3. **Delete any code in the script editor and paste the new code from the `Code.gs` file.**
4. **Save and name your project.**
5. **Close the Apps Script tab and refresh your Google Sheets.**

After refreshing, a new menu item titled 'Rota Management' will appear in your Google Sheets menu. From here, you can access all the functionalities listed above.

## Technologies
- Google Apps Script

## Usage
Once installed, the Rota Management menu will be accessible from within your Google Sheets. Here’s how to use the individual features:

- **Trim White Spaces:** Select this option to clean up the selected cell range in your sheet.
- **Check for Clashes:** This function will check for any scheduling conflicts and highlight them in red.
- **Generate Reports:** Choose the type of report you need to generate from the sub-menu options. You will be prompted to enter necessary details like names or dates depending on the report.

## Contributing
Contributions are what make the open-source community such an amazing place to learn, inspire, and create. Any contributions you make are **greatly appreciated**.

1. Fork the Project
2. Create your Feature Branch (`git checkout -b feature/AmazingFeature`)
3. Commit your Changes (`git commit -m 'Add some AmazingFeature'`)
4. Push to the Branch (`git push origin feature/AmazingFeature`)
5. Open a Pull Request

## License
Distributed under the MIT License. See `LICENSE` for more information.

## Contact
Your Name - [Your Email](mailto:your-email@example.com)

## Acknowledgments
- Google Apps Script documentation
- Stack Overflow
- And any other contributor or site that helped your project!
