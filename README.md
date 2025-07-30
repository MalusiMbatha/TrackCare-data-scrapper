Script Name: GetDataFromWeb

Author: Malusi Mbatha  
Description:  
This VBA script automates data extraction from the NHLS TrakCare WebView portal. It performs the following actions:

1. Launches Internet Explorer and navigates to the login page.
2. Prompts the user to manually log in and apply necessary filters.
3. Scrapes patient data (e.g., Name, MRN, Date of Birth, Location) from the results table.
4. Writes the extracted data into a new Excel workbook.
5. Handles multi-page navigation until all data is retrieved.
6. Closes Internet Explorer upon completion.

Requirements:  
- Windows OS  
- Internet Explorer (must be installed and functional)  
- Excel application (required for writing output)  

Usage Instructions:  
1. Run the script from within a VBA-enabled environment (e.g., Excel or VBScript host).  
2. Log in manually when prompted and apply necessary filters.  
3. Press OK to continue data extraction.  
4. The script will generate and populate an Excel sheet with the retrieved data.

Note:  
Ensure Internet Explorer is not blocked by system policies, and TrakCare WebView is accessible.

Status:  
Tested and functional as of July 2025.
