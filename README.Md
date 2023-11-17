# Asset Manager

Welcome to the Asset Manager project repository! This Excel-based project is designed to help organizations efficiently manage and track their assets, providing insights into inventory, state, and more.

## Table of Contents

- [Overview](#overview)
- [Features](#features)
- [How to Use](#how-to-use)
- [File Structure](#file-structure)
- [Code Samples](#code-samples)
- [Future Enhancements](#future-enhancements)
- [Contributing](#contributing)
- [License](#license)

## Overview

The Asset Manager is an Excel VBA project that enables organizations to keep detailed records of their assets, from gadgets to stationeries. It includes features such as inventory tracking, state assessment, and a user-friendly interface for easy management.

## Features

- **Inventory Management:** Keep a comprehensive inventory of all assets, including details like item code, purchase price, and current value.
- **State Assessment:** Automatically assess the state of assets (Usable, Damaged, Repair) based on quantities, providing a quick overview.
- **User Interface:** Intuitive user interface with buttons, shapes, and macros to facilitate easy navigation and asset summary generation.
- **User Forms:** Utilizes user forms for searching, adding new assets, and updating inventory, enhancing user experience.

## How to Use

To utilize the Asset Manager effectively, follow these steps:

1. Open the "AssetManager.xlsm" Excel file.
2. Explore the different sheets, including "Product Master," "Stock In-Out," and "Inventory."
3. Utilize the user forms ("frm_Search_Product" and "frm_Update_Inventory") for searching, adding, and updating assets.
4. Leverage the buttons and shapes on the interface for quick summaries and actions.

## File Structure

- **Code:** Contains VBA code for various functionalities, including user forms, macros, and state assessment.
- **Documentation:** Houses documentation files, including project breakdowns, code explanations, and future enhancements.
- **Assets:** Store any relevant images or assets used in the project.

## Code Samples

Here are snippets of key VBA code used in the project:

#Code Sample 1: Search and Display Products: 

' frm_Search_Product Userform initialization event


Private Sub UserForm_Initialize()

    Call Show_Product

End Sub

' Code for searching and displaying products


Sub Show_Product()


    ' ... (Code for populating ListBox1 based on search criteria)

End Sub

' Code to be executed when "See More Details..." button is clicked


Private Sub ListBox1_DblClick(ByVal Cancel As MSForms.ReturnBoolean)

    Call Search_Product

End Sub

' Additional code for displaying details on ListBox1 item change


Private Sub ListBox1_Change()


    ' ... (Code for displaying details of the selected item)

End Sub

' Code to be executed when "See More Details..." button is clicked


Private Sub Search_Product()


    ' ... (Code for copying data to "Report" sheet and hiding the userform)


End Sub

#Code Sample 2: Updating Inventory: 


' Code for updating inventory when Submit button is clicked


Private Sub Submit_Click()

    ' ... (Code for validating input, updating "Product Master," "Stock In-Out," and "Inventory" sheets)

    ' ... (Code for refreshing form fields)

    Call Refresh

    MsgBox "Updated!!!"

End Sub

' Refresh function to clear form fields

Sub Refresh()
    ' ... (Code for resetting form fields to default values)
End Sub

#Code Sample 3: Refresh Macro

' Refresh function to clear form fields


    ' ... (Code for resetting form fields to default values)

End Sub


## Future Enhancements

We have exciting plans to enhance the Asset Manager:

- **Connectivity:** Explore integrating the software with online platforms to retrieve live data on asset prices.
- **Web Scraping:** Develop functionalities to scrape information from websites like Amazon and eBay to compare prices.
- **User Authentication:** Implement user authentication for secure access to sensitive data.

Feel free to contribute ideas or collaborate on these enhancements!

## Contributing

We welcome contributions! If you have ideas, improvements, or bug fixes, please open an issue or submit a pull request.

## License

This project is licensed under the [MIT License](LICENSE), allowing you to use, modify, and distribute the code freely.

Happy Asset Managing!
