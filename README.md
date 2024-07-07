# Excel Web Application

This is a web-based application built with ASP.NET Core that allows users to work with Excel sheets, import data, perform calculations, and save the modified sheet. 

## Features

- **Import Excel Sheet:** Upload the Taxes Excel sheet via the web interface.
- **Add a Column:** Adds a new column called “Total Value before Taxing” to the Excel sheet, computing the total value before applying taxes.
- **Calculate Total:** Adds a new row at the end of the sheet, calculating the sum of the “Total Value After Taxing” column.
- **Save Modified Sheet:** Ensures that the changes are reflected in the saved file.
- **Display Result:** Displays the final result (sum of “Total Value After Taxing”) to the user on the application interface.

## Technologies and Patterns Used

### 1. Dependency Injection

The project leverages ASP.NET Core's built-in Dependency Injection (DI) to manage service lifetimes and dependencies. DI helps in keeping the application loosely coupled and easier to test.

### 2.2. Three-Layer Architecture
## The application follows a three-layer architecture to separate concerns and improve maintainability:

- **Presentation Layer:**  Contains the controllers and views to handle user interaction.
- **Business Logic Layer:** (BLL): Contains the core logic of the application (e.g., Excel processing logic).
- **Data Access Layer (DAL):** Manages data access and storage, though in this simple application, this might just involve processing the Excel files.
  
### 3. Library EPPlus
## The project uses the EPPlus library to manage Excel files. EPPlus is a powerful library that simplifies reading from and writing to Excel files.
