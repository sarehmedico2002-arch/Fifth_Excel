# Fifth_Excel
# Car Inventory Analysis

An Excel workbook built from our original sales data, which started as a plain text file and was converted into a structured spreadsheet for analysis.

## File

- `car_inventory.xlsx`

## What it does

- **Text-to-data conversion**: source data started as a text file and was brought into Excel as the base for this workbook
- **LEFT / MID / RIGHT**: used to break each Car ID code apart into its Make, Model, and Manufacture Year components
- **VLOOKUP**: matches the extracted Make and Model codes against reference tables to pull in their full names
- **CONCATENATE**: builds a new, reformatted Car ID from the Make, Year, Model, Color, and original ID
- **UPPER**: used inside the new Car ID formula to standardize letter casing
- **Age and Miles/Year**: calculated from the extracted manufacture year and recorded mileage
- **Warranty coverage check**: flags whether a car's mileage is still within its warranty limit
- **Pivot Table**: summarizes total miles driven by each driver, with a grand total
- **Scatter Chart**: visualizes the relationship between two of the dataset's numeric fields

## Sheets

- `in`: raw and calculated data for each car
- `Sheet1`: pivot table summary of miles by driver
