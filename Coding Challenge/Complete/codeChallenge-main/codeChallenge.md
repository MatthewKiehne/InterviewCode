# Code Challenge

This challenge is designed for you to show your abilities to read data from different file types, work with different data types, validate the data, format and manipulate data to create reports, and store data. You can use any language(s) to accomplish this task. You will write two separate programs for these tasks.

## Program One Design Overview

1. Read data from either `shippingDetails.xlsx or shippingDetails.csv`, your choice.
2. Validate Data.
3. Generate reports.

### Validate Data and Generate Error Reports

1. Your program should validate email addresses and phone numbers.

> Make sure that the email field for all orders is not blank, contains the `@` symbol, and contains one of the following suffixes: `.com, .net, .edu, .org`. A report should be generated with a title of `Invalid Email Addresses` and list each OrderNum and email address that does not comply on its own line. Print the results to the console.
> Make sure that all phone numbers are `not blank, contain 9 digits and two -'s, contain only numbers or -, and are formatted as 123-456-7890`. Generate a report titled `Invalid Phone Numbers` and list each Order Num and Phone Number that does not comply on its own line. Print the results to the console.

### Reports

1. Generate a report for fulfillment that shows Order Num, OrderDate, and ShipDate where the OrderDate and ShipDate are `less than 24 hours apart`. Your program should save this to an external file of your choice.
2. Generate a report for marketing that shows Order Num, OrderTotal, and Email for orders `over a user specified dollar amount`. The user should be prompted for the dollar amount that the program will then use as a variable. Make sure to validate input. Your program should save this to an external file of your choice.
3. Generate a report that sums and shows the number of orders and the formatted total dollar value(e.g. $1,234.56). Your program should save this to an external file that prepends today's date to the file name(e.g. `10062020-ordertotal.txt`).

## Program Two Design Overview

1. Read data from `products.json`.
2. Generate reports to the console.

### Generate Reports Using User Input

1. Prompt the user for which report they want to generate, `availability`, `multiple variants`, or `price`.
2. If the user selects `availability`, generate a report with two columns, listing all `available` product sku's in a column called `Available`, and `disabled` sku's in an `Unavailable` column.
3. If the user selects `multiple variants`, run a report that prints a column header titled `Multiple Variants` and each product sku that has more than one variant.
4. If the user selects `price`, prompt the user for a `price`, and then prompt for `greater than` or `less than`. Print a report that lists all product `sku's` that meet the criteria of the user selected choices with a heading of `Filtered by Price`.

#### Instructions

1. Include a readme file that explains how to run your programs

#### Source Control

1. Please upload your project to Github

#### Questions

Any questions can be directed to [techteam@serenityhealth.com]

#### Email

Email a link to your project repository when complete to [techteam@serenityhealth.com]
