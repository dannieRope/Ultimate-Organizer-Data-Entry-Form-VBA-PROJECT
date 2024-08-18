# Data Entry Form (Ultimate Organizer)- VBA-PROJECT

## Problem Statement

This project has wide-ranging uses and applications. it involves Creating a VBA user form that will:
1) allow the user to add or delete new categories (user-defined) to a data table (these would be the columns/column headings) on a main worksheet,
2) allow the user to add or delete records (a row of a table), and
3) allow the user to look up different categories for a record (basically searching through the data and outputting a specific user-defined category) with the option of replacing those items.

## Requirements

**1)	Main Form.**  The main form should have options to add a category, delete a category, add a new record (row of a table), delete record, and search through the data (by using one of the categories).  A delete confirmation box should appear to confirm deletion of any categories. 
The main form might look something like this:

![image](https://github.com/user-attachments/assets/f324d7f1-60c5-4d40-b1bd-7f80129840c6)

**2)	Add Category.**  If the user wishes to create a new category, a user form similar to the one below should appear.  This allows the user to input the name of a new category.  This would be the heading of a column of data.  When the user selects “Add Category” then the next available (blank) column would be entitled what the user inputs.  As a simple example, some headings might be “Name”, “Phone Number”, and “Address”.

![image](https://github.com/user-attachments/assets/f31946a5-34fa-4287-9174-d461e9f0b6df)

**3)	Delete Category.**  If the user wishes to delete an entire category (other than the Names category), a user form similar to the one below should appear.  All of the columns are populated in a combo box, with the default being the second column.  This allows the user to select the category that they wish to delete (i.e. an entire column of data from the spreadsheet).  This should be confirmed with a Yes/No message box after “Delete” is clicked!

![image](https://github.com/user-attachments/assets/b0df30d8-8004-443d-a6fd-9ab2e59b3ce8)

**4)	Add Record.**  Next, we wish to be able to add records (rows of the table/spreadsheet).  A user form similar to the one below should allow the user to input a new record.  In the example below, I’ve assumed that the user has already created 3 categories (“Name”, “Phone”, and “Address”).  The user form should be able to display up to 12 different categories, so make sure there is extra space in case later on the user adds a new category.  Note that you should have 12 total labels and 12 total text boxes.  If there are currently n categories of data on the worksheet, then only n of the labels and n of the text boxes should be visible; all others should remain hidden.  See the screencasts on how to do this.  In the diagram below, the dotted lines represent that those labels and text boxes are hidden (i.e., not currently being used), but you should allow up to 12 total categories to be used on the spreadsheet.  

![image](https://github.com/user-attachments/assets/34996806-8880-4b70-9f39-10834a306524)

Again, there is a lot of space here such that if a new category is added then those new categories will show up.  When the user submits “Add Record”, the items in the user form are added in a single row of the spreadsheet.

**5)	Delete Record.**  If the user wishes to delete an entire record (row of the spreadsheet, other than the first row), a user form similar to the one below should appear.  All of the names in column A will appear in a combo box.  This allows the user to select the record (name) that they wish to delete (i.e. an entire row of data from the spreadsheet).  This should be confirmed with a Yes/No message box after “Delete” is clicked!   

![image](https://github.com/user-attachments/assets/da00e304-0d0a-4982-8d5d-1cb2209247a1)

**6)	Search/Replace.**  The final aspect of the project is to allow the user to search through records for information and replace or add information.  The user should be able to select one of up to 12 categories (drop-down list) to use as a search criterion.  The user form will then display what the user is searching for and will also allow them to replace the information.  If the user selects to replace an item for a particular row and column, then the change will be permanently made to the worksheet.  If there is no available information for that item, then the user form will ask the user if they would like to add information to that record and category, and the addition should be made permanent on the worksheet. 

![image](https://github.com/user-attachments/assets/51d250f4-ac9f-4e34-af49-755d89e0b2cb)

**7)	Input Validation.** Make sure that your user form never brings up the “Debug” box and the Visual Basic Editor.  Part 2 (Week 4) of the course has some basic information on input validation for user forms.  There is a lot of input validation and fine tuning that you can do for this project (like what happens if Cancel buttons are pressed, etc.).  You won’t be graded on input validation, but do your best to make sure that there aren’t too many errors encountered during normal operation of your project! 

## Creating the form

**1)	Main Form.**

![Capture](https://github.com/user-attachments/assets/bbf39bf0-b643-4094-a3f8-84af9bf9791d)

Find the VBA codes behind this form [here]()

**2)	Add Category.**

![Capture2](https://github.com/user-attachments/assets/74192edd-ffbe-4f49-8646-b718a70614f0)

Find the VBA codes behind this form [here](https://github.com/dannieRope/Ultimate-Organizer-Data-Entry-Form-VBA-PROJECT/blob/main/Forms/AddCategory.frm)


**3)	Delete Category.**

![Capture3](https://github.com/user-attachments/assets/54805c87-3e15-42ec-be3d-03c525e4a027)

Find the VBA codes behind this form [here]()

**4)	Add Record.** 

![Capture4](https://github.com/user-attachments/assets/9e5eb631-92c8-4ed6-bb6b-61d03dd4c629)

Find the VBA codes behind this form [here](https://github.com/dannieRope/Ultimate-Organizer-Data-Entry-Form-VBA-PROJECT/blob/main/Forms/AddRecord.frm)

**5)	Delete Record.**

![Capture5](https://github.com/user-attachments/assets/8466231c-7cd6-4f71-974a-ea58dda2e769)

Find the VBA codes behind this form [here]()

**6)	Search/Replace.**

![Capture6](https://github.com/user-attachments/assets/6b98180c-9406-47ee-b86c-5ffa3049dbd6)

Find the VBA codes behind this form [here]()

## Structure

- **Modules/**: Contains standard modules (.bas files)
- **Forms/**: Contains user forms (.frm files)
- **Workbook/**: Contains the main workbook with macros (.xlsm file)

## Usage

1. Download the repository.
2. Import the necessary `.bas` or `.frm` files into your VBA editor in Excel.
3. Open the `.xlsm` workbook to see the macros in action

## License
[MIT License](LICENSE)


