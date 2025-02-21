#### Project 2
# Formatting Grades in Excel
- This is a group project. Through GitHub Classroom you should have created a shared GitHub repository with your group members, so as long as you upload your finished code to that group repository, each member in your team will get credit. You will also need to fill out a peer review survey on Learning Suite to recieve credit.
- I do not provide automated tests for projects. You will need to determine yourself whether the code meets the requirements provided in the rubric. After you turn in your code, your code will be manually graded (meaning partial credit may be given for certain requirements). The TAs will update the `Rubric.md` file with your grade and any comments that they have.

- Assume a high school teacher approached your group complaining about the Excel file their grading system produces each quarter. It spits out all the classes they teach with all their students in a single worksheet, with student information stored in a single column. The teacher wants your group to make a program that will automatically format and summarize the important information about each of the classes they teach.

- Place your code in the `p02_formatting_grades_in_excel.py` file. Optionally, you can create any new .py files you want to place functions or classes in. If you do so, just make sure you import those files into `p02_formatting_grades_in_excel.py`

## Libraries Required:
- openpyxl
    - `import openpyxl`
    - for fonts: `from openpyxl.styles import Font`

## External Files Required:
Your GitHub repository should contain 2 excel files: `poorly_organized_data_1.xlsx` and `poorly_organized_data_2.xlsx`. Each file contains the following columns:
- **Class Name**: The class the student is in.
- **Student Info**: Contains last name, first name, and student ID in a single column, separated by underscores.
- **Grade**: A grade between 0 and 100.


## Functions/Classes Required:
- There are no specific custom functions or classes that you need to write. You could write some if you think it would make your code nicer, but no points will be gained or lost either way.

## Requirements:
Using the openpyxl library, import one of the two example Excel files. Your program should be robust enough to work with either of the files (and any other file that is structured in the same way as those two files).
- In other words, you shouldn’t be hardcoding in specific names of classes (History, Calculus), etc because your solution should work no matter what classes, students and grades are in the file. 
- When first writing your program, you can just choose one Excel file to work with, and then test it with the other Excel file afterwards.

Your program should create a new Excel file (that way you’ll still have the original Excel file and a new Excel file at the end).

**REQUIREMENTS**: Your program will need to:
1.	Create new worksheets for each class (e.g., a sheet for Algebra, a sheet for Calculus, etc.)
2.	Create columns for last name, first name, student ID, and grade with the student data for that class placed in each new worksheet created.
3.	Place a filter over the 4 aforementioned columns in each sheet.
4.	Create simple summary information on each class's worksheet using functions. These will be placed in columns F (the titles) and G (the data). It should show:
    - The highest grade
    - The lowest grade
    - The mean grade
    - The median grade
    - The number of students in the class
5.	Bold the headers and change the width of the columns on each worksheet.
    - The width of the columns for A,B,C,D,F,G must each be set to the number of characters in the header + 5. 
        - For example the column D header is “Grade” which has 5 characters, so the width of column D should be 10, etc.
6.	Save the results as a new Excel file named “formatted_grades.xlsx”

See the `example_output_1.xlsx` and `example_output_1.xlsx` files for examples of what your files should look like when you're done. Remember also, that if you view any Excel files within VS Code using the Excel Viewer extension it won't show all formatting (it also has trouble displaying the median function). Because of this, make sure you look at the files using Excel to verify you did everything correctly before you push up your code.

## Hints
There is a lot of variation in how exactly your group could perform this, so there isn’t one specific “logical flow” for the project. All that matters is that you create a program that fulfills the 6 requirements for any Excel file in the same format as the 2 starter Excel files. This project is actually great practice for situations like this where you know what you’re starting with and what the end product should be, but you have to plan out the process of getting from A to B.

However, here are some hints that might help you implement each requirement. What you use is up to you:

1. ### Sheets for each class:
- Potentially useful methods / attributes
    - workbook_object.create_sheet()
        - Creates a new worksheet in a workbook.
    - workbook_object.sheetnames
        - Gives you the names of all the sheets in a workbook.
    - workbook_object.worksheets
        - Gives you all the worksheet objects in a workbook.
    - worksheet_object.iter_rows()
        - Useful if you want to loop through a worksheet.
2. ### last name, first name, student ID, and grade columns
- Potentially useful methods:
    - string_variable.split()
        - splits up a string by a character you specify and returns a list of strings.
    - worksheet_object.append()
        - If you give it a list, it will place each element in the list in the next empty row of the worksheet.
3. ### Filter
- The textbook and class practice show an example of applying a filter
- You need to apply the filter to the range starting in A1 and ending in D(the max number of rows in that sheet). How do you get the max number of rows that have data in a sheet?
4. ### Adding functions
- The textbook and class practice show an example of adding in functions.
- Column F will have the titles of the functions
- Column G will have the actual results
5. ### Simple formatting
- You only need to bold the headers of Columns A, B, C, D, F, and G.
- You need to adjust the width of those same columns based on the number of characters in each of the headers. What function returns the number of letters in a string?
- Remember that viewing the Excel file in VS Code might not show the formatting correctly. To make sure you did it right, open up the .xlsx file in Excel.
6.	Save the results
- worksheet_object.save()

## Grading Rubric:
See Rubric.md. Remember to right click and select "Open Preview" to view the file in a nice format. The TAs will update this file with your grade when they are done grading your submission.


## Example Output
See the `example_output_1.xlsx` and `example_output_1.xlsx` files
