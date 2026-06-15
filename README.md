# Student Course Management System

A menu-driven Python application that manages students, courses, and enrollments using SQLite and generates Excel reports using OpenPyXL.

## Features

* Add Student
* View Students
* Add Course
* View Courses
* Enroll Student into Course
* View Student Enrollments using SQL JOIN
* Update Student Marks
* Delete Enrollment
* Show Top Performing Student
* Show Course Average Marks
* Generate Excel Report

## Technologies Used

* Python
* SQLite3
* OpenPyXL

## Database Tables

### Students

* student_id (Primary Key)
* student_name
* age
* email

### Courses

* course_id (Primary Key)
* course_name
* course_duration

### Enrollments

* enrollment_id (Primary Key)
* student_id (Foreign Key)
* course_id (Foreign Key)
* marks

## Excel Report

The application generates `student_course_report.xlsx` containing:

1. Students Sheet
2. Courses Sheet
3. Enrollments Sheet
4. Summary Sheet

The Summary Sheet includes:

* Total Students
* Total Courses
* Total Enrollments
* Top Student
* Highest Marks
* Lowest Marks
* Average Marks

## Concepts Covered

* CRUD Operations
* SQLite Database Management
* SQL Queries and JOINs
* Aggregation Functions (COUNT, AVG)
* File Handling
* Excel Automation with OpenPyXL
* Loops and Exception Handling
* Data Traversal and Searching

## How to Run

1. Install Python.
2. Install OpenPyXL:

   ```
   pip install openpyxl
   ```
3. Run the program:

   ```
   python student_management.py
   ```

## Files Generated

* `student_management.db` – SQLite database
* `student_course_report.xlsx` – Excel report
