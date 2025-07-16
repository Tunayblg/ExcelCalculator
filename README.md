AutoSchedU – Automated Course Scheduling System
AutoSchedU is a Django-based academic course scheduling system designed for university-level use. It automates the creation of weekly course timetables, optimizes room usage, prevents conflicts, and simplifies student and instructor enrollment workflows. The project includes an Excel-based batch enrollment system and an intelligent scheduling algorithm that dynamically assigns courses based on availability and constraints.

🔧 Core Features
Automatic Course Scheduling Algorithm

Assigns courses to available time slots without overlapping instructors or students

Takes classroom availability, course types (theory/lab), and capacity constraints into account

Supports fixed day/hour enforcement for certain courses if needed

Excel-Based Student Course Enrollment

Upload Excel files to batch assign students to courses

Automatically creates new student/instructor accounts if not found in the system

Supports multi-department course enrollments

Role-Based User System

Coordinator: Generates timetables, manages departments, and performs admin tasks

Instructor: Views assigned courses and schedules

Student: Views personal schedule and registered courses

Manual Adjustments via Admin Panel

Timetables can be edited manually after automatic generation

Schedule is shown in visual grid format (similar to a classic weekly calendar)

🧠 Scheduling Logic Overview
The scheduling algorithm works by:

Analyzing All Constraints

Instructor availability (unavailable days/hours)

Room types and capacities

Course hour requirements (e.g., 2-hour per week)

Avoiding overlaps between departments and shared courses

Greedy + Backtracking Assignment

Greedy logic tries best-fit placements

If unsatisfiable, the algorithm backtracks and tries alternatives

Ensures balanced distribution of courses across the week

Final Output

Saves the generated timetable to the database

Displays it in a table for each class and department

Excel export available

📂 Project Modules
