# ENG101Sort

This program takes a CSV input file of student data and creates a PDF where each student is on an individual page.

How to use:

The CSV file will have one row of headers and then data for each student on a single line.
The data should be in the following order:

1. First Name
2. Last Name
3. Cell phone (unused)
4. Duke email
5. Duke ID (unused)
6. Gender
7. Gender Description (unused)
8. Ethnicity
9. More than 2 yrs in High School
10. AP Credits
11. Electronics background
12. Crafting background
13. Programming background
14. CAD background
15. Rapid Prototyping background
16. Project choice #1
17. Choice #2
18. Choice #3
19. Choice #4
20. Choice #5

Go into the code and modify the code of line `string filename=@"path/file"` and put the correct `path/file` into the string.

Then, run the program within the IDE.  It will create a PDF with the same filename, but ".pdf" added, in the same location.

If the CSV has a different order for the data, either reorder the CSV, or modify the `Students.InputFromCSVValues()` method
such that the indices of the `values` array matches the correct columns.

If there is more than one header line, add additional `ReadLine` commands at the start of the file read.

