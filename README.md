# Personal toolkit
A repository to store some of my personal random tools.

## Tools present as of today (see commit date)
- view_object_info.py:<br>
Requires pandas (`pip3 install pandas`).<br>
A Python script that will print the methods and/or attributes of a given object.<br>
To use it, import the `display_info` function and call it passing to it and object and optionally what you want to display from it (either `methods`, `attributes`, or `all`). Default is `all`. If you want to avoid the message that defines what info is going to be printed, you can just import any of the individual functions and work directly with them.

- docx_highlight_diff.py<br>
Requires python-docx (`pip3 install python-docx`).<br>
A Python script that receives two input Microsoft Word files (need to have `.docx` extension) and an output filename, and creates a Microsoft Word file with the given name, where its contents are exactly the same as the second input file, but it will have highlighted in yellow all the elements present in the second input file that are not present in the first one. 
