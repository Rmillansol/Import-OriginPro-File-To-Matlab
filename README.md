# Import-OriginPro-File-To-Matlab
Function that returns a structure with the books and sheets of the specified OriginPro document.
By means of a single function the code allows importing data from an OriginPro file ( *.opj, *opju). The returned variable is a tree structure with the books and their sheets. With the parameters associated to the books, sheets and columns.
Note that it only retrieves the information of the books and it does not retrieve information of graphs, etc.

# Getting Started
Ensure the code in this project is on your MATLAB path and you have Origin installed. Note only works on Windows. It has been tested with OriginPro 2020 and 2022.
```matlab
outData=ImportOriginFileToMatlabData(filepath);
```
Where OutData is the structure with all the books of the Origin file
# Important Information
This code is licensed under the MIT License.

The author of this code has no relationship with MATLAB or Origin.

When the code is started it will connect to an open instance of Origin. 
