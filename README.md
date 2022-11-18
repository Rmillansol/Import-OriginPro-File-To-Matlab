# Import-OriginPro-File-To-Matlab
Function that returns a structure with the books and sheets of the specified OriginPro document.
By means of a single function the code allows importing data from an OriginPro file ( *.opj, *opju). The returned variable is a tree structure with the books and their sheets. With the parameters associated to the books, sheets and columns.
Note that it only retrieves the information of the books and it does not retrieve information of graphs, etc.

# Getting Started
Ensure the code in this project is on your MATLAB path and you have Origin installed. Note only works on Windows. It has been tested with OriginPro 2020 and 2022.
```matlab
outData=ImportOriginFileToMatlabData(filepath);
```
Where outData is the structure with all the books of the Origin file.
For example, to access data on sheet *i* in workbook *j*. The data will be a cell array of *n* rows for *m* columns.
```matlab
cellarray=outData.books(i).Sheets(j).Data;
```
Then, you can access column properties such as name, long name, units or column type by
```matlab
LongName=outData.books(i).Sheets(j).Columns(k).LongName;
```
Remember that you can query the structure using the variable viewer provided by the Matlab IDE or by using the instruction.
```matlab
fieldnames(struct array)
```
# Important Information
This code is licensed under the MIT License.

The author of this code has no relationship with MATLAB or Origin.

When the code is started it will connect to an open instance of Origin. 
