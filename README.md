<div align='center'>
     
<img src="https://github.com/kbkozlev/Excelize/blob/master/.github/Excelize1.png" alt="logo" width="540" height="211"><br/>

<a href="ttps://github.com/kbkozlev/Excelize/blob/master/LICENSE.md" alt="License">
  <img src="https://img.shields.io/github/license/kbkozlev/Excelize?color=blue&style=for-the-badge" />
</a>

<a href="https://github.com/kbkozlev/Excelize/releases" alt="GitHub release">
  <img src="https://img.shields.io/github/v/release/kbkozlev/Excelize?color=blue&style=for-the-badge" />
</a>
     
</div>

<div>

# Excelize
Excelize is a small app that will allow users to easily manipulate Excel files, by merging, converting and splitting them as needed. 
<br>Multiple file selection is supported in order to optimize the processing of a large number of files.
     
# Download and Run
You can download the latest <a href="https://github.com/kbkozlev/Excelize/releases">release</a> or use the command line to download the source code:
     
```
     curl -LO https://github.com/kbkozlev/Excelize/archive/refs/heads/master.tar.gz
     tar -xf master.tar.gz
     cd Excelize-master
     pip install -r requirements.txt
     python main.py     
```
# How to use
<b>Prerequisites:</b> All files must have same header.

1. Select one or multiple files of the same type.
2. Select output folder
3. Select input file type 
4. If you select worksheets you can combine them in a single file, or you can split the original file into seperate files where each worksheet is a workbook.
   <br>If you select workbooks you can combine multiple workbooks into a single file.
5. The output files can be converted to either xlsx or csv, or both if both checkboxes are selected.
</div>
