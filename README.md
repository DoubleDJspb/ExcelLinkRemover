
# Introduction #
This project provides a C# / .Net console application for removing all external links from Excel .xlsx files. The app does not use Office components and libraries.

## Downloads ##

The latest release of ExcelLinkRemover is v1.0.3, released 12-10-2020. 

* [Source](https://github.com/DoubleDJspb/ExcelLinkRemover/archive/main.zip)
* [Binaries](https://github.com/DoubleDJspb/ExcelLinkRemover/releases/download/v1.0.3/ExcelLinkRemover-1.0.3-bin.zip)

## Overview ##
The project is provided as a .Net 4.7.2 assembly, but can just as easily be changed to .Net 4.0 or later. The solution was created in Visual Studio 2019.

## Usage ##

#### From Explorer ####
Just drag and drop the required .xlsx files to the **ExcelLinkRemover.exe** program file to remove links.

#### From Console ####
Alternative launch from the command line: **ExcelLinkRemover.exe path_to_file_1.xlsx path_to_file_2.xlsx path_to_file_N.xlsx**

The original files will not be modified. The program will save changes in the ExcelLinkRemover subfolder in new files with the same names.

#### Important! ####
The user must have **write permissions** in the folder where the source files are in order to be able to create new files.
