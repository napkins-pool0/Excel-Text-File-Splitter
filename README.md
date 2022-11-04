# Excel-Text-File-Splitter
 Excel based tool to split flat file DB extracts to table derived text exports.

# Description

During the development of a personal project, learning SQL, I came up against the issue of cleaning data. The source I wanted to use to learn SQL was the IRS' [Political Organization Filing and Disclosure](https://www.irs.gov/charities-non-profits/political-organizations/political-organization-filing-and-disclosurer) dataset, which is released weekly (I think).

The process of having to split line-by-line the flat data file into multiple data files was time consuming and I figured I would learn some VBA to try and automate what I could.

# Installation
The file is there for download. I have also extracted the .cls and .bas files for visibility

# Instructions
1. ***Table Names and Prefix***

The user enters the table names that the file is split into, and the table delimiter that each line begins with. i.e., Pol Org Disclosure dataset identifies a line belonging to the Political Organisation Reporting table with "2". Then with a pipe to then indicate that is the end of the identifier, "2|"

3. ***Carriage Returns***

Set if you wish to clear carriage returns or not, either "Yes" or "No". I removed carriage returns as sometimes when cleaning the data, Power Query would insist on reading them as new lines, instead of advising them as the same cell. 

4. ***Cycle Size***

This is how many lines it loops through, storing in arrays and then writes to. I found 1,000 to be a sweet spot, but others might have better runs with it on different cycle sizes. 

5. ***Start***

Press Start to choose whether to perform the task on a file or folder. The output files will be created in the same path as the source file or folder. There is a simple % indication on file progress, and if the source is a file, the file count will indicate progress as well. 

# How it works
I've attempted to break down the workings of it, by function, below, as splitting numerous files in a folder with no carriage returns

![Folder No CR Diagram](https://i.imgur.com/yfgOHTS.png)

# Feedback
Any feedback, please provide comments or raise issues, I'm entirely new to all of this.
