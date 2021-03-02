# Filepath Validator
A tool for mass validating large list of filepaths.

Using this tool you can read an Excel spreadsheet that has filepaths in a specific column. Thousands of rows can be processed with this tool without using much of your systems RAM.

The tool utilises a SAX approach for reading the spreadsheet. Using the SAX approach, OpenXMLReader reads the XML in the file one element at a time, without having to load the entire file into memory.

https://docs.microsoft.com/en-us/office/open-xml/how-to-parse-and-read-a-large-spreadsheet

![img1](https://i.ibb.co/YLfCDV7/Screenshot-1.png)

Choose a spreadsheet using the application UI and specify the column in which the filepaths are listed. Progress is displayed while the filepaths are being checked.

![img2](https://i.ibb.co/5L0bs9J/Screenshot-2.png)

The result will be displayed in the UI and a log file will be created. All the filepaths where a file couldn't be found are listed in the log file.

![img3](https://i.ibb.co/XypGBDH/Screenshot-3.png)

![img4](https://i.ibb.co/pz2hSGy/Screenshot-4.png)
