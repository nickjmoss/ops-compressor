# Ops Compressor

This tool was created to help in the compression and decompression(explosion) of certain Excel
spreadsheets. This tool was requested to be created by Travis Mecham and will be for him and
others on his team that will be manipulating specific Excel spreadsheets of records.


## Authors

- [@nickjmoss](https://www.github.com/nickjmoss) Nick Moss (Data Quality Intern)


## How It Works

This tool uses the Java Library [Apache POI](https://poi.apache.org/) which can read and write to
Excel files. It has an API that simulates an Excel file with sheets, columns, rows, and cells. This
library reads the data from an Excel file into its API so that the data can be manipulated and then
written to a new Excel file.

This tool also makes use of the Java Library [Swing](https://www.javatpoint.com/java-swing) to 
create a very simple GUI for the user.

### Reading The Excel File
