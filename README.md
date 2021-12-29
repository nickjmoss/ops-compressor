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


The interface for the tool looks like this:

![gui](https://github.com/fs-eng/ops-compresser/blob/main/readme-images/gui.jpg)

The user is prompted to first choose an Excel file on their machine that they would like to
compress or decompress(more on that later). Using the JFileChooser API in the Java Swing library,
when the user clicks on 'Choose a File...' that will open a dialog box where they can browse their
machine for the file they need. Once the file is picked, the path of the file will be displayed on
the GUI.

This program only accepts an Excel file with the extension .xlsx.


### Compressing


Below is an example of an Excel file that we want to compress:

![compress](https://github.com/fs-eng/ops-compresser/blob/main/readme-images/decompressed.jpg)

Notice how there are numerous rows that contain the exact same information except for their IMAGE column.
What we want to do is compress all of these identical rows down into just one row that contains the info found
in all of these rows. But because the IMAGE column is different in each row, the compressor tool will
add a FIRST_IMAGE and LAST_IMAGE column to the compressed Excel file and populate those columns with the image number
where the identical rows begin and the image number where the identical rows end.

The tool will do this for each set of identical rows until it reaches the end of the file. The newly compressed Excel file
will then be saved in the same directory as the original file and will have the same name as the original file except the keyword
COMPRESSED will be in the name along with the date the compressed file was created.

The compressed file should look something like this:

![compress](https://github.com/fs-eng/ops-compresser/blob/main/readme-images/compressed.jpg)


### Decompressing(Exploding)


Decompressing or exploding an Excel file is just the opposite of compressing. 

Below is an example of an that same file we just compressed:

![compress](https://github.com/fs-eng/ops-compresser/blob/main/readme-images/compressed.jpg)

Now, to decompress this file the compressor tool will read through each row in the file and copy the values of each
column into a new row. The number of rows made is determined by the number of images between the FIRST_IMAGE and the LAST_IMAGE.
For example, in the first row of the file above, the FIRST_IMAGE value is 6 and the LAST_IMAGE value is 37, the compressor tool will
then iterate from 6 to 37 and with each iteration it will create a new row in the decompressed Excel file. This new row will not contain
the FIRST_IMAGE and LAST_IMAGE columns, instead those will both be replaced with the IMAGE_NBR column and its value will be the number of
the iteration that the tool is on, so if it were the first iteration for the file above, the value would be 6.

The tool will do this for each row in the Excel file and when it is finished it will save the newly decompressed Excel file in the same directory
as the original file and it will have the same name as the original file except the keyword DECOMPRESSED will be in the name along with the date
the decompressed file was created.

The decompressed file should look something like this:

![compress](https://github.com/fs-eng/ops-compresser/blob/main/readme-images/decompressed.jpg)

### Output

If the file is either compressed or decompressed successfully, then the Output pane below the 'Compress' and 'Decompress' buttons will display messaged that
the operation completed successfully and it will show the path of the new file that was created.

If there is an error, an error message will be displayed in the Output pane with possible suggestions as to why something went wrong.

IMPORTANT: This only works with specific Excel files that have a certain order to their columns, these Excel files were provided to me by Travis Mecham and I 
built this tool according to the files he provided me.
