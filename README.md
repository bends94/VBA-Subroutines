# VBA-Subroutines
Some examples of VBA Subs I have created. None of this will probably work for anyone else, as I created them for specific workbook formats.

clear_contents(ws_clear,start_cell)<br>
ws_clear is the name of the worksheet to clear.<br>
start_cell is the cell to start the clearing in. The range goes through column Z and the last row in the column specified.<br>
clear contents just cleares a section of the worksheet for me. I typically use it inside other subroutines so I don't have to type out the full code every time.
<br><br>
TotBoM()<br>
TotBoM totalized the BoM worksheet, into the Totalized BoM worksheet. Adds all quantities of repeatedly recurring parts together.
<br><br>
Product_Cutsheets()<br>
Product_Cutsheets_Save(OpenPath,SavePath,partStr)<br>
Using the Totalized Bill of Material (BoM), create marked up pdf datasheets and install sheets for each device. This is used in tandem with a access database that has part names, and filesystem locations of saved datasheets and install sheets for each part. Uses the Adobe Acrobat API to markup and save files into a directory based on the job number.
