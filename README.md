# excel-product-file-maker
A selection of modules and procedures in excel VBA to generate multiple 3rd party data file types from one source excel file

**DISCLAIMER**
I am very new to programming as a whole, everything I write is a learning experience for myself and I am certain that there will be many mistakes and many "You are doing everything wrong!!" moments. I am always looking for an honest opinion or a better way to do something.

What I am trying to accomplish

Amazon, Overstock, Wayfair and all the 3rd party marketplaces that I have experience with rely heavily on creating excel files that contain product information and order information.

The column order is different for each 3rd party platform but the source file information always remains the same (product or otherwise). Instead of having to copy and paste multiple times to get the same information just into a different column, I would like to run a macro or routine to do so.

Basically

1) The ability to assign column names to variables based on the value of the cell
2) The ability take said values and create new excel files according to a different header order then from source file
3) Perform calculations that depend on user input when the source file is opened
4) Create a new excel file specific to the 3rd party marketplace with as much informattion as possible.

