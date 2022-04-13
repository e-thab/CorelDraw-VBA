# Auto Import

Actually not VBA, but AutoIt. A BASIC-like scripting language for general Windows automation.
Thought I'd put it here since it's another work program.

PDFs are placed into the RIP software as needed to print on sublimation paper or vinyl. Since they go from the printer to the CNC machine to be cut out,
they need registration marks for the CNC to recognize.

Normally, the process of placing PDFs into the RIP software is very slow and requires tedious file browsing and repitition. This program automates the majority of that.

All that’s needed is to select the product name, enter the SKU, and go. A string is procedurally built to determine the file path and input is sent to window controls in the RIP software to automatically import the file.

With logic for quantity! For example: The PDF for one of the products is naturally 4-up since we do packs in multiples of 4 and that’s what fits best on a sheet. Entering a quantity of ‘16’ imports 4 files. Not super complex but pretty nifty.

# How to Implement

I'm not sure there's a reasonable way on this one. The entire process of building the string that the program is based on requires hard-coded references to folder hierarchy.
And it requires that you're using a specific version of ErgoSoft RIP software. If you do want to try, just install AutoIt and edit or compile the script.

You're welcome to take a look at the code and use whatever you like to Frankenstein something or just use as a reference.