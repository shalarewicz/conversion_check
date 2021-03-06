Date: 2020-05-11
Author: Stephan Halarewicz

Validates IRBNet Conversion Kit
Conversion Kit tabs should be titled Research Administration, R&D, IRB, IACUC, IBC, Research Safety, Biosafety, Determinations

Checks for common errors in conversion data including
	- Characters not supported by the latin-1 encoding
	- Line breaks
	- Missing required fields
	- Invalid date formats. Data should still be inspected for visual date format YYYY-MM-DD as Pandas exports all 
	  date formatted values in the correct format. 
	- Non-numeric vote format
	- Valid Values based on mapping defined in valid_values_map.xlsx
	  This file defines the types available in IRBNet based on the configured Board Type. These do not strictly match the
	  values defined on the Valid Values tab. Values on the valid values tab may be not be supported by the board type but 
	  also may not cause the conversion to fail and therefore a warning is printed. Furthermore, note that types such as 
	  Designated Review (only allowed for IACUCS on the Valid Values Tab) is allowed for RDC and IBC board types per the 
	  system. 
	- Prints warning for any submissions Pending Review

Requirements
	pip install -r requirements.txt
	- python 3
	- pandas
	- valid_values_map.xlsx in root directory

Run with
	py -3 conversion-check.py <filename>

Test on 
	-"Test Data Conversion Model.xlsx" compare output to gold.txt
	-"Test Data Conversion Model2.xlsx compare output to gold2.txt

	- Compare txt files with fc <file1> <file2> in cmd
	- Out file will need to be encoded with UTF-8


