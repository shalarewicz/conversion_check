Date: 2020-05-11
Author: Stephan Halarewicz

Validates IRBNet Conversion Kit
Conversion Kit tabs should be titled Research Administration, R&D, IRB, IACUC, IBC, Research Safety, Biosafety, Determinations

Checks for common errors in conversion data including
	- Characters not supported by the latin-1 encoding
	- Line breaks
	- Commas in the Internal Reference Number
	- Missing required fields
	- Invalid date formats. Data should still be inspected for visual date format YYYY-MM-DD as Pandas exports all date formatted values in the correct format. 
	- Non-numeric vote format
	- Valid Values based on mapping defined in valid_values_map.xlsx
	  This file defines the types available in IRBNet based on the configured Board Type. These do not strictly match the
	  values defined on the Valid Values tab. Values on the valid values tab may be not be supported by the board type but 
	  also may not cause the conversion to fail and therefore a warning is printed. Furthermore, note that types such as 
	  Designated Review (only allowed for IACUCS on the Valid Values Tab) is allowed for RDC and IBC board types per the 
	  system. 
	- Prints warning for any submissions Pending Review
	- Checks that all projects have been reviewed by at least one board
	- Reasonable Date Checks
		- Expiration and Report Due Dates greater than the Effective Date
		- Expiration and Report Due Dates less than 1 year from today (3 years for Expiration in IACUC)
		- Submission Date is in the past
		- Initial Approval Date before Effective Date
		- Next Report Due not equal to Expiration
		- Effective Date is in the past

Requirements
	pip install -r requirements.txt
	- python 3.8.3
	- pandas
	- valid_values_map.xlsx in root directory

Run with
	py -3 conversion-check.py filename.xlsx > output.txt

Test on
	-"Test Data Conversion Model.xlsx" compare output to gold.txt
		- Refresh date formulas by opening and saving sheet.
	-"Test Data Conversion Model 2.xlsx compare output to gold2.txt

	- Compare txt files with fc <file1> <file2> in cmd


