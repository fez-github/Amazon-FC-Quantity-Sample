# Amazon-FC-Quantity-Sample
This is a sample file showcasing my work with Excel and VBA.  
While working with Amazon, I had to use their Daily Inventory Report to cross-reference their FBA inventory count with our database count.
I had encountered a problem in which the report would not list if a certain item's quantity had dropped to 0.

	Ex. On Day 1, SKU1 would have 10 quantity in warehouse NCI3, and 5 quantity in KQO8.  On day 2, SKU1 in NCI3 would have 8 quantity, and KQO8 would have 0.
		-Day 1's report would show both warehouses's quantity just fine.  However, day 2 would only return a record of NCI3's quantity, and no record of KQO8.
		-No record would show, which meant a simple import of the report would not update the quantity field in our database to 0.
		
I had to keep an Excel file which would store all the active records, and detect when a certain item record was not imported in again.

Files Used:
Amazon FC Quantity Sheet[Base file with all code.  Included as FC Quantity Test.xlsb]
Daily Inventory Reports[The sheets included utilize dummy data for testing.  Included as FCQuantity Report 1 & 2]
Parent Item & SKU Name[Sheet pulled from database.  Used to link Amazon SKU name with database SKU name.  Included as Parent Name Reference Book.]

Userform:
Originally I had planned to generate a VBA Userform to allow users to activate the code.  However, the form turned out to be cumbersome to use for something as simple as hitting a button.
Instead, I created a sheet with 3 button controls that the user can click.

Exported Files:
2 CSV files will be saved.  
	One contains all of the base report records with additional data used to link them to our database records.
	Another uses summarized data that tells us how many of each SKU is in Amazon Warehouses.
		-This sheet is included as a sample.
