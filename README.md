As a systems analyst I often need to quickly confirm and/or assess the effects of various Essbase data operations I perform, such as loading and calculating new data, or moving existing data between different scenarios and versions. To facilitate that, I created a process that uses Excel VBA to do the following:

• Loops through each of the FP&A department's Essbase cubes and captures multi-level "Before" and "After" time-stamped data snapshots 

• The "After" sheets are analyzed by a second VBA process that compares their values with the corresponding "Before" values and applies a yellow background to all variant cells

• Most of the sheets are hundreds of rows deep and dozens of columns wide, so the VBA also compiles a list of columns containing at least one yellow cell, and writes that list to the upper left corner of each "After" sheet

This gives me a quick and powerful visual summary of which cubes changed, exactly where they changed, and to what degree.
