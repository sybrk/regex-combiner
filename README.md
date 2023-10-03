# regex-combiner
Regex combiner for Trados regex files

# Changes in version 1.7:

ID column has been hidden. Regex IDs will be autosorted by macro now. Users do not need to worry about ID sorting.

Users will be warned if Description field has no value or duplicate values. regex file will not be created until these issues are fixed. This is to help avoid issues in Trados.

Users can use drop-down to choose the correct values for IgnoreCase and RuleCondition columns.

## Guide for version 1.7:

Drag and drop your Trados regex files (.sdlqasettings) into an excel spreadsheet. Ignore the warning and click ‘Yes’. Choose the first option (As an XML table) and click ‘OK’. Ignore the schema warning and click ‘OK’ again.

Repeat above step for the regex files you want to merge.

Open regex combiner.xlsm file and enable macro features if warned.

Copy the relevant content in your separate regex files you opened in excel before and paste them under the same columns in the regex combiner.

You can delete rows from the table except the one on row 6 saying "Insert your regexes below this row. Leave this row as it is." You can delete the entire row in excel if you cannot delete a table row. Some cells in the rows 5 and 6 are locked to prevent issues.

Click on the Combine button at top of of the excel file. If you have emtpy cells under Description field or have duplicate descriptions, you will be warned and your file will not be created. If all is ok you will see a message saying "Regex file created".

You fill find your combined regex file under the same folder of this excel file. It’s name will be combined.sdlqasettings. You can import this new file in Studio QA settings under regex profiles.

## Guide for version 1.6

Drag and drop your Trados regex files (.sdlqasettings) into an excel spreadsheet. Ignore the warning and click ‘Yes’. Choose the first option (As an XML table) and click ‘OK’. Ignore the schema warning and click ‘OK’ again.

Repeat above step for the regex files you want to merge.

Open regex combiner.xlsm file and enable macro features if warned.

Copy the content under below columns in your separate regex files you opened in excel before and paste them under the same columns in the regex combiner.

If there are more than 66 regex rules, you need to increase the number on the ID column like below as needed. The ID should start with RegExRules and increase from RegExRules0. Do not add or delete any row in ID column or it will create problems. 

![image](https://github.com/sybrk/regex-combiner/assets/18572636/73d3c604-36bb-4f60-afb7-4d7ddecdf862)

Finally, click on the Combine button at top of of the excel file and you fill find your combined regex file under the same folder of this excel file. It’s name will be combined.sdlqasettings. You can import this new file in Studio QA settings under regex profiles.
