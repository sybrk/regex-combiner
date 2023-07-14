# regex-combiner
Regex combiner for Trados regex files
1.      Drag and drop your Trados regex files (.sdlqasettings) into an excel spreadsheet. Ignore the warning and click ‘Yes’. Choose the first option (As an XML table) and click ‘OK’. Ignore the schema warning and click ‘OK’ again.
2.      Repeat above step for the regex files you want to merge.
3.      Open regex combiner and enable macro features if warned.
4.      Copy the content under below columns in your separate regex files you opened in excel before and paste them under the same columns in the regex combiner.
5.      If there are more than 66 regex rules, you need to increase the number on the ID column like below as needed. The ID should start with RegExRules and increase from RegExRules0. Do not add or delete any row in ID column or it will create problems. 
![image](https://github.com/sybrk/regex-combiner/assets/18572636/73d3c604-36bb-4f60-afb7-4d7ddecdf862)
6.      Finally, click on the Combine button at top of of the excel file and you fill find your combined regex file under the same folder of this excel file. It’s name will be combined.sdlqasettings. You can import this new file in Studio QA settings under regex profiles.

