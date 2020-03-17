# RPA-ExceptionLogging
## Introduction

This is a quick tip on how to build dictionary from data in excel file. Your excel file can contain multiple dictionaries and you can dynamically create required dictionary object by calling the same utility workflow. 

This workflow takes in two input parameters ( config excel file location and dictionary name ) and gives a dictionary output object.

The workflow reads the desired tab from excel sheet and builds the datatable. Then it finds the start and end index of dictionary elements using LINQ based on name of the dictionary provided. 
<code>
startIndex = (From row In dtDictionaryData.AsEnumerable() Let r = row.Field(Of String)("Value") Where r = in_strDictionaryName Select dtDictionaryData.Rows.IndexOf(row)).First() 

endIndex = (From row In dtDictionaryData.AsEnumerable().Skip(startIndex) Let r = row.Field(Of String)("Name") Where r = "EndDictionary" Select dtDictionaryData.Rows.IndexOf(row)).First()
</code>

Then it selects the data  between those two indexes and builds the dictionary from it. 






