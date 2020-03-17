# RPA-ExceptionLogging
## Introduction

This is a quick tip on how to build dictionary from data in excel file. Your excel file can contain multiple dictionaries and you can dynamically create required dictionary object by calling the same utility workflow. 

This workflow takes in two input parameters ( config excel file location and dictionary name ) and gives a dictionary output object.

The workflow reads the desired tab from excel sheet and builds the datatable. Then it finds the start and end index of dictionary elements using LINQ based on name of the dictionary provided. 


Start Index will be the name of dictionary in value column.


<code>
startIndex = (From row In dtDictionaryData.AsEnumerable() Let r = row.Field(Of String)("Value") Where r = in_strDictionaryName Select dtDictionaryData.Rows.IndexOf(row)).First() 
</code>

First Skip elements upto start index as you want to find EndDictionary element after the name of dictionary appear in your excel file. End Index will be the "EndDictionary" element in Name column.

<code>
endIndex = (From row In dtDictionaryData.AsEnumerable().Skip(startIndex) Let r = row.Field(Of String)("Name") Where r = "EndDictionary" Select dtDictionaryData.Rows.IndexOf(row)).First()
</code>

Then it selects the data  between those two indexes and builds the dictionary from it. 






