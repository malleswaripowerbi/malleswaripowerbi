ibutor
26 lines (26 sloc)  3.32 KB
   
let
    Source = Excel.Workbook(File.Contents("C:\Users\USER\Downloads\EmployeeData.xlsx"), null, true),
    EmployeeData1_Sheet = Source{[Item="EmployeeData1",Kind="Sheet"]}[Data],
    #"Promoted Headers" = Table.PromoteHeaders(EmployeeData1_Sheet, [PromoteAllScalars=true]),
    #"Changed Type" = Table.TransformColumnTypes(#"Promoted Headers",{{"EmployeeID", Int64.Type}, {"NationalIDNumber", Int64.Type}, {"ContactID", Int64.Type}, {"LoginID", type text}, {"ManagerID", Int64.Type}, {"Title", type text}, {"BirthDate", type datetime}, {"MaritalStatus", type text}, {"Gender", type text}, {"HireDate", type datetime}, {"Dept", type text}, {"Salary", Int64.Type}, {"Job Grade", type text}, {"CurrentFlag", Int64.Type}, {"rowguid", type text}}),
    #"Removed National id" = Table.RemoveColumns(#"Changed Type",{"NationalIDNumber"}),
    #"title to full name" = Table.RenameColumns(#"Removed National id",{{"Title", "full name"}}),
    #"Changed birth date" = Table.TransformColumnTypes(#"title to full name",{{"BirthDate", type date}}),
    #"Replaced m to married" = Table.ReplaceValue(#"Changed birth date","M","Married",Replacer.ReplaceText,{"MaritalStatus"}),
    #"Replaced marial status s to single" = Table.ReplaceValue(#"Replaced m to married","S","Single",Replacer.ReplaceText,{"MaritalStatus"}),
    #"Replaced Gender m to male" = Table.ReplaceValue(#"Replaced marial status s to single","M","Male",Replacer.ReplaceText,{"Gender"}),
    #"Replaced Gender f to female" = Table.ReplaceValue(#"Replaced Gender m to male","F","Female",Replacer.ReplaceText,{"Gender"}),
    #"Changed Hiredate" = Table.TransformColumnTypes(#"Replaced Gender f to female",{{"HireDate", type date}}),
    #"Removed currentflag" = Table.RemoveColumns(#"Changed Hiredate",{"CurrentFlag"}),
    #"Split full name" = Table.SplitColumn(#"Removed currentflag", "full name", Splitter.SplitTextByEachDelimiter({" "}, QuoteStyle.Csv, false), {"full name.1", "full name.2"}),
    #"fullnamde to first name" = Table.RenameColumns(#"Split full name",{{"full name.1", "first name"}, {"full name.2", "Last name"}}),
    #"Removed rowguid" = Table.RemoveColumns(#"fullnamde to first name",{"rowguid", "ContactID"}),
    #"Inserted Age" = Table.AddColumn(#"Removed rowguid", "Age", each Date.From(DateTime.LocalNow()) - [BirthDate], type duration),
    #"Calculated Total Years" = Table.TransformColumns(#"Inserted Age",{{"Age", each Duration.TotalDays(_) / 365, type number}}),
    #"Changed to whole no" = Table.TransformColumnTypes(#"Calculated Total Years",{{"Age", Int64.Type}}),
    #"Removed Managerid" = Table.RemoveColumns(#"Changed to whole no",{"ManagerID"}),
    #"Added Conditional respect name" = Table.AddColumn(#"Removed Managerid", "Respect name", each if [Gender] = "Male" then "Mr" else if [MaritalStatus] = "Single" then "Ms" else if [MaritalStatus] = "Married" then "Mrs" else null),
    #"Merged eid to dept" = Table.CombineColumns(Table.TransformColumnTypes(#"Added Conditional respect name", {{"EmployeeID", type text}}, "en-IN"),{"EmployeeID", "Dept"},Combiner.CombineTextByDelimiter("-->", QuoteStyle.None),"empid & dept"),
    #"Reordered Columns" = Table.ReorderColumns(#"Merged eid to dept",{"empid & dept", "LoginID", "first name", "Last name", "BirthDate", "MaritalStatus", "Gender", "HireDate", "Salary", "Job Grade", "Age", "Respect name"})
in
    #"Reordered Columns"
