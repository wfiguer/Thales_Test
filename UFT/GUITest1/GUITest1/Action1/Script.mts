'Starts the execution, Opens the page and searchs the location
Browser("thales colombia - Google").Page("thales colombia - Google").WebEdit("search").Set "thales colombia" @@ script infofile_;_ZIP::ssf1.xml_;_
Browser("thales colombia - Google").Page("thales colombia - Google").WebButton("Search").Click

'Checks the WebElements with the DataTable information

Browser("thales colombia - Google").Page("Thales International Suc.").WebElement("addressWebElement").CheckProperty "innertext", DataTable.GetSheet("Main").GetParameter("Address").ValueByRow(1)
Browser("thales colombia - Google").Page("Thales International Suc.").WebElement("urlWebElement").CheckProperty "innertext", fixUrl(DataTable.GetSheet("Main").GetParameter("URL").ValueByRow(1))
Browser("thales colombia - Google").Page("Thales International Suc.").WebElement("contactPhoneWebElement").CheckProperty "innertext", fixContactPhone(DataTable.GetSheet("Main").GetParameter("Contact phone").ValueByRow(1))

'Creates information sheet about routes
DataTable.AddSheet ("RoutesResult").AddParameter "Time",""
DataTable.AddSheet ("RoutesResult").AddParameter "Via",""

' Creates information about better option of route
DataTable.AddSheet("BetterOption").AddParameter"Starting point",""
DataTable.AddSheet("BetterOption").AddParameter"Time",""
DataTable.AddSheet("BetterOption").AddParameter"Via",""

'Runs 5 times (5 rows with differents Starting points)
RunAction "DifferentRoutes", 1

'Notifies the better option
MsgBox "The better option is from: " & DataTable.GetSheet("BetterOption").GetParameter("Starting point").ValueByRow(1) & " With " & DataTable.GetSheet("BetterOption").GetParameter("Time").ValueByRow(1) & " min, with the route: "&DataTable.GetSheet("BetterOption").GetParameter("Via").ValueByRow(1),"0","BETTER  OPTION"


