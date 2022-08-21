'Stores time and via of each point to Thales in a DataTable (RoutesResult)
'Finds the route for each starting point
Browser("Thales International Suc.").Page("Thales International Suc.").WebButton("directionWebButton").Click
Browser("Thales International Suc.").Page("Google Maps").WebEdit("choosePointWebEdit").Set DataTable.Value("Starting Point","DifferentRoutes")
Browser("Thales International Suc.").Page("Google Maps").WebButton("searchWebButton").Click

'Store information about each route
rowCount = DataTable.GetSheet ("RoutesResult").GetRowCount
DataTable.SetCurrentRow(rowCount+1)
DataTable.Value("Via","RoutesResult")=Browser("Thales International Suc.").Page("Google Maps").WebElement("viaLabel").GetROProperty ("innertext")
DataTable.Value("Time","RoutesResult")=  changeToMin(Browser("Thales International Suc.").Page("Google Maps").WebElement("timeLabel").GetROProperty ("innertext"))

'Store information about better route
rowCount2 = DataTable.GetSheet("BetterOption").GetRowCount
DataTable.SetCurrentRow(rowCount2)
If DataTable.Value("Time","BetterOption") > changeToMin(Browser("Thales International Suc.").Page("Google Maps").WebElement("timeLabel").GetROProperty ("innertext")) or DataTable.Value("Time","BetterOption")=""Then
	DataTable.Value("Starting point","BetterOption")=DataTable.GetSheet("DifferentRoutes").GetParameter("Starting Point").ValueByRow(rowCount+1)
	DataTable.Value("Time","BetterOption")=changeToMin(Browser("Thales International Suc.").Page("Google Maps").WebElement("timeLabel").GetROProperty ("innertext"))
	DataTable.Value("Via","BetterOption") = Browser("Thales International Suc.").Page("Google Maps").WebElement("viaLabel").GetROProperty ("innertext")
End If

Browser("Thales International Suc.").Page("Google Maps").WebButton("closeDirectionsButton").Click


