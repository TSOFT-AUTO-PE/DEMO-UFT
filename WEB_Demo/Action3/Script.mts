Option Explicit

Dim name, lastname, creditnumber,Num_Iter

name = DataTable("i_Name", dtLocalsheet)
lastname = DataTable("i_LastName", dtLocalsheet)
creditnumber = DataTable("i_Creditnumber", dtLocalsheet)
Num_Iter = Environment.Value("ActionIteration")

Sub EnterData()
	Browser("Book a Flight: Mercury").Page("Book a Flight: Mercury").WebEdit("passFirst0").Set name
	Browser("Book a Flight: Mercury").Page("Book a Flight: Mercury").WebEdit("passLast0").Set lastname
	Browser("Book a Flight: Mercury").Page("Book a Flight: Mercury").WebEdit("creditnumber").Set creditnumber
	Browser("Book a Flight: Mercury").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"EnterData.png", True
	imagenToWord "Ingreso de datos cliente", RutaEvidencias() &Num_Iter&"_"&"EnterData.png"
	Browser("Book a Flight: Mercury").Page("Book a Flight: Mercury").Image("buyFlights").Click
	If DataTable("i_Navegador", "1_LoginWeb") = "chrome" Then
	Browser("Book a Flight: Mercury").Page("Flight Confirmation: Mercury").WebElement("Total Taxes").Output CheckPoint("Total_Price")
	ElseIf DataTable("i_Navegador", "1_LoginWeb") = "ie" Then
	Browser("Book a Flight: Mercury").Page("Flight Confirmation: Mercury").WebElement("WebTable").Output CheckPoint("Total_Price")
	ElseIf DataTable("i_Navegador", "1_LoginWeb") = "ff" Then
	Browser("Book a Flight: Mercury").Page("Flight Confirmation: Mercury").WebElement("Total Taxes").Output CheckPoint("Total_Price")
	End If
	Browser("Book a Flight: Mercury").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Validacion.png", True
	imagenToWord "Validación de compra", RutaEvidencias() &Num_Iter&"_"&"Validacion.png"
	Browser("Book a Flight: Mercury").Page("Flight Confirmation: Mercury").Image("backtoflights").Click @@ script infofile_;_ZIP::ssf2.xml_;_
End Sub

Call EnterData()

