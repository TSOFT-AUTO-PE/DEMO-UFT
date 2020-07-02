Option Explicit

Dim num_passengers, from_port, from_month, from_day, to_port, to_month, to_day, ser_class, airline, Num_Iter

		num_passengers = DataTable("i_NumPassengers", dtLocalsheet)
		from_port = DataTable("i_From", dtLocalsheet)
		from_month = DataTable("i_FromMonth", dtLocalsheet)
		from_day = DataTable("i_FromDay", dtLocalsheet)
		to_port = DataTable("i_ToPort", dtLocalsheet)
		to_month = DataTable("i_ToMonth", dtLocalsheet)
		to_day = DataTable("i_ToDay", dtLocalsheet)
		airline = DataTable("i_Airline", dtLocalsheet)
		Num_Iter = Environment.Value("ActionIteration")	
	
Sub FlightDetails()
	
	Select Case DataTable("i_TypePass", dtLocalsheet)
		Case "1"
		Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebRadioGroup("tripType").Select "roundtrip"
		Case "2"
		Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebRadioGroup("tripType").Select "oneway"
	End Select
 @@ script infofile_;_ZIP::ssf1.xml_;_
	Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebList("passCount").Select num_passengers @@ script infofile_;_ZIP::ssf3.xml_;_
	Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebList("fromPort").Select from_port @@ script infofile_;_ZIP::ssf4.xml_;_
	Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebList("fromMonth").Select from_month @@ script infofile_;_ZIP::ssf5.xml_;_
	Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebList("fromDay").Select from_day @@ script infofile_;_ZIP::ssf6.xml_;_
	Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebList("toPort").Select to_port @@ script infofile_;_ZIP::ssf7.xml_;_
	Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebList("toMonth").Select to_month @@ script infofile_;_ZIP::ssf9.xml_;_
	Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebList("toDay").Select to_day @@ script infofile_;_ZIP::ssf10.xml_;_

End Sub
Sub Preferences()
		Select Case DataTable("i_Class", dtLocalsheet)
			Case "Business"
			Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebRadioGroup("servClass").Select "Business" @@ script infofile_;_ZIP::ssf15.xml_;_
			Case "Economy"
			Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebRadioGroup("servClass").Select "Coach" @@ script infofile_;_ZIP::ssf16.xml_;_
			Case "First"
			Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebRadioGroup("servClass").Select "First" @@ script infofile_;_ZIP::ssf17.xml_;_
		End Select
		Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").WebList("airline").Select airline
		Browser("Find a Flight: Mercury").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"FlightFinder.png", True
		imagenToWord "Detalle de Vuelo", RutaEvidencias() &Num_Iter&"_"&"FlightFinder.png"
		Browser("Find a Flight: Mercury").Page("Find a Flight: Mercury").Image("findFlights").Click @@ script infofile_;_ZIP::ssf18.xml_;_
		While Browser("Find a Flight: Mercury").Page("Select a Flight: Mercury").Image("reserveFlights").Exist = False
			wait 1
		Wend
		wait 2		
		Browser("Find a Flight: Mercury").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"ReserveFlight.png", True
		imagenToWord "Reserva de Vuelo", RutaEvidencias() &Num_Iter&"_"&"ReserveFlight.png"
		Browser("Find a Flight: Mercury").Page("Select a Flight: Mercury").Image("reserveFlights").Click @@ hightlight id_;_Browser("Find a Flight: Mercury").Page("Select a Flight: Mercury").Image("reserveFlights")_;_script infofile_;_ZIP::ssf19.xml_;_
		While Browser("Find a Flight: Mercury").Page("Book a Flight: Mercury").WebEdit("passFirst0").Exist = False
			wait 1
		Wend
		wait 2		
End Sub
 @@ script infofile_;_ZIP::ssf14.xml_;_
Call FlightDetails() @@ script infofile_;_ZIP::ssf2.xml_;_
Call Preferences()

