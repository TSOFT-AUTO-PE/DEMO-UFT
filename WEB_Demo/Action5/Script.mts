Option Explicit

Dim url, ie, chrome, ff, user, pass, start_time, stop_time, Num_Iter

		url = "http://newtours.demoaut.com/"
		ie = "C:\Program Files\internet explorer\iexplore.exe"
		chrome = "C:\Program Files (x86)\Google\Chrome\Application\chrome.exe"
		ff = "C:\Program Files\Mozilla Firefox\firefox.exe"
		user = DataTable("i_User", dtLocalsheet)
		pass = DataTable("i_Password", dtLocalsheet)
		Num_Iter = Environment.Value("ActionIteration")		
Sub CloseAll()
	
	SystemUtil.CloseProcessByName ("chrome.exe")
	SystemUtil.CloseProcessByName ("iexplore.exe")
	SystemUtil.CloseProcessByName ("firefox.exe")
End Sub
Sub OpenBrowser()
	start_time = Timer
	Select Case  DataTable("i_Navegador", dtLocalsheet)
			Case "ie"
			SystemUtil.Run ie, url
			Case "chrome"
			SystemUtil.Run chrome, url
			Case "ff"
			SystemUtil.Run ff, url		
	End Select
End Sub
Sub LoginWeb()
	Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours").WebEdit("userName").Set user
	Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours").WebEdit("password").Set pass
	Browser("Welcome: Mercury Tours").CaptureBitmap RutaEvidencias() &Num_Iter&"_"&"Login_Mercury.png", True
	imagenToWord "Logeo de WEB MERCURY", RutaEvidencias() &Num_Iter&"_"&"Login_Mercury.png"
	Browser("Welcome: Mercury Tours").Page("Welcome: Mercury Tours").Image("Sign-In").Click
	While Browser("Welcome: Mercury Tours").Page("Find a Flight: Mercury").WebElement("Type:").Exist = False
		wait 1
	Wend
	wait 2
	stop_time = Timer
	DataTable("Timer", dtLocalsheet) = "Se Ejecutó en: ["&stop_time - start_time&"] segundos."	
End Sub

Call CloseAll()
Call OpenBrowser()
Call LoginWeb()
