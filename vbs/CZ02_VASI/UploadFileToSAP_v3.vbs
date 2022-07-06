Option Explicit
Const PROJECT = "SK01_LOAD4ME"
Const WDIR = "C:\!AUTO\SK01_SAP_CSV_UPLOADER"
Const CRFILE = "C:\!AUTO\CREDENTIALS\logins.txt"
Const COFILE = "C:\!AUTO\CONFIGURATION\LOAD4ME.conf"
Const SPSOURCE = "Load4Me"
Const COMPANYCODE = 0
Const DOCTYPE = 1
Const POSTINGSUBMODUL = 2
Const SESSIONNAME = 3
Const DOCHDRTXT = 4
Const CURRENCYDEC = 5
Const DATESTAMP = 6
Const adSaveCreateOverWrite = 2
Dim oRX : Set oRX = New RegExp 
Dim oSP : Set oSP = New SP
Dim oSAP : Set oSAP = New SAPLauncher
Dim oWSH : Set oWSH = CreateObject("Wscript.Shell")
Dim oXML : Set oXML = CreateObject("MSXML2.DOMDocument")
Dim oXMLCONF : Set oXMLCONF = CreateObject("MSXML2.DOMDocument")
Dim oXMLREST : Set oXMLREST = CreateObject("MSXML2.DOMDocument")
Dim oHTTP : Set oHTTP = CreateObject("MSXML2.XMLHTTP")
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim oFILES : Set oFILES = CreateObject("Scripting.Dictionary")
Dim oMAIL : Set oMAIL = New Mailer
Dim oMAILADMIN : Set oMAILADMIN = New Mailer 
Dim oSTREAM : Set oSTREAM = CreateObject("ADODB.Stream")
Dim oSS ' SAP session
Dim arrFileParts
Dim strProcessingReport
Dim file
Dim arg
Dim strSAPSystem : strSAPSystem = Null
Dim strSAPClient : strSAPClient = Null
oRX.Pattern = "^.*\.(csv|CSV)$"
oMAIL.AddAdmin = "tomas.ac@volvo.com;tomas.chudik@volvo.com"
oMAILADMIN.AddAdmin = "tomas.chudik@volvo.com"
oXMLREST.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
oXMLREST.setProperty "SelectionNamespaces","xmlns:d=""http://schemas.microsoft.com/ado/2007/08/dataservices"""

For arg = 0 To WScript.Arguments.Count - 1

	Select Case WScript.Arguments.Item(arg)
	
		Case "-s","--system"
			strSAPSystem = WScript.Arguments.Item(arg + 1)
			
		Case "-c","--client"
			strSAPClient = WScript.Arguments.Item(arg + 1)
			
	End Select 
	
Next

'Verify arguments have been passed to the script
If IsNull(strSAPSystem) Or IsNull(strSAPClient) Then 
	WScript.Quit
End If 

'Create working directory if it doesn't exist
If Not oFSO.FolderExists(WDIR) Then
	oFSO.CreateFolder WDIR
End If

'Load credentials file
oXMLCONF.load CRFILE

'Get SP credentials
Dim strSPUser : strSPUser = oXMLCONF.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/username").text
Dim strSPSecret : strSPSecret = oXMLCONF.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/password").text
Dim strSPHost : strSPHost = oXMLCONF.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/host").text
Dim strSPDomain : strSPDomain = oXMLCONF.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/domain").text

'Load configuration file
oXMLCONF.load COFILE

'Initialize SP object
oSP.Init strSPHost,strSPDomain,strSPUser,strSPSecret 

'Request files in the queue
With oHTTP
	.open "GET", "https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it" _
	           & "/_api/web/lists/getbytitle('Load4Me')/items?" _
	           & "$select=Processed&$filter=(Processed eq 'Queued')&$expand=File", False 
	           
	.setRequestHeader "Authorization", "Bearer " & oSP.GetToken
	.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
	.setRequestHeader "X-RequestDigest", oSP.GetDigest
	.send
End With 

'If HTTP request fails, send mail and exit the script with the HTTP Status
If Not oHTTP.status = 200 Then
	SpRestError
End If 

'If HTTP request OK, load xml response
oXML.loadXML oHTTP.responseText
oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
oXML.setProperty "SelectionNamespaces","xmlns:d=""http://schemas.microsoft.com/ado/2007/08/dataservices"""
Dim nodes,node,fname,furl,confignodes,iurl,fsize,pname
Set nodes = oXML.selectNodes("//entry//entry//d:Name")

'No queued files, quit, send message and exit
If nodes.length = 0 Then
	oMAIL.SendMessageSAP Now & "<br>No queued files<br><br>","I",strSAPSystem
	WScript.Quit
End If 

'Some queued files, continue
'Launch SAP
oSAP.SetClientName = strSAPClient
oSAP.SetSystemName = strSAPSystem
oSAP.SetLocalXML = oWSH.ExpandEnvironmentStrings("%APPDATA%") & "\SAP\Common\SAPUILandscape.xml"
oSAP.CheckSAPLogon
oSAP.FindSAPSession
If Not oSAP.SessionFound Then
	debug.WriteLine "Session found -> " & oSAP.SessionFound
	oMAIL.SendMessageSAP "Error: SAP session not found","E",strSAPSystem
	WScript.Quit(1)
Else
	Set oSS = oSAP.GetSession ' get sap session
	oSAP.KillPopups(oSS)
End If 

'For loop where files get verified and processed based on rules defined in the config file
For node = 0 To nodes.length - 1
	fname = nodes.item(node).text ' File name
	furl = nodes.item(node).parentNode.parentNode.parentNode.selectSingleNode("id").text ' Full file url
	fsize = CInt(nodes.item(node).parentNode.selectSingleNode("d:Length").text)			 ' File size in bytes. Skip empty
	
	arrFileParts = Split(fname,"_") ' Split the file name
	pname = UCase(Right(arrFileParts(SESSIONNAME),1)) ' Project name i.e 'A'
	
	Set confignodes = oXMLCONF.selectNodes("//companycode[@code=""" & UCase(arrFileParts(COMPANYCODE)) & """]")
	
	If confignodes.length > 1 Then 
		
		PatchRecord "{""Processed"":""Error"",""Details"":""[Preupload] Duplicate configuration entry""}"
		oFILES.Add fname,"Duplicate configuration entry for company code '" & arrFileParts(COMPANYCODE) & "'"
		                        
	ElseIf confignodes.length < 1 Then 
						   
		PatchRecord "{""Processed"":""Error"",""Details"":""[Preupload] Missing configuration entry""}"
		oFILES.Add fname, "Missing configuration entry for company code '" & arrFileParts(COMPANYCODE) & "'"
				   
	ElseIf UCase(arrFileParts(COMPANYCODE)) <> Left(UCase(arrFileParts(SESSIONNAME)),4) Then
		
		PatchRecord "{""Processed"":""Error"",""Details"":""[Preupload] Company code / session name mismatch""}"
		oFILES.Add fname, "Company code - session name mismatch"
		
	ElseIf fsize = 0 Then
	
		PatchRecord "{""Processed"":""Error"",""Details"":""[Preupload] Empty file""}"
		oFILES.Add fname, "Empty file"
		
	ElseIf Not oRX.Test(fname) Then 
	
		PatchRecord "{""Processed"":""Error"",""Details"":""[Preupload] Bad file type""}"
		oFILES.Add fname, "Bad file type"
	
	ElseIf confignodes.length = 1 Then
		On Error Resume Next
		err.Clear
		If err.number <> 0 Or err.number = 424 Then 'If attribute missing then not allowed
			PatchRecord "{""Processed"":""Error"",""Details"":""[Preupload] Doctype '" & arrFileParts(DOCTYPE) & "' not allowed""}"
			oFILES.Add fname, "Doctype '" & arrFileParts(DOCTYPE) & "' not allowed"
			On Error GoTo 0 
		Else  	
		'1) Download the file
			On Error GoTo 0
			DownloadFile
		'2) Upload to SAP
		 	On Error GoTo 0
			UploadToSAP
		End If
	End If  	
Next 

'Exit SAP session
oSS.findById("wnd[0]/tbar[0]/okcd").text="/NEX"
oSS.findById("wnd[0]").sendVKey 0

'Build the report message
For Each file In oFILES.Keys
	strProcessingReport = strProcessingReport & file & " :::> " & oFILES.Item(file) & "<br>"
Next 

'wdapp - checkin
Checkin PROJECT, CRFILE
oMAIL.SendMessageSAP Now & "<br>Processing report<br><br>" & strProcessingReport,"I",strSAPSystem
WScript.Quit(0)
'************************************
'********* M A I N   E N D  *********
'************************************

Sub UploadToSAP
		Randomize
		arrFileParts(SESSIONNAME) = UCase(arrFileParts(SESSIONNAME)) & "-" & Right(CDbl(Rnd(1)) * (Second(Time) + Minute(Time)),5) ' Session from the file name + the 'randomly' generated suffix
		debug.WriteLine "UploadToSAP() -> SESSIONNAME: " & arrFileParts(SESSIONNAME)
		oSAP.KillPopups(oSS)
		oSS.findById("wnd[0]/tbar[0]/okcd").text="/NZTC_ZCVDOC47"
		oSS.findById("wnd[0]").sendVKey 0
		oSAP.KillPopups(oSS)
		oSS.findById("wnd[0]/usr/ctxtDSNSET1").text= WDIR & "\" & fname		  			'path&filenamewithmax128characters
		oSS.findById("wnd[0]/usr/txtW_SEPCHR").text=";"									'characterforseparation
		oSS.findById("wnd[0]/usr/txtW_FEEDID").text=""									'FeederId(exit)
		oSS.findById("wnd[0]/usr/txtW_FILEID").text=""									'FileId(exit)
		oSS.findById("wnd[0]/usr/ctxtW_BUKRS").text= UCase(arrFileParts(COMPANYCODE))	'CompanyCode
		oSS.findById("wnd[0]/usr/txtW_HTXT").text= arrFileParts(DOCHDRTXT)				'DocumentHeaderText
		oSS.findById("wnd[0]/usr/ctxtW_BLART").text= arrFileParts(DOCTYPE)				'Documenttype
		
		If UCase(arrFileParts(POSTINGSUBMODUL)) = "GL" Then
			oSS.findById("wnd[0]/usr/radW_TYPE1").select								'PostinginsubmoduleAR (GL)
		ElseIf UCase(arrFileParts(POSTINGSUBMODUL)) = "AR" Then 
			oSS.findById("wnd[0]/usr/radW_TYPE2").select								'PostinginsubmoduleAP (AR)
		ElseIf UCase(arrFileParts(POSTINGSUBMODUL)) = "AP" Then
			oSS.findById("wnd[0]/usr/radW_TYPE3").select								'PostinginsubmoduleAP (AP)
		End If 
		
		If UCase(arrFileParts(CURRENCYDEC)) = "2D" Then
			oSS.findById("wnd[0]/usr/radW_CURR1").select								' 2 decimal places
		ElseIf UCase(arrFileParts(CURRENCYDEC)) = "0D" Then
			oSS.findById("wnd[0]/usr/radW_CURR2").select								' 0 decimal places
		End If 
		
		oSS.findById("wnd[0]/usr/txtP_SESS").text = arrFileParts(SESSIONNAME) 			'Session name
		oSS.findById("wnd[0]/usr/chkP_SUBM").selected = -1								'Automaticsubmitselected (-1 or 1 )
		oSS.findById("wnd[0]").sendVKey 8												'Execution F8
		oSAP.KillPopups(oSS)
		

		If Len(oSS.findById("wnd[0]/sbar/pane[0]").text) > 0 Then
			PatchRecord "{""Processed"":""Error"",""Details"":""[Upload] " & Replace(oSS.findById("wnd[0]/sbar/pane[0]").text,"\","\\") & """}"
			oFILES.Add fname, oSS.findById("wnd[0]/sbar/pane[0]").text
		Else
			PatchRecord "{""Processed"":""Processed"",""SessionID"":""" & arrFileParts(SESSIONNAME) & """}" ' Upload part is OK, patch only Processed column
			SM35 ' Call SM35 check
		End If 
		
End Sub 



'**************************
'SM35 check function
'**************************
Sub SM35()
	debug.WriteLine "SM35() -> SESSIONNAME: " & arrFileParts(SESSIONNAME)
	Dim i,badmin
	
	If oXMLCONF.selectNodes("//contact[@companycode=""" & UCase(arrFileParts(COMPANYCODE)) & """ and @project=""" & pname & """ and @doctype=""" & UCase(arrFileParts(DOCTYPE)) & """]").length = 0 Then
		badmin = "Business admin contact missing"
	Else
		badmin = oXMLCONF.selectNodes("//contact[@companycode=""" & UCase(arrFileParts(COMPANYCODE)) & """ and @project=""" & pname & """ and @doctype=""" & UCase(arrFileParts(DOCTYPE)) & """]").item(0).text
	End If 
	
	oSAP.KillPopups(oSS)
	If Not LCase(oSS.Info.Transaction) = "sm35" Then	
		oSS.findById("wnd[0]/tbar[0]/okcd").text = "/nSM35"
		oSS.findById("wnd[0]").sendVKey 0
	End If 
	oSAP.KillPopups(oSS)
	oSS.findById("wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/txtD0100-MAPN").text = arrFileParts(SESSIONNAME)
	oSS.findById("wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/ctxtD0100-VON").text = Day(Date) & "." & Month(Date) & "." & Year(Date)
	oSS.findById("wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/ctxtD0100-BIS").text = Day(Date) & "." & Month(Date) & "." & Year(Date)
	'oSS.findById("wnd[0]/usr/subD1000_HEADER:SAPMSBDC_CC:1005/txtD0100-CREATOR").text = strUser
	oSS.findById("wnd[0]").sendVKey 0
	oSAP.KillPopups(oSS)
	
	'At this point there should be at least one session
	If CInt(oSS.findById("wnd[0]/usr/subD1000_FOOT:SAPMSBDC_CC:1015/txtTC_APQI-LINES").text) = 0 Then
		PatchRecord "{""Checked"":""Errors"",""Details"":""[SM35] No SM35 session found""}"
		oFILES.Add fname, "No SM35 session found"
		Exit Sub 
	End If 
	
	'Simple loop. Wait total of 25 seconds max, then move on
	For i = 0 To 5
		oSS.findById("wnd[0]").sendVKey 0
		If LCase(oSS.findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI/lblITAB_APQI-STATUS[1,0]").tooltip) = "processed" Then
			PatchRecord "{""Checked"":""Processed"",""Details"":""[SM35] OK""}" ' SM35 check ok, patch SM35(Checked) column
			oFILES.Add fname, "Upload -> OK, SM35 -> OK"
			Exit Sub
		ElseIf LCase(oSS.findById("wnd[0]/usr/tabsD1000_TABSTRIP/tabpALLE/ssubD1000_SUBSCREEN:SAPMSBDC_CC:1010/tblSAPMSBDC_CCTC_APQI/lblITAB_APQI-STATUS[1,0]").tooltip) = "errors" Then
			PatchRecord "{""Checked"":""Errors"",""Details"":""[SM35] ERRORS"",""BusinessAdmin"":""" & badmin & """}" ' SM35 chec ok, patch SM35(Checked) column
			oFILES.Add fname, "Upload -> OK, SM35 -> ERRORS"
			Exit Sub
		End If
		WScript.Sleep 5000
	Next 
	

	PatchRecord "{""Processed"":""Processed"",""Checked"":""Timed out"",""Details"":""Could not verify after 5 retries"",""BusinessAdmin"":""" & badmin & """}"
	oFILES.Add fname, "Upload -> OK, SM35 -> Timed out after 5 retries"
	'Unable to verify file status after 25 seconds. Patch the record's 'details' column	
End Sub

Sub PatchRecord(strJson)
	With oHTTP
		.open "PATCH", Replace(furl,"/File",""),False
		.setRequestHeader "Accept","application/json;odata=verbose"
		.setRequestHeader "Content-Type","application/json"
		.setRequestHeader "Authorization","Bearer " & oSP.GetToken
		.setRequestHeader "If-Match","*"
		.send strJson
	End With
	debug.WriteLine oHTTP.responseText
End Sub 


Sub SpRestError()
	oXMLREST.loadXML oHTTP.responseText
	oMAIL.SendMessageSAP "Error code: " & oXMLREST.selectSingleNode("//m:code").text & "<br>" _
				       & "Error message: " & oXMLREST.selectSingleNode("//m:message").text, "E",strSAPSystem
	WScript.Quit(oHTTP.status)
End Sub 


Sub DownloadFile()
	With oHTTP
		.open "GET", furl & "/$value",False
		.setRequestHeader "Authorization","Bearer " & oSP.GetToken
		.send
	End With 
	If Not oHTTP.status = 200 Then
		SpRestError
	End If 
	oSTREAM.Open
	oSTREAM.Type = 1
	oSTREAM.Write oHTTP.responseBody
	oSTREAM.SaveToFile WDIR & "\" & fname, adSaveCreateOverWrite
	oSTREAM.Close
End Sub 

Class SP
	
	Private oXML
	Private strAuthUrlPart1
	Private strAuthUrlPart2
	Private vti_bin_clientURL
	Private oHTTP
	Private oFSO
	Private strClientID
	Private strSecurityToken
	Private strClientSecret
	Private strFormDigestValue
	Private strTenantID
	Private strResourceID
	Private strURLbody
	Private strSiteURL
	Private strDomain
	Private numHTTPstatus
	Private strSite
	Private errDescription
	Private errNumber
	Private errSource
	
	Private Sub Class_Initialize
		
		errDescription = ""
		errNumber = 0
		numHTTPstatus = 0
		strAuthUrlPart1 = "https://accounts.accesscontrol.windows.net/"
		strAuthUrlPart2 = "/tokens/OAuth/2"
		Set oFSO = CreateObject("Scripting.FileSystemObject")
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		Set oXML = CreateObject("MSXML2.DOMDocument")
	End Sub 
	
	Public Function Init(sSiteUrl,sDomain,sClientID,sClientSecret)
	
		Dim parts
		If Right(sSiteUrl,1) = "/" Then
			strSiteURL = sSiteUrl
		Else
			strSiteURL = sSiteUrl & "/"
		End If
		
		parts = Split(strSiteURL,"/")
		
		strSite = parts(UBound(parts) - 1) 
		
		If Left(sDomain,1) = "/" Then
			sDomain = Right(sDomain,Len(sDomain) - 1)
		End If
		If Right(sDomain,1) = "/" Then
			sDomain = Left(sDomain,Len(sDomain) - 1)
		End If 
		strDomain = sDomain
		strClientID = sClientID
		strClientSecret = sClientSecret
		
		If Right(sSiteUrl,1) = "/" Then
			vti_bin_clientURL = sSiteUrl & "_vti_bin/client.svc"
		Else
			vti_bin_clientURL = sSiteUrl & "/_vti_bin/client.svc"
		End If 
		
		GetTenantID      ' Obtain the Tenant/Realm ID
		GetSecurityToken ' Obtain the Security Token
		GetXDigestValue  ' Obtain the form digest value
	
	End Function 
		
	'********************** P R I V A T E   F U N C T I O N S ************************
	Private Function GetTenantID()
	
		Dim part,parts,header
		oHTTP.open "GET",vti_bin_clientURL,False
		oHTTP.setRequestHeader "Authorization","Bearer"
		oHTTP.send
	
		parts = Split(oHTTP.getResponseHeader("WWW-Authenticate"),",")
	
		For Each part In parts 
	
			If InStr(part,"Bearer realm") > 0 Then
				header = Split(part,"=")
				strTenantID = header(1)
				strTenantID = Mid(strTenantID,2,Len(strTenantID) - 2)
			End If 
		
			If InStr(part,"client_id") > 0 Then
				header = Split(part,"=")
				strResourceID = header(1)
				strResourceID = Mid(strResourceID,2,Len(strResourceID) - 2)
			End If 		
		Next
	
	End Function
	
	Private Function GetXDigestValue()
		
		Dim colElements
		oHTTP.open "POST", strSiteURL & "_api/contextinfo", False 
		oHTTP.setRequestHeader "accept","application/atom+xml;odata=verbose"
		oHTTP.setRequestHeader "authorization", "Bearer " & strSecurityToken
		oHTTP.send
		oXML.loadXML oHTTP.responseText
		Set colElements = oXML.getElementsByTagName("d:FormDigestValue")
		strFormDigestValue = colElements.item(0).text 
		
	End Function
	
	Private Function GetSecurityToken
	
		Dim oHTTP,part,parts,tokens,token
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		strURLbody = "grant_type=client_credentials&client_id=" & strClientID & "@" & strTenantID & "&client_secret=" & strClientSecret & "&resource=" & strResourceID & "/" & strDomain & "@" & strTenantID
		oHTTP.open "POST", strAuthUrlPart1 & strTenantID & strAuthUrlPart2, False
		oHTTP.setRequestHeader "User-Agent","Mozilla/5.0 (compatible; MSIE 9.0; Windows NT 6.1; WOW64; Trident/5.0)"
		oHTTP.setRequestHeader "Host","accounts.accesscontrol.windows.net"
		oHTTP.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		oHTTP.setRequestHeader "Content-Length", CStr(Len(strURLbody))
		oHTTP.send strURLbody
		parts = Split(oHTTP.responseText,",")
		For Each part In parts
			If InStr(part,"access_token") > 0 Then
				tokens = Split(part,":")
				Exit For
			End If
		Next
		
		
		token = Mid(tokens(1),2,Len(tokens(1)) - 3)
		strSecurityToken = token
		
		
	End Function 
	
	Private Function Strip(sString)
		If Right(sString,1) = "/" Then
			sString = Mid(sString,1,Len(sString) - 1)
		End If 
		If Left(sString,1) = "/" Then
			sString = Mid(sString,2,Len(sString) - 1)
		End If
		
		Strip = sString
	End Function 
	
	Public Function GetListItem(sListName,sFieldName,sFieldValue)
		Dim oHTTP
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items?$select=" & sFieldName & "&$filter=" & sFieldName & " eq " & "'" & sFieldValue & "'", False 
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
		
		If Not oHTTP.status = 200 Then
			GetListItem = False ' Something went wrong. Assume the item doesn't exist an owervrite it. Or lose it !
			Exit Function
		End If 
		
		oXML.loadXML oHTTP.responseText
		
		If oXML.getElementsByTagName("d:Title").length > 0 Then
			If sFieldValue = oXML.getElementsByTagName("d:Title").item(0).text Then
				GetListItem = True
				Exit Function
			Else
				GetListItem = False
				Exit Function
			End If
		Else
			GetListItem = False
			Exit Function
		End If 
	End Function
	
	Public Function UpdateList()
	
		Dim oHTTP,body,oSTREAM
		body = "<?xml version=""1.0"" encoding=""utf-8""?>" _
		& "<soap12:Envelope xmlns:xsi=""http://www.w3.org/2001/XMLSchema-instance"" xmlns:xsd=""http://www.w3.org/2001/XMLSchema"" xmlns:soap12=""http://www.w3.org/2003/05/soap-envelope"">" _
		& "<soap12:Body><UpdateListItems xmlns=""http://schemas.microsoft.com/sharepoint/soap/"">" _
		& "<listName>SK01_Manual_Payments_QA</listName><updates><Field Name=""ID"">21<Field><Field Name=""Title"">HELLO</Field></updates></UpdateListItems></soap12:Body></soap12:Envelope>"
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")
		oHTTP.open "POST","https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it/_vti_bin/Lists.asmx",False
		oHTTP.setRequestHeader "Host","volvogroup.sharepoint.com"
		oHTTP.setRequestHeader "Content-Type","application/soap+xml; charset=utf-8"
		oHTTP.setRequestHeader "Content-Length",Len(body)
		oHTTP.send body 
	End Function 
	
	Public Function GetFileInfo(sServerRelFilePath)
		
		If Not Left(sServerRelFilePath,1) = "/" Then
			sServerRelFilePath = "/" & sServerRelFilePath
		End If 
		 
		With oHTTP
			.open "GET", strSiteURL & "_api/web/getFileByServerRelativeUrl('/sites/" & strSite & sServerRelFilePath & "')/Properties"
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With 
		
		
		
		If Not oHTTP.status = 200 Then
			GetFileInfo = -1
			Exit Function
		End If 
		
		debug.WriteLine oHTTP.responseText
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		debug.WriteLine oXML.getElementsByTagName("d:vti_x005f_filesize").item(0).text
		
	End Function 
	
			
			
	Public Function DownloadFile(sServerRelFilePath,sSaveAsPath)
	
		Dim oHTTP,oSTREAM
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP")
		With oHTTP
			.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & sServerRelFilePath
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.send
		End With
		
		If oHTTP.status = 200 Then
			Set oSTREAM = CreateObject("ADODB.Stream")
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			oSTREAM.SaveToFile sSaveAsPath
			oSTREAM.Close
		Else
			debug.WriteLine oHTTP.status
			debug.WriteLine oHTTP.responseText
		End If 
		
	End Function 
	
			
	Public Function GetFileCount(sServerRelDirPath)
	
		Dim oHTTP,oXML,colElements
		sServerRelDirPath = Strip(sServerRelDirPath)
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		
		With oHTTP
			.open "GET", oSP.GetSiteURL & "_api/web/GetFolderByServerRelativeUrl('/sites/" & strSite & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & oSP.GetToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
		
		If Not oHTTP.status = 200 Then 
			GetFileCount = -1
			Exit Function
		End If 

		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		Set colElements = oXML.getElementsByTagName("d:Name")
		
		GetFileCount = colElements.length
				
	End Function 
	
	Public Function FolderExists(sRelDirPath)
		
		Dim oHTTP,oXML,colElements
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sRelDirPath & "')", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
		
		If oHTTP.status = 200 Then
			oXML.loadXML oHTTP.responseText
			Set colElements = oXML.getElementsByTagName("d:Exists")
			If colElements.length > 0 Then
				If LCase(colElements.item(0).text) = "true" Then
					FolderExists = True
					Exit Function
				Else
					FolderExists = False
					Exit Function
				End If 
			End If  
		Else
			FolderExists = False
		End If 
	
	End Function 
	
	Public Function DownloadFilesA(sServerRelDirPath,sDestinationFolder,ByRef arrFiles)
		
		Dim item,nodes
		Dim path
		Dim oSTREAM : Set oSTREAM = CreateObject("ADODB.Stream")
		
		For item = 0 To UBound(arrFiles)
			
			With oHTTP
				.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sServerRelDirPath & "')/Files('" & arrFiles(item) & "')", False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With
			
			
			If Not oHTTP.status = 200 Then
				errSource = "MSXML2.XMLHTTP"
				errDescription = oHTTP.responseText
				errNumber = oHTTP.status
				DownloadFilesA = -1
				Exit Function
			End If 
			
			oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
			oXML.loadXML oHTTP.responseText
			
			Set nodes = oXML.getElementsByTagName("d:ServerRelativeUrl")
			
			If Not nodes.length > 0 Then
				errSource = "MSXML2.XMLHTTP"
				errDescription = "d:serverRelativeUrl node missing. Affected file: " & arrFiles(item)
				errNumber = Hex(1000)
				DownloadFilesA = -1
				Exit Function
			End If 
			
			path = nodes.nextNode.text ' Save relative URl
			debug.WriteLine strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & path
			With oHTTP
				.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & path, False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With
			
			If Not oHTTP.status = 200 Then
				errSource = "MSXML2.XMLHTTP"
				errDescription = oHTTP.responseText
				errNumber = oHTTP.status
				DownloadFiles = -1
				Exit Function
			End If 
			
			
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			On Error Resume Next 
			oSTREAM.SaveToFile oFSO.BuildPath(sDestinationFolder,oFSO.GetFileName(path))
			
			If err.number > 0 Then 
				errSource = err.Source
				errDescription = err.Description
				errNumber = err.number
				DownloadFiles = -1
				oSTREAM.Close
				Exit Function
			End If 
			
			oSTREAM.Close
			
		Next
		
	End Function 
	
	
	
	
	
	
	Public Function DownloadFiles(sServerRelDirPath,sDestinationFolder)
	
		Dim item,nodes
		Dim path
		Dim oSTREAM : Set oSTREAM = CreateObject("ADODB.Stream")
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('/sites/" & strSite & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.send
		End With 
		
		If Not oHTTP.status = 200 Then
			errDescription = oHTTP.responseText
			errNumber = oHTTP.status
			DownloadFiles = -1
			Exit Function
		End If 
		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		
		Set nodes = oXML.getElementsByTagName("d:ServerRelativeUrl")
		
		For item = 0 To nodes.length - 1
			path = nodes.nextNode.text
			With oHTTP
				.open "GET", strSiteURL & "_layouts/15/download.aspx?SourceUrl=https://" & strDomain & path, False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With
			
			If Not oHTTP.status = 200 Then
				errDescription = oHTTP.responseText
				errNumber = oHTTP.status
				DownloadFiles = -1
				Exit Function
			End If 
			
		
			oSTREAM.Open
			oSTREAM.Type = 1
			oSTREAM.Write oHTTP.responseBody
			On Error Resume Next 
			oSTREAM.SaveToFile oFSO.BuildPath(sDestinationFolder,oFSO.GetFileName(path))
			
			If err.number > 0 Then 
				errSource = err.Source
				errDescription = err.Description
				errNumber = err.number
				DownloadFiles = -1
				Exit Function
			End If 
			
			oSTREAM.Close
			
		Next 
		
		DownloadFiles = nodes.length
		
	End Function 
	
	Public Function GetFilesA(sServerRelDirPath,ByRef dictFiles)
	
		Dim item,nodes
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('/sites/" & strSite & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With 
		
		If Not oHTTP.status = 200 Then
			GetFilesA = -1
			Exit Function
		End If 
		
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		oXML.loadXML oHTTP.responseText
		
		Set nodes = oXML.getElementsByTagName("d:Name")
		
		For item = 0 To nodes.length - 1
			dictFiles.Add nodes.nextNode.text,""
		Next 
		
		GetFilesA = dictFiles.Count
	
	End Function 
		
	Public Function GetFiles(sServerRelDirPath,ByRef colFiles) ' sType "json" or "atom+xml"
	
		Dim dictFilesInSourceDir
		Dim dictFiles
		Dim oHTTP,oXML,colItems,item,colPaths,path
		Set dictFiles = CreateObject("Scripting.Dictionary")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		If Left(sServerRelDirPath,1) = "/" Then
			sServerRelDirPath = Right(sServerRelDirPath,Len(sServerRelDirPath) - 1)
		End If
		If Right(sServerRelDirPath,1) = "/" Then
			sServerRelDirPath = Left(sServerRelDirPath,Len(sServerRelDirPath) - 1)
		End If 
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/GetFolderByServerRelativeUrl('" & sServerRelDirPath & "')/Files", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With
	
		If oHTTP.status = 200 Then
			oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
			oXML.loadXML oHTTP.responseText
			Set colItems = oXML.getElementsByTagName("d:Name")
			Set colPaths = oXML.getElementsByTagName("d:ServerRelativeUrl")
		
			For item = 0 To colItems.length - 1
				colFiles.add colItems.item(item).text,colPaths.item(item).text
			Next
			
		End If 
	
	End Function
			
			
	Public Function MoveFile2(sSourceRelDirPath,sDestRelDirPath)

		If Left(sSourceRelDirPath,1) = "/" Then
			sSourceRelDirPath = Right(sSourceRelDirPath,Len(sSourceRelDirPath) - 1)
		End If 
		If Left(sDestRelDirPath,1) = "/" Then
			sDestRelDirPath = "/" & Right(sDestRelDirPath,Len(sDestRelDirPath) - 1)
		End If 
		 
		Dim oHTTP,strBody
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	
		strBody = "{""srcPath"": {""__metadata"": {""type"": ""SP.ResourcePath""},""DecodedUrl"": """ & strSiteURL & sSourceRelDirPath & """},""destPath"": {""__metadata"": {""type"": ""SP.ResourcePath""},""DecodedUrl"": """ & strSiteURL & sDestRelDirPath & """}}"
		
		With oHTTP
			.open "POST","https://volvogroup.sharepoint.com/sites/unit-rc-sk-bs-it/_api/SP.MoveCopyUtil.MoveFileByPath(overwrite=@a1)?@a1=true"
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/json;odata=nometadata"
			.setRequestHeader "Content-Type", "application/json;odata=verbose"
			.setRequestHeader "Content-Length", Len(strBody)
			.send strBody
		End With
		
		If oHTTP.status = 200 Then
			MoveFile2 = 1 ' Return 1 or True if successfull
			Exit Function
		Else
			MoveFile2 = 0 ' Return 0 or False if failed
			Exit Function
		End If 
	
	End Function
	
	
	Public Function AddListItem(sListName,sJsonRequest)
'		To do this operation, you must know the ListItemEntityTypeFullName property of the list And
'		pass that as the value of type in the HTTP request body. Following is a sample rest call to get the ListItemEntityTypeFullName

		Dim oHTTP,oXML,strEntityTypeFullName,colElements,request
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		oXML.setProperty "SelectionNamespaces","xmlns=""http://www.w3.org/2005/Atom"""
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')?$select=ListItemEntityTypeFullName", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
			.send
		End With 
		
		If oHTTP.status = 200 Then
			oXML.loadXML oHTTP.responseText
			Set colElements = oXML.getElementsByTagName("d:ListItemEntityTypeFullName")
			If colElements.length >= 1 Then
				strEntityTypeFullName = colElements.item(0).text
			Else
				AddListItem = -1 ' Couldn't obtain the EntityTypeFullName
				Exit Function 
			End If
		Else 
			AddListItem = -2 ' http error
			Exit Function 
		End If
		
		
		sJsonRequest = "{""__metadata"": { ""type"": """ & strEntityTypeFullName & """ }," & sJsonRequest ' Prepend metadata part
		With oHTTP
			.open "POST", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items", False
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accept", "application/json;odata=verbose"
			.setRequestHeader "Content-Type", "application/json;odata=verbose"
			.setRequestHeader "If-None-Match", "*"
			.setRequestHeader "Content-Length", Len(sJsonRequest)
			.setRequestHeader "X-RequestDigest", strFormDigestValue
			.send sJsonRequest
		End With
		
		If oHTTP.status = 201 Then
			AddListItem = 0 ' Success
			Exit Function
		Else 
			AddListItem = oHTTP.status
			debug.WriteLine oHTTP.responseText
			Exit Function
		End If 
		
		
	End Function 
	
	Public Function DeleteAllItemsInList(sListName)
		
		Dim oHTTP,oXML,element
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		
		With oHTTP
			.open "GET", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items", False 
			.setRequestHeader "Authorization", "Bearer " & strSecurityToken
			.setRequestHeader "Accpet", "application/json;odata=verbose"
			.setRequestHeader "Content-Type", "application/json"
			.send
		End With 
		
		oXML.loadXML oHTTP.responseText
		For Each element In oXML.getElementsByTagName("d:Id")
			With oHTTP
				.open "POST", strSiteURL & "_api/web/lists/GetByTitle('" & sListName & "')/items(" & element.text & ")", False
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.setRequestHeader "Accpet", "application/json;odata=verbose"
				.setRequestHeader "Content-Type", "application/json"
				.setRequestHeader "If-Match", "*"
				.setRequestHeader "X-HTTP-Method", "DELETE"
				.send
			End With 
		Next 
	End Function 
		
			
	Public Function RenewToken()
		
		debug.WriteLine GetSecurityToken
		RenewToken = strSecurityToken ' Return token to the caller
		
	End Function 
	
	
	
	' ************************** P R O P E R T I E S ******************************
	Public Property Get LastErrorSource
	
		LastErrorSource = errSource
		
	End Property 
	
	Public Property Get LastErrorCode
	
		LastErrorCode = errNumber
		
	End Property
	
	Public Property Get LastErrorDesc
	
		LastErrorDesc = errDescription
		
	End Property 
	
	Public Property Get GetDigest
		
		GetDigest = strFormDigestValue
		
	End Property 
	
	Public Property Get GetToken
	
		GetToken = strSecurityToken
		
	End Property 
	
	Public Property Get GetHttpResponse
		GetHttpResponse = oHTTP.responseText
	End Property 
	
	Public Property Get GetHttpResponseHeaders(strHeader) ' If strHeader "*" then get all headers
		If strHeader = "*" Then
		
			GetHttpResponseHeaders = oHTTP.getAllResponseHeaders
			Exit Property
		
		End If
		
		GetHttpResponseHeaders = oHTTP.getResponseHeader(strHeader)
		
	End Property 
	
	Public Property Get GetRealmTenantID
		GetRealmTenantID = strTenantID
	End Property
	
	Public Property Get GetClientID
		GetClientID = strClientID
	End Property 
	
	Public Property Get GetResourceID
		GetResourceID = strResourceID
	End Property 
	
	Public Property Get GetClientSecret
		GetClientSecret = strClientSecret
	End Property 
	
	Public Property Get GetAuthURL
		GetAuthURL = strAuthUrlPart1 & strTenantID & strAuthUrlPart2
	End Property 
	
	Public Property Get GetSiteURL
		GetSiteURL = strSiteURL
	End Property 
	
	Public Property Get GetSiteDomain
		GetSiteDomain = strDomain
	End Property 
	
End Class 



Class Mailer

	Private oEmail
	Private oSysInfo
	Private oUser
	Private strAdmins
	Private oNET
	Private strUserName
	Private strComputerName
	
	Private Sub Class_Initialize
	
		Set oEmail = CreateObject("CDO.Message")
		Set oSysInfo = CreateObject("ADSystemInfo")
		Set oUser = GetObject("LDAP://" & oSysInfo.UserName)
		Set oNET = CreateObject("Wscript.Network")
		strUserName = oNET.UserName
		strComputerName = oNET.ComputerName
		
	End Sub 
	
	Private Sub Class_Terminate
	
	End Sub 
	
	
	Public Function SendMessage(strMessage,strSeverity,strAppendToSubject)
	
		Dim admin,strFrom,oUser
		Set oUser = GetObject("LDAP://" & oSysInfo.UserName)
		For Each admin In Split(GetAdmins,",")
			
			With oEmail 
				.From = oUser.Mail
				.To = admin
				.Subject = strSeverity & ";" & WScript.ScriptName & ";" & Year(Date()) & "-" & right("00" & Month(Date()),2) & "-" & right("00" & Day(Date()),2) & ";" & Time() & ";" & strUserName & ";" & strComputerName & ";" & strAppendToSubject
				.Configuration.Fields.Item _
   				("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				.Configuration.Fields.Item _
    			("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
        		"mailgot.it.volvo.net" 
				.Configuration.Fields.Item _
  	    		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
  	    		'.TextBody = strMessage
  	    		.HTMLBody = strMessage
  	    		.Configuration.Fields.Item("urn:schemas:mailheader:X-MSMail-Priority") = "High"
  	    		.Configuration.Fields.Item("urn:schemas:httpmail:importance") = 2
				.Configuration.Fields.Item("urn:schemas:mailheader:X-Priority") = 2
				.Configuration.Fields.Update
				.Send
			End With 
			
		Next
	
	End Function
	
	Public Function SendMessageSAP(strMessage,strSeverity,strSAPsystemName)

		Dim admin,strFrom,oUser
		Set oUser = GetObject("LDAP://" & oSysInfo.UserName)
		For Each admin In Split(GetAdmins,",")
			
			With oEmail 
				.From = oUser.Mail
				.To = admin
				.Subject = strSeverity & ";" & WScript.ScriptName & ";" & strSAPsystemName & ";" & Year(Date()) & "-" & right("00" & Month(Date()),2) & "-" & right("00" & Day(Date()),2) & ";" & Time() & ";" & strUserName & ";" & strComputerName
				.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/sendusing") = 2
				.Configuration.Fields.Item _
				("http://schemas.microsoft.com/cdo/configuration/smtpserver") = _
	    		"mailgot.it.volvo.net" 
				.Configuration.Fields.Item _
	    		("http://schemas.microsoft.com/cdo/configuration/smtpserverport") = 25
	    		'.TextBody = strMessage
	    		.HTMLBody = strMessage
	    		.Configuration.Fields.Item("urn:schemas:mailheader:X-MSMail-Priority") = "High"
	    		.Configuration.Fields.Item("urn:schemas:httpmail:importance") = 2
				.Configuration.Fields.Item("urn:schemas:mailheader:X-Priority") = 2
				.Configuration.Fields.Update
				.Send
			End With 
			
		Next
	
	End Function
	
	Public Property Let AddAdmin(strEmailAddress)
		strAdmins = strAdmins & strEmailAddress & ","
	End Property 
	
	Public Property Get GetAdmins
		GetAdmins = Left(strAdmins,Len(strAdmins) - 1)
	End Property 
	
End Class




Class SAPLauncher
	
	Private oHTTP
	Private oXML
	Private oWSH
	Private oFSO
	Private oSAPGUI
	Private oAPP
	Private oCON
	Private oSES
	Private strGlobalURL
	Private strLocalLandscapePATH
	Private boolSAPRunning  		' Indicates whether SPA Logon is runniied files
	Private boolSessionFound		' Set to true if session was found or created
	Private strSSN 					' Sap System Name e.g FQ2
	Private strSCN  	    		' Sap Client Name e.g. 105
	Private strSSD          		' Sap System Description. This string is found in the local landscape xml and used to connect to the sap system


	
	
	
	' ============== Constructor & Destructor ==================
	Private Sub Class_Initialize
	
		Set oHTTP = CreateObject("MSXML2.XMLHTTP")
		Set oXML = CreateObject("MSXML2.DOMDocument")
		Set oFSO = CreateObject("scripting.filesystemobject")
		Set oWSH = CreateObject("wscript.shell")
		oSAPGUI = Null
		oAPP = Null
		oCON = Null
		oSES = Null 
		strSSN = Null
		strSCN = Null
		strGlobalURL = Null
		strLocalLandscapePATH = Null
		strSSD = Null
		boolSAPRunning = False
		boolSessionFound = False
		

	End Sub
	
	Private Sub Class_Terminate
	
	End Sub
	
	' ============ P U B L I C  &  P R I V A T E   M E T H O D S  &   S U B R O U T I N E S ===========
	
	

	' ---------- CheckSAPLogon
	Public Sub CheckSAPLogon
	
		Dim oWmi,colProc,proc,oSAP,waitfor
		Set oWmi = GetObject("winmgmts:\\.\root\cimv2")
		Set colProc = oWmi.ExecQuery("SELECT Name, ProcessId FROM Win32_Process")
		
		On Error Resume Next 
		For Each proc In colProc
			If InStr(LCase(proc.Name),"saplogon") > 0 Then 
				Do While True 
					Set oSAPGUI = GetObject("SAPGUI") ' Wait until the object is instantiated
						If IsObject(oSAPGUI) Then
							boolSAPRunning = True
							Exit Sub  ' At this point we can safely assume that SAPLogon is running and SAPGUI object is available
					End If 
				Loop
			End If 
		Next 
		
		On Error GoTo 0 ' Reenable error handling
			
		'Start SAPLogon and open system passed in the command line parameter
		Set proc = oWSH.Exec("C:\Program Files (x86)\SAP\FrontEnd\SapGui\saplogon.exe")
		Set colProc = oWmi.ExecQuery("SELECT Name, ProcessId FROM Win32_Process")
		
		On Error Resume Next 
		For Each proc In colProc
			If InStr(LCase(proc.Name),"saplogon") > 0 Then 
				Do While True 
					Set oSAPGUI = GetObject("SAPGUI") ' Wait until the object is instantiated
						If IsObject(oSAPGUI) Then
							boolSAPRunning = True
							Exit Sub  ' At this point we can safely assume that SAPLogon is running and SAPGUI object is available
					End If 
				Loop
			End If 
		Next 
		
		On Error GoTo 0 ' Reenable error handling
		

	End Sub 
	
	
	
	' ---------- FindSAPSession
	Public Sub FindSAPSession
		Dim waitPeriod,waitTurns,currentTurn
		waitPeriod = 5000 ' miliseconds
		waitTurns = 5 ' 5 x 5000 = 20000 ms / 20 s
		currentTurn = 1
		
		If Not boolSAPRunning Then
			oSES = Null
			Exit Sub 
		End If  
		
		FindSAPSystemDescription
		
		If IsNull(strSSD) Then
			oSES = Null
			Exit Sub
		End If
		 
		Set oAPP = oSAPGUI.GetScriptingEngine
		Select Case oAPP.Children.Count
	
			Case 0 ' No open connections exist
				For currentTurn = 1 To waitTurns
					Set oCON = oAPP.OpenConnection(strSSD,True,False) ' Open a new connection asynchronously
					On Error Resume Next
					Set oSES = oCON.Children(0) ' Attach to the first session
					On Error GoTo 0
					If Not IsObject(oSES) Or IsNull(oSES) Then
						Debug.WriteLine "No session found, waiting " & currentTurn & " out of " & waitTurns & " turns"
						WScript.Sleep waitPeriod
					ElseIf IsObject(oSES) And IsNull(oSES) Then
						Debug.WriteLine "No session found, waiting " & currentTurn & " out of " & waitTurns & " turns"
						WScript.Sleep waitPeriod
					Else
						Exit For 
					End If 	
				Next
				If Not IsObject(oSES) Or IsNull(oSES) Then
						Debug.WriteLine "No session found after 5 retries"
						boolSessionFound = False
						Exit Sub
				End If 
				
				If IsObject(oSES) And Not IsNull(oSES) Then
					If InStr(oSES.findById("wnd[0]/sbar/pane[0]").text,"No user exists") > 0 Then
						oCON.CloseConnection
						boolSessionFound = False 
						debug.WriteLine "Session found: " & boolSessionFound
						Exit Sub 
					End If 
					boolSessionFound = True
					debug.WriteLine "Session found: " & boolSessionFound
					Exit Sub 
				End If
				
				oCON.CloseConnection
				boolSessionFound = False
				debug.WriteLine "Session found: " & boolSessionFound
				Exit Sub 
				
		
			Case Else ' Atleast one connection exists
		
				For Each oCON In oAPP.Children ' connections
					For Each oSES In oCON.Children ' sessions
						If LCase(oSES.Info.SystemName) = LCase(strSSN) Then
							If InStr(oSES.findById("wnd[0]/sbar/pane[0]").text,"No user exists") > 0 Then
								oCON.CloseConnection
								boolSessionFound = False 
								debug.WriteLine "Session found: " & boolSessionFound
								Exit Sub 
							End If 
							boolSessionFound = True
							debug.WriteLine "Session found: " & boolSessionFound
							Exit Sub ' Stop here. We found our desired system. oCON and oSES objects hold our target system
						End If 
					Next
				Next
			
				
			End Select 
			
			Set oCON = oAPP.OpenConnection(strSSD,True,False) ' Open a new connection asynchronously
			On Error Resume Next 
			Set oSES = oCON.Children(0) ' Attach to the first session
			On Error GoTo 0
			
			If IsObject(oSES) And Not IsNull(oSES) Then
				If InStr(oSES.findById("wnd[0]/sbar/pane[0]").text,"No user exists") > 0 Then
					oCON.CloseConnection
					boolSessionFound = False 
					debug.WriteLine "Session found: " & boolSessionFound
					Exit Sub 
				End If 
					boolSessionFound = True
					debug.WriteLine "Session found: " & boolSessionFound
					Exit Sub
			End If 
			
			oCON.CloseConnection
			debug.WriteLine "Session found: " & boolSessionFound
			
	End Sub 

	
	
	' --------- FindSAPSystemDescription
	Private Sub FindSAPSystemDescription
	
		Dim n_ChildNodes,n_ChildNode,uuid,i,j
	
		oXML.load(strLocalLandscapePATH) ' Locally stored XML

		Set n_ChildNodes = oXML.getElementsByTagName("Landscape")
	
		For Each n_ChildNode In n_ChildNodes 
			For Each i In n_ChildNode.childNodes
				If i.baseName = "Services" Then 
					Set n_ChildNodes = i.childNodes
					On Error Resume Next
					For Each j In n_ChildNodes
						
						If Left(LCase(j.attributes.getNamedItem("name").text),3) = LCase(strSSN) Then 
							debug.WriteLine "SAP system name: " & strSSN
							strSSD = j.attributes.getNamedItem("name").text
							debug.WriteLine "SAP system description: " & strSSD
							CheckSAPLogon
							Exit Sub  
						End If
					Next
				End If 
			Next
		Next
		strSSD = Null ' Not found
		debug.WriteLine strSSD
	End Sub 
	
	Public Function GetSession
	
		If IsNull(oSES) Or Not IsObject(oSES) Then
			GetSession = Null
		else
			Set GetSession = oSES
		End If 

	End Function
	
	Public Function KillPopups(ByRef objSession)
		Do While objSession.Children.Count > 1
			If InStr(objSession.ActiveWindow.Text, "System Message") > 0 Then
				objSession.ActiveWindow.sendVKey 12
			ElseIf InStr(objSession.ActiveWindow.Text, "Information") > 0 And InStr(objSession.ActiveWindow.PopupDialogText, "Exchange rate adjusted to system settings") > 0 Then
				objSession.ActiveWindow.sendVKey 0
			ElseIf InStr(objSession.ActiveWindow.Text, "Copyright") > 0 Then
				objSession.ActiveWindow.sendVKey 0
			ElseIf InStr(objSession.ActiveWindow.Text, "License Information for Multiple Logon") > 0 Then
				objSession.findById("wnd[1]/usr/radMULTI_LOGON_OPT2").select
				objSession.ActiveWindow.sendVKey 0
			'ElseIF   'Insert next type of popup windows which you want to kill
			Else
				Exit Do
			End If
		Loop
	End Function 

	' ================= P R O P E R T I E S ====================
	Public Property Get SAPLogonRunning
		SAPLogonRunning = boolSAPRunning
	End Property 	
		
	Public Property Get SAPSessionExists
		If boolSAPRunning And Not IsNull(oSES) Then
			SAPSessionExists = True
		Else 
			SAPSessionExists = False
		End If
	End Property 
	
	Public Property Get SAPsysName
		SAPsysName = strSSN
	End Property 
	
	Public Property Get SAPcliName
		SAPcliName = strSCN
	End Property 
	
	Public Property Get LandscapeURL
		LandscapeURL = strGlobalURL
	End Property 
	
	
	Public Property Get SAPsysDescription
		SAPsysDescription = strSSD
	End Property 
	
	Public Property Get GetGlobalURL
	
		GetGlobalURL = strGlobalURL
		
	End Property 
	
	Public Property Let SetGlobalURL(url)
	
		strGlobalURL = url
		
	End Property 
	
	Public Property Let SetLocalXML(xml)
	
		strLocalLandscapePATH = xml
		
	End Property 
	
	Public Property Get GetLocalXML
	
		GetLocalXML = strLocalLandscapePATH
		
	End Property 
	
	Public Property Let SetSystemName(sys)
	
		strSSN = sys
		
	End Property 
	
	Public Property Let SetClientName(cli)
	
		strSCN = cli
	
	End Property 
	
	Public Property Get SessionFound
	
		SessionFound = boolSessionFound
		
	End Property 
	
		
End Class 




Function Checkin(sProjectName,sCredFilePath)
	Dim sUserName,sUserSecret,sSiteUrl,sDomain,sTenantID,sClientID,sXDigest,sAccessToken,tmp,rxResult
	Dim oHTTP : Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
	Dim oXML : Set oXML = CreateObject("MSXML2.DOMDocument")
	Dim oRX : Set oRX = New RegExp
	
	'Load credentials
	oXML.load sCredFilePath
	sUserName = oXML.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/username").text
	sUserSecret = oXML.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/password").text
	sSiteUrl = oXML.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/host").text
	sDomain = oXML.selectSingleNode("//service[@name=""unit-rc-sk-bs-it""]/domain").text
	
	'Get TenantID & ClientID/ResourceID
	With oHTTP
		.open "GET",sSiteUrl & "/_vti_bin/client.svc",False
		.setRequestHeader "Authorization","Bearer"
		.send
	End With
		
	oRX.Pattern = "Bearer realm=""([a-zA-Z0-9]{1,}-)*[a-zA-Z0-9]{12}"
	Set rxResult = oRX.Execute(oHTTP.getResponseHeader("WWW-Authenticate"))
	oRX.Pattern = "[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
	sTenantID = oRX.Execute(rxResult(0))(0)
	
	oRX.Pattern = "client_id=""[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
	Set rxResult = oRX.Execute(oHTTP.getResponseHeader("WWW-Authenticate"))
	oRX.Pattern = "[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
	sClientID = oRX.Execute(rxResult(0))(0)
	
	'Get AccessToken
	Dim sBody : sBody = "grant_type=client_credentials&client_id=" & sUserName & "@" & sTenantID & "&client_secret=" & sUserSecret & "&resource=" & sClientID & "/" & sDomain & "@" & sTenantID
	With oHTTP
		.open "POST", "https://accounts.accesscontrol.windows.net/" & sTenantID & "/tokens/OAuth/2", False
		.setRequestHeader "Host","accounts.accesscontrol.windows.net"
		.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
		.setRequestHeader "Content-Length", CStr(Len(sBody))
		.send sBody
	End With 
	
	oRX.Pattern = "access_token"":"".*"
	Set rxResult = oRX.Execute(oHTTP.responseText)
	rxResult = Split(rxResult(0),":")
	rxResult(1) = Replace(rxResult(1),"""","")
	rxResult(1) = Replace(rxResult(1),"}","")
	sAccessToken = rxResult(1) ' Save the token 
	
	'Get XDigest
	With oHTTP
		oHTTP.open "POST", sSiteUrl & "/_api/contextinfo", False 
		oHTTP.setRequestHeader "accept","application/atom+xml;odata=verbose"
		oHTTP.setRequestHeader "authorization", "Bearer " & sAccessToken
		oHTTP.send
	End With 
	
	oXML.loadXML oHTTP.responseText
	sXDigest = oXML.selectSingleNode("//d:FormDigestValue").text
	
	
	'Send query
	With oHTTP
		.open "GET", sSiteUrl & "/_api/web/lists/getbytitle('WDAPP')/items?$select=Title&$filter=(Title eq '" & sProjectName & "')", False        
		.setRequestHeader "Authorization", "Bearer " & sAccessToken
		.setRequestHeader "Accept", "application/atom+xml;odata=verbose"
		.setRequestHeader "X-RequestDigest", sXDigest
		.send
	End With 
	

	'Patch record
	Dim oNet : Set oNet = CreateObject("WScript.Network")
	Dim oSysInfo : Set oSysInfo = CreateObject("ADSystemInfo")
	Dim oLDAP : Set oLDAP = GetObject("LDAP://" & oSysInfo.UserName)
	oXML.loadXML oHTTP.responseText
	Dim url : url = oXML.selectSingleNode("//feed").attributes.getNamedItem("xml:base").text
	url = url & oXML.selectSingleNode("//entry/link[@rel=""edit""]").attributes.getNamedItem("href").text
  	
	With oHTTP
		.open "PATCH", url, False
		.setRequestHeader "Accept","application/json;odata=verbose"
		.setRequestHeader "Content-Type","application/json"
		.setRequestHeader "Authorization","Bearer " & sAccessToken
		.setRequestHeader "If-Match","*"
		.send "{""ComputerName"":""" & Trim(oNet.ComputerName) & """,""UserName"":""" & Trim(oLDAP.displayName) & """,""UserID"":""" & Trim(oLDAP.sAMAccountName) & """}"
	End With
	
End Function