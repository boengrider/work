Option Explicit

Const UNIT = "unit-vasi"
Const PROJECT = "MAKE4ME"
Const CONFIG_PATH = "C:\!AUTO\CONFIGURATION\MAKE4ME.conf"
Dim oXML : Set oXML = CreateObject("MSXML2.DOMDOCUMENT")
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim oSPL : Set oSPL = New SharePointLite
Dim config, country
config = Null
country = Null
Dim credfile, spUser, spSecret, spSite
spSite = Null
spUser = Null
spSecret = Null
credfile = Null
Dim retval, arg, token, xdigestvalue



'#############################################
'################ M A I N ####################
'#############################################

'**************************
'Process cli args and load
'Sharepoint credentials
'**************************
'**************************
'@At least one parameter expected
If WScript.Arguments.Count < 1 Then
	debug.WriteLine "Usage: " & WScript.ScriptName & " -c|--country $COUNTRY [-cf|--config $CONFIGPATH]"
	WScript.Quit
End If 
'@Process only selected arguments
For arg = 0 To WScript.Arguments.Count - 1

	Select Case WScript.Arguments.Item(arg)
	
		Case "-c","--country"
			country = UCase(WScript.Arguments.Item(arg + 1))
			
		Case "-cf","--config"
			config = WScript.Arguments.Item(arg + 1)
			
	End Select 
	
Next
'@Can't continue w/o country
If IsNull(country) Then
	debug.WriteLine "Parameter $COUNTRY not optional"
	debug.WriteLine "Usage: " & WScript.ScriptName & " -c|--country $COUNTRY [-cf|--config $CONFIGPATH]"
	WScript.Quit
End If 
'@Fallback config path
If IsNull(config) Then
	config = CONFIG_PATH
End If  
'@Verify config file path
If Not oFSO.FolderExists(oFSO.GetParentFolderName(config)) Then
	'Parent folder does not exist, no need to check further
	'Notify admin
	debug.WriteLine "Parent folder " & oFSO.GetParentFolderName(config) & " does not exist"
	WScript.Quit
ElseIf Not oFSO.FileExists(config) Then
	'Either file is a folder or file does not exist
	'Notify admin
	debug.WriteLine oFSO.GetFileName(config) & " is either folder or file does not exist"
	WScript.Quit
End If  
'@Load config first
oXML.load config
credfile = oXML.selectSingleNode("//Country[@name=""" & country & """]/CredentialsFile").text
'@Can't continue w/o credentials
If IsNull(credfile) Then
	'Notify admin
	WScript.Quit
End If
'@Load credentials file
oXML.load credfile
spUser = oXML.selectSingleNode("//service[@name=""" & UNIT & """]/username").text
spSecret = oXML.selectSingleNode("//service[@name=""" & UNIT & """]/password").text
spSite = oXML.selectSingleNode("//service[@name=""" & UNIT & """]/host").text
'@Can't continue w/o either user,secret or site
If IsNull(spUser) Or IsNull(spSecret) Or IsNull(spSite) Then
	'Notify admin
	debug.WriteLine "Missing info"
	WScript.Quit
End If 
'@Initiliaze SPLite
retval = oSPL.SharePointLite(spSite,spUser,spSecret,False)
If retval <> 0 Then
	'Notify admin
	debug.WriteLine oSPL.LastErrorNumber
	debug.WriteLine oSPL.LastErrorSource
	debug.WriteLine oSPL.LastErrorDesc
	WScript.Quit
Else 'Initialization OK
	token = oSPL.AccessToken
	xdigestvalue = oSPL.XDigest
End If
'@Load configuration subtree based on country
oXML.load config
Set oXML = oXML.selectNodes("//Country[@name=""" & country & """]").item(0)
debug.WriteLine oXML.firstChild.baseName








'#########################################################################
Class SharePointLite
	
	Private oRX 
	Private oXML
	Private strAuthUrlPart1
	Private strAuthUrlPart2
	Private vti_bin_clientURL
	Private oHTTP
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
	Private boolRaise

	
	Private Sub Class_Initialize
		errDescription = ""
		errNumber = 0
		numHTTPstatus = 0
		strAuthUrlPart1 = "https://accounts.accesscontrol.windows.net/"
		strAuthUrlPart2 = "/tokens/OAuth/2"
		boolRaise = False 
		Set oRX = New RegExp
		Set oHTTP = CreateObject("MSXML2.ServerXMLHTTP.3.0")
		Set oXML = CreateObject("MSXML2.DOMDocument")
	End Sub 
	

	Public Function SharePointLite(sSiteUrl,sClientID,sClientSecret,bRaise)
		Dim strErrSource : strErrSource = "SharepointLite.SharePointLite()"
		Dim tmp,retval
		boolRaise = bRaise
		oRX.Global = True
		oRX.Multiline = True
		oRX.IgnoreCase = True
		oRX.Pattern = "(http:\/\/|https:\/\/)([^\/])*\/sites\/([^\/])*\/{0,1}"
		
		If oRX.Test(sSiteUrl) Then
			tmp = oRX.Execute(sSiteUrl)(0)
			If Right(tmp,1) <> "/" Then
				strSiteURL = tmp & "/"
			Else
				strSiteURL = tmp
			End If 
		ElseIf boolRaise Then
			err.Raise 100, strErrSource, "Bad URL -> " & sSiteUrl
		Else
			errSource = strErrSource
			errNumber = 100
			errDescription = "Bad URL -> " & sSiteUrl
			SharePointLite = 100
			Exit Function
		End If 
		
		
		vti_bin_clientURL = strSiteURL & "_vti_bin/client.svc"
		tmp = Split(strSiteURL,"/")
		strDomain = tmp(2)
		strClientID = sClientID
		strClientSecret = sClientSecret
		
		retval = GetTenantID      ' Obtain the Tenant/Realm ID
		If retval <> 0 Then
			SharePointLite = retval
			Exit Function 
		End If 
		
		retval = GetSecurityToken ' Obtain the Security Token
		If retval <> 0 Then
			SharePointLite = retval
			Exit Function
		End If 
		
		retval = GetXDigestValue  ' Obtain the form digest value
		If retval <> 0 Then
			SharePointLite = retval
			Exit Function
		End If 
		
		
		SharePointLite = 0
	
	End Function 
		
	'********************** P R I V A T E   F U N C T I O N S ************************
	
	
	'##############################
	'######### GetTenantID ########
	'##############################
	Private Function GetTenantID()
		Dim rxResult
		Dim strErrSource : strErrSource = "Sharepoint.GetTenantID()"
		
		With oHTTP
			.open "GET",vti_bin_clientURL,False
			.setRequestHeader "Authorization","Bearer"
			.send
		End With
		
		If Not oHTTP.status = 401 And boolRaise Then
			err.Raise oHTTP.status, strErrSource, oHTTP.responseText
		ElseIf Not oHTTP.status = 401 And Not boolRaise Then 
			GetTenantID = oHTTP.status
			Exit Function
		End If 
		
		oRX.Pattern = "Bearer realm=""([a-zA-Z0-9]{1,}-)*[a-zA-Z0-9]{12}"
		If oRX.Test(oHTTP.getResponseHeader("WWW-Authenticate")) Then
			Set rxResult = oRX.Execute(oHTTP.getResponseHeader("WWW-Authenticate"))
			oRX.Pattern = "[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
			If oRX.Test(rxResult(0)) Then 
				strTenantID = oRX.Execute(rxResult(0))(0)
			ElseIf boolRaise Then
				err.Raise 1000, strErrSource, "Bearer realm not found"
			End If 
		ElseIf boolRaise Then
			err.Raise 1000, strErrSource, "Bearer realm not found"
		Else
			errSource = strErrSource
			errNumber = 1000
			errDescription = "Bearer realm not found"
			GetTenantID = 1000
			Exit Function
		End If 
		
		oRX.Pattern = "client_id=""[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
		If oRX.Test(oHTTP.getResponseHeader("WWW-Authenticate")) Then
			Set rxResult = oRX.Execute(oHTTP.getResponseHeader("WWW-Authenticate"))
			oRX.Pattern = "[a-fA-F0-9]{8}-([a-fA-F0-9]{4}-){3}[a-fA-F0-9]{12}"
			If oRX.Test(rxResult(0)) Then 
				strResourceID = oRX.Execute(rxResult(0))(0)
			ElseIf boolRaise Then
				err.Raise 1000, strErrSource, "Client_id not found"
			Else
				GetTenantID = 1000
				Exit Function 
			End If  
		ElseIf boolRaise Then
			err.Raise 1000, strErrSource, "Client_id not found"
		Else
			errSource = strErrSource
			errNumber = 1000
			errDescription = "Client_id not found"
			GetTenantID = 1000
			Exit Function 
		End If 
		
		GetTenantID = 0
	End Function
	
	
	'##############################
	'####### GetXDigestValue ######
	'##############################
	Private Function GetXDigestValue()
		Dim strErrSource : strErrSource = "Sharepoint.GetXDigestValue()"
		Dim colNodes
		
		With oHTTP
			oHTTP.open "POST", strSiteURL & "_api/contextinfo", False 
			oHTTP.setRequestHeader "accept","application/atom+xml;odata=verbose"
			oHTTP.setRequestHeader "authorization", "Bearer " & strSecurityToken
			oHTTP.send
		End With 
		
		If Not oHTTP.status = 200 And boolRaise Then
			err.Raise oHTTP.status, strErrSource, oHTTP.responseText
		ElseIf Not oHTTP.status = 200 Then
			errSource = strErrSource
			errNumber = oHTTP.status
			errDescription = oHTTP.responseText
			GetXDigestValue = oHTTP.status
			Exit Function 
		End If 
		
		oXML.loadXML oHTTP.responseText
		
		Set colNodes = oXML.selectNodes("//d:FormDigestValue")
		
		If colNodes.length = 0 And boolRaise Then
			err.Raise 1100, strErrSource, "FormDigestValue not found"
		ElseIf colNodes.length = 0 Then
			errSource = strErrSource
			errNumber = 1100
			errDescription = "FormDigestValue not found"
			GetXDigestValue = 1100
			Exit Function
		Else 
			strFormDigestValue = colNodes.item(0).text
		End If 
		
		GetXDigestValue = 0	
	End Function
	
	
	'##############################
	'###### GetSecurityToken ######
	'##############################
	Private Function GetSecurityToken()
		Dim rxResult
		Dim strErrSource : strErrSource = "Sharepoint.GetSecurityToken()"
		Dim strURLbody : strURLbody = "grant_type=client_credentials&client_id=" & strClientID & "@" & strTenantID & "&client_secret=" & strClientSecret & "&resource=" & strResourceID & "/" & strDomain & "@" & strTenantID
		
		With oHTTP
			.open "POST", strAuthUrlPart1 & strTenantID & strAuthUrlPart2, False
			.setRequestHeader "Host","accounts.accesscontrol.windows.net"
			.setRequestHeader "Content-Type","application/x-www-form-urlencoded"
			.setRequestHeader "Content-Length", CStr(Len(strURLbody))
			.send strURLbody
		End With 
		
		If Not oHTTP.status = 200 And boolRaise Then
			err.Raise oHTTP.status, strErrSource, oHTTP.responseText
		ElseIf Not oHTTP.status = 200 Then
			errSource = strErrSource
			errNumber = oHTTP.status
			errDescription = oHTTP.responseText
			GetSecurityToken = oHTTP.status
			Exit Function 
		End If 

		oRX.Pattern = "access_token"":"".*"
		If oRX.Test(oHTTP.responseText) Then
			Set rxResult = oRX.Execute(oHTTP.responseText)
			rxResult = Split(rxResult(0),":")
			rxResult(1) = Replace(rxResult(1),"""","")
			rxResult(1) = Replace(rxResult(1),"}","")
			strSecurityToken = rxResult(1) ' Save the token 
		ElseIf boolRaise Then
			 err.Raise 1200, strErrSource, "Access token not found"
		Else
			errSource = strErrSource
			errNumber = 1200
			errDescription = "Access token not found"
			GetSecurityToken = 1200
			Exit Function
		End If 
		
		GetSecurityToken = 0 	
	End Function 
	
	Public Property Get XDigest
		XDigest = strFormDigestValue
	End Property 
		
	Public Property Get AccessToken
		AccessToken = strSecurityToken
	End Property 
	
	Public Property Get LastErrorNumber
		LastErrorNumber = errNumber
	End Property
	
	Public Property Get LastErrorDesc
		LastErrorDesc = errDescription
	End Property
	
	Public Property Get LastErrorSource
		LastErrorSource = errSource
	End Property 
End Class 