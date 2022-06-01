Option Explicit
Const UNIT = "unit-vasi"
Const PROJECT = "MAKE4ME"
Const CONFIG_PATH = "C:\!AUTO\CONFIGURATION\MAKE4ME.conf"
Const HOST = "https://volvogroup.sharepoint.com/sites/unit-vasi"
Const LIST = "CZ02_VASI_Portal"
Dim oXML : Set oXML = CreateObject("MSXML2.DOMDOCUMENT")
Dim oXMLCONF : Set oXMLCONF = CreateObject("MSXML2.DOMDOCUMENT")
Dim oFSO : Set oFSO = CreateObject("Scripting.FileSystemObject")
Dim oSPL : Set oSPL = New SharePointLite
Dim oDF : Set oDF = New DateFormatter
Dim oMyItems : Set oMyItems = New MyItems
Dim oMailer : Set oMailer = New Mailer
oMailer.AddAdmin = "tomas.ac@volvo.com;tomas.chudik@volvo.com"
Dim config, country
config = Null
country = Null
Dim credfile, spUser, spSecret, spSite
spSite = Null
spUser = Null
spSecret = Null
credfile = Null
Dim retval, arg, token, xdigestvalue
Dim buffer,entryNode,itemNode,i
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
oXMLCONF.load config
credfile = oXMLCONF.selectSingleNode("//Country[@name=""" & country & """]/CredentialsFile").text
'@Can't continue w/o credentials
If IsNull(credfile) Then
	'Notify admin
	WScript.Quit
End If
'@Load credentials file
oXMLCONF.load credfile
spUser = oXMLCONF.selectSingleNode("//service[@name=""" & UNIT & """]/username").text
spSecret = oXMLCONF.selectSingleNode("//service[@name=""" & UNIT & """]/password").text
spSite = oXMLCONF.selectSingleNode("//service[@name=""" & UNIT & """]/host").text
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
oXMLCONF.load config
Set oXMLCONF = oXMLCONF.selectNodes("//Country[@name=""" & country & """]").item(0)
'****************************
'@Load 'PostingAP' 'ready'  *
'****************************
buffer = oSPL.GetListItems(HOST,LIST,"select=*&$filter=PostingAP eq 'Ready'")
If IsNull(buffer) Then 
	'Notify admin
	'http request failed
	WScript.Quit
End If 
'@Continue if there are any 'entries'
oXML.loadXML buffer
oXML.setProperty "SelectionNamespaces", "xmlns:d=""http://schemas.microsoft.com/ado/2007/08/dataservices"""
oXML.setProperty "SelectionNamespaces", "xmlns:m=""http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"""
If oXML.selectNodes("//entry").length = 0 Then
	'Notify admin
	'Nothing to process
	'Exit
	WScript.Quit
End If 
'@Consume every //entry node in the collection
For Each entryNode In oXML.selectNodes("//entry")
	oMyItems.Consume entryNode, oXMLCONF
Next


'TEMP. Remove later
Dim f : Set f = oFSO.OpenTextFile("C:\!AUTO\CZ02_VASI\out.txt",2,True)
'TEMP. Remove later

'@Write out
Dim vasiItems : Set vasiItems = oMyItems.Items
For Each itemNode In vasiItems.Keys
	f.Write vasiItems.Item(itemNode).GetHeader & vasiItems.Item(itemNode).GetLineItems
	debug.Write vasiItems.Item(itemNode).GetHeader & vasiItems.Item(itemNode).GetLineItems
Next 
f.Close
'@Send message
oMailer.SendMessage "Processed items: " & oMyItems.Count & vbCrLf & vbCrLf & "CCP: " & oMyItems.CCPCount & vbCrLf  _
                  & "OUT: " & oMyItems.OUTCount & vbCrLf & "Other: " & oMyItems.Count - (oMyItems.CCPCount + oMyItems.OUTCount),"I",""
                  


'**********************************************************
'@Load 'PostingAP' 'ready' items of type OUT from SP list *
'**********************************************************
Class MyItems
	
	Private listItems__
	Private invoice__
	Private oRX__
	Private ccpItems__
	Private outItems__
	Private otherItems__
	
	Private Sub Class_Initialize()
		ccpItems__ = 0
		outItems__ = 0
		otherItems__ = 0
		Set oRX__ = New RegExp
		oRX__.Pattern = "[0-9]{1,}"
		oRX__.Global = True 
		Set listItems__ = CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate()
	End Sub 
	
	Public Function Consume(xmlNode__,xmlConfigNode__)
		invoice__ = xmlNode__.selectSingleNode("content").selectSingleNode("m:properties").selectSingleNode("d:Title").text
		
		Select Case UCase(xmlNode__.SelectSingleNode("//content/m:properties/d:DocType").text)
		
			Case "CCP"
				
				If Not listItems__.Exists(invoice__) Then
					'Create a new instance of MyItem
					listItems__.Add invoice__, New MyItem
					ccpItems__ = ccpItems__ + 1
				End If
				 
				'************
				'*  Header  *
				'************
				'DocDate -> header
				listItems__.Item(invoice__).HeaderDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'DocDate -> line item
				listItems__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'PostingDate -> header
				listItems__.Item(invoice__).HeaderPostingDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'PostingDate -> line item
				listItems__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'Currency -> header
				listItems__.Item(invoice__).HeaderCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
				'Currency -> line item
				listItems__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
				'Reference -> header
				listItems__.Item(invoice__).HeaderReference = invoice__
				'Reference -> line items
				listItems__.Item(invoice__).LineItemReference = invoice__
				'Parma -> header
				listItems__.Item(invoice__).HeaderParma = xmlConfigNode__.SelectSingleNode("//VendorParma").text
				'GL -> line item
				listItems__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("//CaseCCP/GL").text
				'TotalAmount -> header
				listItems__.Item(invoice__).HeaderTotalAmount = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text
				'TotalAmount -> line item
				listItems__.Item(invoice__).LineItemTotalAmount = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text
				'TaxCode -> header
				listItems__.Item(invoice__).HeaderTaxCode = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("TaxCode").text
				'TaxCode -> line item
				listItems__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("TaxCode").text
				'TaxAmount -> header
				listItems__.Item(invoice__).HeaderTaxAmount = "0,00"
				'TaxAmount -> line item
				listItems__.Item(invoice__).LineItemTaxAmount = "0,00"
				'CostCenter -> line item
				listItems__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("CC").text
				'ProfitCenter -> line item
				listItems__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("PC").text
				'Allocation -> line item
				listItems__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
				'TradingPartner -> header
				listItems__.Item(invoice__).HeaderTradingPartner = xmlConfigNode__.SelectSingleNode("//TradingPartner").text
				'TradingPartner -> line item
				listItems__.Item(invoice__).LineItemTradingPartner = xmlConfigNode__.SelectSingleNode("//TradingPartner").text
				'AmountInLocCurrency
				listItems__.Item(invoice__).HeaderAmInLocCur = "0,00"
				'AmountInLocCurrency -> line item
				listItems__.Item(invoice__).LineItemAmInLocCur = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text
				'LineText
				listItems__.Item(invoice__).HeaderLineText = invoice__
				'PaymentTerms
				listItems__.Item(invoice__).HeaderPaymentTerms = xmlConfigNode__.SelectSingleNode("//PaymentTerms").text
				
				'Additonal properties
				listItems__.Item(invoice__).ItemType = invoice__
				
			Case "OUT"
				outItems__ = outItems__ + 1
			
			Case Else
				otherItems__ = otherItems__ + 1 
			
		End Select 
		
	End Function
	
	Public Property Get CCPCount
		CCPCount = ccpItems__
	End Property 
	
	Public Property Get OUTCount
		OUTCount = outItems__
	End Property
	
	Public Property Get Count
		Count = listItems__.Count
	End Property 
	
	Public Property Get Items
		Set Items = listItems__
	End Property 
	 
End Class
'This class represents one invoice
Class MyItem
	
	Public 	outHeader__(13) ' Header line array
	Private outBuffer__ 	' GL lines string delimited with CRLF
	'Following variables will be used tu build header line
	Private invoice__ ' invoice string
	Private isInvoice__       ' Invoice or credit note. If true -> invoice else credite note
	Private rx__
	
	Private Sub Class_Initialize()
		Set rx__ = New RegExp
		rx__.Global = True
		outBuffer__ = ""
	End Sub
	
	'***************
	'Properties
	'***************
	
	'Setters
	'Index 0 DocDate
	Public Property Let HeaderDocdate(d)
		rx__.Pattern = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
		outHeader__(0) = Replace(rx__.Execute(d)(0),"-","")
	End Property
	
	'Index 1 PostingDate
	Public Property Let HeaderPostingDate(d)
		rx__.Pattern = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
		outHeader__(1) = Replace(rx__.Execute(d)(0),"-","")
	End Property
	
	'Index 2 Currency
	Public Property Let HeaderCurrency(c)
		outHeader__(2) = c
	End Property 
	
	'Index 3 Reference
	Public Property Let HeaderReference(r)
		rx__.Pattern = "[0-9]{1,}"
		outHeader__(3) = rx__.Execute(r)(0)
	End Property 
	
	'Index 4 Parma
	Public Property Let HeaderParma(p)
		outHeader__(4) = p
	End Property 
	
	'Index 5 & 7 TotalAmount + TAX amount
	Public Property Let HeaderTotalAmount(n)
		outHeader__(5) = outHeader__(5) + CDbl(n)
		outHeader__(7) = outHeader__(5)
	End Property 
	
	'Index 6 TaxCode
	Public Property Let HeaderTaxCode(t)
		outHeader__(6) = t
	End Property 
	
	'Index 8 LineText
	Public Property Let HeaderLineText(t)
		outHeader__(8) = t
	End Property 
	
	'Index 9 PaymentTerms
	Public Property Let HeaderPaymentTerms(p)
		outHeader__(9) = p
	End Property
	
	'Index 10 TradingPartner
	Public Property Let HeaderTradingPartner(p)
		outHeader__(10) = p
	End Property
	
	'Index 11 AmountInLocCurr
	Public Property Let HeaderAmInLocCur(a)
		outHeader__(11) = a
	End Property
	
	'Index 12 TaxAmount
	Public Property Let HeaderTaxAmount(a)
		outHeader__(12) = a
	End Property
	
	'Sets isInvoice__ and invoice__
	Public Property Let ItemType(t)
		If Left(UCase(t),1) = "I" Then
			invoice__ = t
			isInvoice__ = True
		Else
			invoice__ = t 
			isInvoice__ = False
		End If
	End Property
	
	Public Property Let LineItemDocDate(d)
		rx__.Pattern = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
		outBuffer__ = outBuffer__ & Replace(rx__.Execute(d)(0),"-","") & ";"
	End Property
	
	Public Property Let LineItemCurrency(c)
		outBuffer__ = outBuffer__ & c & ";"
	End Property
	
	Public Property Let LineItemReference(r)
		rx__.Pattern = "[0-9]{1,}"
		outBuffer__ = outBuffer__ & rx__.Execute(r)(0) & ";"
	End Property
	
	Public Property Let LineItemGLAccount(a)
		outBuffer__ = outBuffer__ & a & ";;;"
	End Property 
		
	Public Property Let LineItemTotalAmount(a)
		If isInvoice__ Then
			outBuffer__ = outBuffer__ & a & ";"
		Else
			outBuffer__ = outBuffer__ & a & "-;"
		End If 
	End Property 
	
	Public Property Let LineItemTaxCode(t)
		outBuffer__ = outBuffer__ & t & ";"
	End Property
	
	Public Property Let LineItemTaxAmount(a)
		outBuffer__ = outBuffer__ & a & ";"
	End Property 
	
	Public Property Let LineItemCostCenter(c)
		outBuffer__ = outBuffer__ & c & ";"
	End Property 
	
	Public Property Let LineItemProfitCenter(p)
		outBuffer__ = outBuffer__ & p & ";;;;;;;;;;;"
	End Property 
	
	Public Property Let LineItemAllocation(a)
		outBuffer__ = outBuffer__ & a & ";;"
	End Property 
	
	Public Property Let LineItemTradingPartner(p)
		outBuffer__ = outBuffer__ & p & ";;"
	End Property 
	
	Public Property Let LineItemAmInLocCur(a)
		If isInvoice__ Then
			outBuffer__ = outBuffer__ & a & ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf
		Else
			outBuffer__ = outBuffer__ & a & "-;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf
		End If
	End Property 
	
	
	
	
	
	'Getters
	'Returns header string
	Public Property Get GetHeader
		GetHeader = outHeader__(0) & ";" & outHeader__(1) & ";" & outHeader__(2) & ";" & outHeader__(3) _
		          & ";;" & outHeader__(4) & ";;" & GetTotalAmount & ";" & outHeader__(6) & ";" & GetTotalAmount _
		          & ";;;;;;" & outHeader__(8) & ";;" & outHeader__(9) & ";;;;" & outHeader__(10) & ";;" _
		          & outHeader__(11) & ";" & outHeader__(12) & ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf 
	End Property 
	
	'Return line items string
	Public Property Get GetLineItems
		GetLineItems = outBuffer__
	End Property 
	
	'Returns True if item is invoice, otherwise false
	Public Property Get IsInvoice
		IsInvoice = isInvoice__
	End Property 
	
	Private Property Get GetTotalAmount
		If isInvoice__ Then
			GetTotalAmount = CStr(outHeader__(5)) & "-"
		Else
			GetTotalAmount = CStr(outHeader__(5))
		End If 
	End Property 
	
End Class




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
	
	Function GetListItems(sSite,sList,sQuery)
		With oHTTP
			.open "GET", sSite & "/_api/web/lists/GetByTitle('" & sList & "')/items?$" & sQuery
			.setRequestHeader "accept","application/atom+xml;odata=verbose"
			.setRequestHeader "authorization", "Bearer " & strSecurityToken
			.send
		End With 
		
		If Not oHTTP.status = 200 Then
			GetListItems = Null
			Exit Function
		End If 
		
		GetListItems = oHTTP.responseText
		
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
	
	Public Property Get HttpResponse
		HttpResponse = oHTTP.responseText
	End Property 
End Class 


' DateFormatter Class
Class DateFormatter
	' Convert from YYYY-MM-DD to DD.MM.YYYY
	Public Function FromYyyyMmDd_WithDashesTo_DdMmYyyy_WithDots(strDate)
		Dim temp
		temp = Right(strDate,2) & "." ' Day
		temp = temp & Mid(strDate,6,2) & "." ' Month
		temp = temp & Left(strDate,4) ' Year
		FromYyyyMmDd_WithDashesTo_DdMmYyyy_WithDots = temp
	End Function 
	
	Public Function ToYearMonthDay(D) ' Date object
		ToYearMonthDay = Right("0000" & Year(D),4) & Right("00" & Month(D),2) & Right("00" & Day(D),2)
	End Function
	
	Public Function ToYearMonthDayWithDashes(D)
		ToYearMonthDayWithDashes = Right("0000" & Year(D),4) & "-" & Right("00" & Month(D),2) & "-" & Right("00" & Day(D),2)
	End Function
	
	Public Function ToDayMonthYearWithDots(D)
		ToDayMonthYearWithDots = Right("00" & Day(D),2) & "." & Right("00" & Month(D),2) & "." & Right("0000" & Year(D),4)
	End Function
	
	Public Function ToYearMonthDayHourMinuteSecondWithZeros(D,T)
		ToYearMonthDayHourMinuteSecondWithZeros = Right("0000" & Year(D),4) & Right("00" & Month(D),2) & Right("00" & Day(D),2) & Right("00" & Hour(T),2) & Right("00" & Minute(T),2) & Right("00" & Second(T),2)
	End Function
	
	Public Function ToYearMonthDayHourMinuteWithZeros(D,T)
		ToYearMonthDayHourMinuteWithZeros = Right("0000" & Year(D),4) & Right("00" & Month(D),2) & Right("00" & Day(D),2) & Right("00" & Hour(T),2) & Right("00" & Minute(T),2)
	End Function
	
	
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
  	    		.TextBody = strMessage
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
	    		.TextBody = strMessage
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

