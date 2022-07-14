'#####################################################
'#####################################################
'############## Project name: CZ02_VASI ##############
'############### Script name: ProcessSPList ########## 
'################## Major version: 4 #################
'################## Minor version: 6 #################
'#################### Version: 4.6 ###################
'#####################################################
'#####################################################
'#################### Changelog ######################
'##  27.06.2022										##
'##  SCCoverage implemented							##
'##  04.07.2022										##
'##	 SCCoverage added to all main cases/clauses		##
'##  06.07.2022 									##
'##  CC and PC based on configuration implemented	##
'##  for all conditions except 0 for now			##
'##  07.07.2022										##
'##  Fixed TotalAmount and other fields 			##
'##  in all conditions based on	diagram				##
'##  08.07.2022										##
'##  Changed output file creation. Create file		##
'##  when 1st OK document encountered to prevent	##
'##  empty output files								##
'##  12.07.2022									    ##
'##  Added SharePointLite 4.5. Updates uploaded 	##
'##  file metadata (load4me -> queued)				##
'##  14.07.2022									    ##
'##  Fixed some minor issues.					    ##
'#####################################################
Option Explicit
Const VERSION = "4.6"
Const PROJECT = "CZ02_VASI"
Const UNIT_SOURCE = "unit-vasi"
Const UNIT_DEST = "unit-rc-sk-bs-it"
Const LIBRARY_DEST = "Load4Me"
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
oMailer.AddAdmin = "tomas.ac@volvo.com"
Dim config, country
config = Null
country = Null
Dim credfile, spUserSource, spSecretSource, spSiteSource, spUserDestination, spSecretDestination, spSiteDestination
spSiteSource = Null
spUserSource = Null
spSecretSource = Null
spSiteDestination = Null
spUserDestination = Null
spSecretDestination = Null
credfile = Null
Dim retval, arg, token, xdigestvalue
Dim buffer, entryNode, itemNode, id, report
Dim outFile
Dim oOutFile
Dim bOutFileCreated : bOutFileCreated = False
'#############################################
'################ M A I N ####################
'#############################################

'***************************
'Process cli args and load *
'***************************

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
	oMailer.SendMessage "Configuration folder " & oFSO.GetParentFolderName(config) & " does not exist","E",""
	debug.WriteLine "Configuration folder " & oFSO.GetParentFolderName(config) & " does not exist"
	WScript.Quit
ElseIf Not oFSO.FileExists(config) Then
	'Either file is a folder or file does not exist
	'Notify admin
	oMailer.SendMessage oFSO.GetFileName(config) & " is either a folder or configuration file does not exist","E",""
	debug.WriteLine oFSO.GetFileName(config) & " is either a folder or configuration file does not exist"
	WScript.Quit
End If
'*************************
'@Load config and verify *
'*************************
oXMLCONF.load config
'@Check if country code exists within configuration
If oXMLCONF.selectNodes("//Country[@name=""" & country & """]").length <> 1 Then
	'Missing or duplicate configuration entry
	'Notify amdin
	oMailer.SendMessage "Missing or duplicate configuration entry for country '" & country & "'","E",""
	debug.WriteLine "Missing or duplicate configuration entry for country '" & country & "'"
	WScript.Quit
End If 
credfile = oXMLCONF.selectSingleNode("//Country[@name=""" & country & """]/CredentialsFile").text
'@Can't continue w/o credentials
If IsNull(credfile) Then
	'Notify admin
	oMailer.SendMessage "Unable to load the credentials file","E",""
	WScript.Quit
End If
'************************
'@Load credentials file
'and verify 
'************************
oXMLCONF.load credfile
spUserSource = oXMLCONF.selectSingleNode("//service[@name=""" & UNIT_SOURCE & """]/username").text
spSecretSource = oXMLCONF.selectSingleNode("//service[@name=""" & UNIT_SOURCE & """]/password").text
spSiteSource = oXMLCONF.selectSingleNode("//service[@name=""" & UNIT_SOURCE & """]/host").text
spUserDestination = oXMLCONF.selectSingleNode("//service[@name=""" & UNIT_DEST & """]/username").text
spSecretDestination = oXMLCONF.selectSingleNode("//service[@name=""" & UNIT_DEST & """]/password").text
spSiteDestination = oXMLCONF.selectSingleNode("//service[@name=""" & UNIT_DEST & """]/host").text

If IsNull(spUserSource) Or IsNull(spSecretSource) Or IsNull(spSiteSource) Or IsNull(spUserDestination) Or IsNull(spSecretDestination) Or IsNull(spSiteDestination) Then
	'Notify admin
	oMailer.SendMessage "Sharepoint credentials could not be obtained","E",""
	debug.WriteLine "Sharepoint credentials could not be obtained"
	WScript.Quit
End If 
'********************
'@Initiliaze SPLite *
'********************
retval = oSPL.SharePointLite(spSiteSource,spUserSource,spSecretSource,False)
If retval <> 0 Then
	'Notify admin
	oMailer.SendMessage "Failed to initialize SharePointLite" & vbCrLf & vbCrLf _
	                  & "Error code: " & oSPL.LastErrorNumber & vbCrLf & vbCrLf _
	                  & "Error decription: " & oSPL.LastErrorDesc,"E",""
	debug.WriteLine oSPL.LastErrorNumber
	debug.WriteLine oSPL.LastErrorSource
	debug.WriteLine oSPL.LastErrorDesc
	WScript.Quit
Else 'Initialization OK
	token = oSPL.AccessToken
	xdigestvalue = oSPL.XDigest
End If
'************************************************
'@Load configuration subtree based on country   *
'and create the workind dir if it doesn't exist *
'************************************************
oXMLCONF.load config
Set oXMLCONF = oXMLCONF.selectNodes("//Country[@name=""" & country & """]").item(0)
Dim pathComponents : pathComponents = Split(oXMLCONF.selectSingleNode("//WorkingDirectory").text,"\")
Dim pathComponent
Dim path : path = ""

For pathComponent = 0 To UBound(pathComponents) - 1
	path = path & pathComponents(pathComponent) & "\"
	
	If Not oFSO.FolderExists(path) Then
		oFSO.CreateFolder path
	End If 
Next
'****************************
'@Load 'PostingAP' 'ready'  *
'****************************
buffer = oSPL.GetListItems(HOST,LIST,"select=*&$filter=PostingAP eq 'Ready'")
If IsNull(buffer) Then 
	'http request failed. buffer is null
	oMailer.SendMessage "HTTP request failed" & vbCrLf & vbCrLf & oSPL.LastErrorDesc,"E",""
	WScript.Quit
End If
'**************************************
'@Continue if there are any 'entries' *
'**************************************
oXML.loadXML buffer
oXML.setProperty "SelectionNamespaces", "xmlns:d=""http://schemas.microsoft.com/ado/2007/08/dataservices"""
oXML.setProperty "SelectionNamespaces", "xmlns:m=""http://schemas.microsoft.com/ado/2007/08/dataservices/metadata"""
If oXML.selectNodes("//entry").length = 0 Then
	'nothing to process, no apposting ready items
	oMailer.SendMessage "No PostingAP ready items found","I",""
	WScript.Quit
End If
'***********************************************
'@Consume every //entry node in the collection *
'***********************************************
For Each entryNode In oXML.selectNodes("//entry")
	oMyItems.Consume entryNode, oXMLCONF
Next
'*****************************
'@Create the output csv file *
'*****************************
If oMyItems.CCPCount + oMyItems.OUTCount > 0 Then
	outFile = oXMLCONF.selectSingleNode("//CompanyCode").text & "_A7_AP_CZ02-D_VASI_2d_" & oDF.ToYearMonthDayHourMinuteWithZeros(Date,Time) & ".csv"

	Dim vasiItems : Set vasiItems = oMyItems.ItemsCCP 'Get all instances of MyItem that are of type CCP
	For Each itemNode In vasiItems.Keys
		If vasiItems.Item(itemNode).IsOK Then
			If Not bOutFileCreated Then
				Set oOutFile = oFSO.OpenTextFile(oXMLCONF.selectSingleNode("//WorkingDirectory").text & outFile, 2, True)
				bOutFileCreated = True
			End If 
			oOutFile.Write vasiItems.Item(itemNode).GetHeader & vasiItems.Item(itemNode).GetLineItems
		End If 
	Next
	
	Set vasiItems = oMyItems.ItemsOUT 'Get all instances of MyItem that are of type OUT
	For Each itemNode In vasiItems.Keys
		If vasiItems.Item(itemNode).IsOK Then
			If Not bOutFileCreated Then
				Set oOutFile = oFSO.OpenTextFile(oXMLCONF.selectSingleNode("//WorkingDirectory").text & outFile, 2, True)
				bOutFileCreated = True
			End If 
			oOutFile.Write vasiItems.Item(itemNode).GetHeader & vasiItems.Item(itemNode).GetLineItems
		End If 
	Next
	
	oOutFile.Close
End If


'DEBUG
'WScript.Quit
'DEBUG

'**********************
'@Update list items & *
'build report string  *
'**********************
report = "OUT: " & oMyItems.OUTCount & vbCrLf & "CCP: " & oMyItems.CCPCount & vbCrLf & "OTHER: " & oMyItems.OTHERCount _
       & vbCrLf & "Filename: " & outFile & vbCrLf & "__________________________" & vbCrLf & "Invoice        DocType" & vbCrLf & vbCrLf  

'@OUT
Set vasiItems = oMyItems.ItemsOUT
For Each itemNode In vasiItems.Keys
	debug.WriteLine vasiItems.Item(itemNode).GetInvoice & " " & vasiItems.Item(itemNode).GetDocType & " is OK -> " & CStr(vasiItems.Item(itemNode).IsOK)
	If vasiItems.Item(itemNode).IsOK Then
		For Each id In vasiItems.Item(itemNode).GetIds
			debug.WriteLine "  Patching record with ID: " & id & " Query: " & "{""PostingAP"":""Processed"",""UploadFileAP"":""" & outFile & """,""MatchedCondition"":""" & vasiItems.Item(itemNode).GetItemCondition & """}"
			'oSPL.PatchSingleItem oSPL.SiteUrl,LIST,id,"{""PostingAP"":""Processed"",""UploadFileAP"":""" & outFile & """}"
			oSPL.PatchSingleItem oSPL.SiteUrl,LIST,id,"{""PostingAP"":""Processed"",""UploadFileAP"":""" & outFile & """,""MatchedCondition"":""" & vasiItems.Item(itemNode).GetItemCondition & """}"
		Next
	Else
		report = report & vasiItems.Item(itemNode).GetInvoice & "   (" & vasiItems.Item(itemNode).GetDocType & ")" & vbCrLf 
	End If  
Next

'@CCP
Set vasiItems = oMyItems.ItemsCCP
For Each itemNode In vasiItems.Keys
	debug.WriteLine vasiItems.Item(itemNode).GetInvoice & " " & vasiItems.Item(itemNode).GetDocType & " is OK -> " & CStr(vasiItems.Item(itemNode).IsOK)
	If vasiItems.Item(itemNode).IsOK Then 
		For Each id In vasiItems.Item(itemNode).GetIds
			debug.WriteLine "  Patching record with ID: " & id & " Query: " & "{""PostingAP"":""Processed"",""UploadFileAP"":""" & outFile & """,""MatchedCondition"":""" & vasiItems.Item(itemNode).GetItemCondition & """}"
			oSPL.PatchSingleItem oSPL.SiteUrl,LIST,id,"{""PostingAP"":""Processed"",""UploadFileAP"":""" & outFile & """,""MatchedCondition"":""" & vasiItems.Item(itemNode).GetItemCondition & """}"
		Next
	Else
		report = report & vasiItems.Item(itemNode).GetInvoice & " (" & vasiItems.Item(itemNode).GetDocType & ")" & vbCrLf 
	End If 
Next

'@OTHER
Set vasiItems = oMyItems.ItemsOTHER
For Each itemNode In vasiItems.Keys
	report = report & vasiItems.Item(itemNode).GetInvoice & " (" & vasiItems.Item(itemNode).GetDocType & ")" & vbCrLf 
Next
'***********************************
'@Upload the output .csv file      *
'to the sharepoint library for     *
'***********************************
'reuse SPLite instance
retval = oSPL.SharePointLite(spSiteDestination,spUserDestination,spSecretDestination,False)
If retval <> 0 Then
	'Notify admin
	debug.WriteLine oSPL.LastErrorNumber
	debug.WriteLine oSPL.LastErrorSource
	debug.WriteLine oSPL.LastErrorDesc
	WScript.Quit
Else 'Initialization OK
	token = oSPL.AccessToken
	xdigestvalue = oSPL.XDigest
	oSPL.UploadSingleFile Null, LIBRARY_DEST, oXMLCONF.selectSingleNode("//WorkingDirectory").text & outFile, True, "{""Processed"":""Queued""}"
End If

'********************
'@Send final report *
'********************
report = report & vbCrLf & vbCrLf & spSiteDestination & "/" & LIBRARY_DEST & "/" & outFile & vbCrLf & vbCrLf & "Automatically generated message. Do not respond."
oMailer.SendMessage report,"I",""
                  
'#############################################
'############ M A I N   E N D ################
'#############################################


'Classes/functions
Class MyItems
	
	Private listItemsCCP__
	Private listItemsOUT__
	Private listItemsOTHER__
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
		Set listItemsCCP__ = CreateObject("Scripting.Dictionary")
		Set listItemsOUT__ = CreateObject("Scripting.Dictionary")
		Set listItemsOTHER__ = CreateObject("Scripting.Dictionary")
	End Sub
	
	Private Sub Class_Terminate()
	End Sub 
	
	Public Function Consume(xmlNode__,xmlConfigNode__)
		invoice__ = xmlNode__.selectSingleNode("content").selectSingleNode("m:properties").selectSingleNode("d:Title").text
		Select Case UCase(xmlNode__.SelectSingleNode("content").selectSingleNode("m:properties").selectSingleNode("d:DocType").text)
		
			'###################################################################################################		
			'*************************************** C C C P ***************************************************
			'###################################################################################################	
			Case "CCP"
				
				If Not listItemsCCP__.Exists(invoice__) Then
					listItemsCCP__.Add invoice__, New MyItem
					ccpItems__ = ccpItems__ + 1
				End If
				
				'Item type Invoice/Credit note,sharepoint id,doctype
				listItemsCCP__.Item(invoice__).ItemType = invoice__
				listItemsCCP__.Item(invoice__).Id = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:ID").text
				listItemsCCP__.Item(invoice__).DocType = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocType").text
				'*************************************
				'Header line 
				'************************************* 
				
				'DocDate -> header (1)
				listItemsCCP__.Item(invoice__).HeaderDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'PostingDate -> header (2)
				listItemsCCP__.Item(invoice__).HeaderPostingDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'Currency -> header (3)
				listItemsCCP__.Item(invoice__).HeaderCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
				'Reference -> header (4)
				listItemsCCP__.Item(invoice__).HeaderReference = invoice__
				'GLAccount -> header(5)
				listItemsCCP__.Item(invoice__).HeaderGLAccount = ""
				'Parma -> header (6)
				listItemsCCP__.Item(invoice__).HeaderParma = xmlConfigNode__.SelectSingleNode("//VendorParma").text
				'SpecialGL -> header (7)
				listItemsCCP__.Item(invoice__).HeaderSpecialGL = ""
				'TotalAmount -> header (8)
				listItemsCCP__.Item(invoice__).HeaderTotalAmount = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text
				'TaxCode -> header (9)
				listItemsCCP__.Item(invoice__).HeaderTaxCode = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("TaxCode").text
				'TaxAmount -> header (10)
				listItemsCCP__.Item(invoice__).HeaderTaxAmount = "0,00"
				'CostCenter -> header (11)
				listItemsCCP__.Item(invoice__).HeaderCostCenter = ""
				'ProfitCenter -> header (12)
				listItemsCCP__.Item(invoice__).HeaderProfitCenter = ""
				'Order -> header (13)
				listItemsCCP__.Item(invoice__).HeaderOrder = ""
				'SerialChassiNumber -> header (14)
				listItemsCCP__.Item(invoice__).HeaderSerialChassiNumber = ""
				'ProductVariant -> header (15)
				listItemsCCP__.Item(invoice__).HeaderProductVariant = ""
				'DueDate -> header (16)
				listItemsCCP__.Item(invoice__).HeaderDueDate = ""
				'Quantity -> header (17)
				listItemsCCP__.Item(invoice__).HeaderQuantity = ""
				'LineText -> header (18)
				listItemsCCP__.Item(invoice__).HeaderLineText = ""
				'NumberOfDays -> header (19)
				listItemsCCP__.Item(invoice__).HeaderNumberOfDays = ""
				'PaymentTerms -> header (20)
				listItemsCCP__.Item(invoice__).HeaderPaymentTerms = xmlConfigNode__.SelectSingleNode("//PaymentTerms").text
				'PaymentBlock -> header (21)
				listItemsCCP__.Item(invoice__).HeaderPaymentBlock = ""
				'PaymentMethod -> header (22)
				listItemsCCP__.Item(invoice__).HeaderPaymentMethod = ""
				'Allocation -> header (23)
				listItemsCCP__.Item(invoice__).HeaderAllocation = ""
				'TradingPartner -> header (24)
				listItemsCCP__.Item(invoice__).HeaderTradingPartner = xmlConfigNode__.SelectSingleNode("//TradingPartner").text
				'ExchangeRate -> header (25)
				listItemsCCP__.Item(invoice__).HeaderExchangeRate = ""
				'AmountInLocCurrency -> header (26)
				listItemsCCP__.Item(invoice__).HeaderAmInLocCur = ""
				'TaxAmountInLocCur -> header (27)
				listItemsCCP__.Item(invoice__).HeaderTaxAmountInLocCur = ""
				
				
				'*******************************************
				'GL Lines (must be in order) 
				'*******************************************
				
				'DocDate -> line item (1)
				listItemsCCP__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'PostingDate -> line item (2)
				listItemsCCP__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'Currency -> line item (3)
				listItemsCCP__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
				'Reference -> line items (4)
				listItemsCCP__.Item(invoice__).LineItemReference = invoice__
				'GL -> line item (5)
				listItemsCCP__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("//CaseCCP/GL").text
				'PARMA -> line item (6)
				listItemsCCP__.Item(invoice__).LineItemPARMA = ""
				'SpecialGL -> line item (7)
				listItemsCCP__.Item(invoice__).LineItemSpecialGL = ""
				'TotalAmount -> line item (8)
				listItemsCCP__.Item(invoice__).LineItemTotalAmount = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text
				'TaxCode -> line item (9)
				listItemsCCP__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("TaxCode").text
				'TaxAmount -> line item (10)
				listItemsCCP__.Item(invoice__).LineItemTaxAmount = "0,00"
				'CostCenter -> line item (11)
				listItemsCCP__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("CC").text
				'ProfitCenter -> line item (12)
				listItemsCCP__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseCCP").SelectSingleNode("PC").text
				'Order -> line item (13)
				listItemsCCP__.Item(invoice__).LineItemOrder = ""
				'SerialChassiNumber -> line item (14)
				listItemsCCP__.Item(invoice__).LineItemSerialChassiNumber = ""
				'ProductVariant -> line item (15)
				listItemsCCP__.Item(invoice__).LineItemProductVariant = ""
				'DueDate -> line item (16)
				listItemsCCP__.Item(invoice__).LineItemDueDate = ""
				'Quantity -> line item (17)
				listItemsCCP__.Item(invoice__).LineItemQuantity = ""
				'LineText -> line item (18)
				listItemsCCP__.Item(invoice__).LineItemLineText = ""
				'NumberOfDays -> line item (19)
				listItemsCCP__.Item(invoice__).LineItemNumberOfDays = ""
				'PaymentTerms -> line item (20)
				listItemsCCP__.Item(invoice__).LineItemPaymentTerms = ""
				'PaymentBlock -> line item (21)
				listItemsCCP__.Item(invoice__).LineItemPaymentBlock = ""
				'PaymentMethod -> line item (22)
				listItemsCCP__.Item(invoice__).LineItemPaymentMethod = ""
				'Allocation -> line item (23)
				listItemsCCP__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
				'TradingPartner -> line item (24)
				listItemsCCP__.Item(invoice__).LineItemTradingPartner = ""
				'ExchangeRate -> line item (25)
				listItemsCCP__.Item(invoice__).LineItemExchangeRate = ""
				'AmountInLocCurrency -> line item (26)
				listItemsCCP__.Item(invoice__).LineItemAmInLocCur = ""
				'TaxAmountInLocCur -> line item (27)
				listItemsCCP__.Item(invoice__).LineItemTaxAmountInLocCur = ""
				listItemsCCP__.Item(invoice__).OK = True
				
				
			'###################################################################################################		
			'*************************************** O U T *****************************************************
			'###################################################################################################	
			Case "OUT"
				
				If Not listItemsOUT__.Exists(invoice__) Then
					'Create a new instance of MyItem
					listItemsOUT__.Add invoice__, New MyItem
					outItems__ = outItems__ + 1
				End If
				
				'Item type Invoice/Credit note,sharepoint id,doctype
				listItemsOUT__.Item(invoice__).ItemType = invoice__
				listItemsOUT__.Item(invoice__).Id = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:ID").text
				listItemsOUT__.Item(invoice__).DocType = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocType").text
				'*************************************
				'Header line for case 
				'*************************************
				debug.WriteLine invoice__ & " Header line"
				'DocDate -> header (1)
				listItemsOUT__.Item(invoice__).HeaderDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'PostingDate -> header (2)
				listItemsOUT__.Item(invoice__).HeaderPostingDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
				'Currency -> header (3)
				listItemsOUT__.Item(invoice__).HeaderCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
				'Reference -> header (4)
				listItemsOUT__.Item(invoice__).HeaderReference = invoice__
				'GLAccount -> header(5)
				listItemsOUT__.Item(invoice__).HeaderGLAccount = ""
				'Parma -> header (6)
				listItemsOUT__.Item(invoice__).HeaderParma = xmlConfigNode__.SelectSingleNode("//VendorParma").text
				'SpecialGL -> header (7)
				listItemsOUT__.Item(invoice__).HeaderSpecialGL = ""
				'TotalAmount -> header (8)
				listItemsOUT__.Item(invoice__).HeaderTotalAmount = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:TotalAmount").text
				'TaxAmount -> header (10)
				listItemsOUT__.Item(invoice__).HeaderTaxAmount = "0,00"
				'CostCenter -> header (11)
				listItemsOUT__.Item(invoice__).HeaderCostCenter = ""
				'ProfitCenter -> header (12)
				listItemsOUT__.Item(invoice__).HeaderProfitCenter = ""
				'Order -> header (13)
				listItemsOUT__.Item(invoice__).HeaderOrder = ""
				'SerialChassiNumber -> header (14)
				listItemsOUT__.Item(invoice__).HeaderSerialChassiNumber = ""
				'ProductVariant -> header (15)
				listItemsOUT__.Item(invoice__).HeaderProductVariant = ""
				'DueDate -> header (16)
				listItemsOUT__.Item(invoice__).HeaderDueDate = ""
				'Quantity -> header (17)
				listItemsOUT__.Item(invoice__).HeaderQuantity = ""
				'LineText -> header (18)
				listItemsOUT__.Item(invoice__).HeaderLineText = ""
				'NumberOfDays -> header (19)
				listItemsOUT__.Item(invoice__).HeaderNumberOfDays = ""
				'PaymentTerms -> header (20)
				listItemsOUT__.Item(invoice__).HeaderPaymentTerms = xmlConfigNode__.SelectSingleNode("//PaymentTerms").text
				'PaymentBlock -> header (21)
				listItemsOUT__.Item(invoice__).HeaderPaymentBlock = ""
				'PaymentMethod -> header (22)
				listItemsOUT__.Item(invoice__).HeaderPaymentMethod = ""
				'Allocation -> header (23)
				listItemsOUT__.Item(invoice__).HeaderAllocation = ""
				'TradingPartner -> header (24)
				listItemsOUT__.Item(invoice__).HeaderTradingPartner = xmlConfigNode__.SelectSingleNode("//TradingPartner").text
				'ExchangeRate -> header (25)
				listItemsOUT__.Item(invoice__).HeaderExchangeRate = ""
				'AmountInLocCurrency -> header (26)
				listItemsOUT__.Item(invoice__).HeaderAmInLocCur = ""
				'TaxAmountInLocCur -> header (27)
				listItemsOUT__.Item(invoice__).HeaderTaxAmountInLocCur = ""
				
				'**************************************************************************
				' (0) GL Line1 (optional) may or may not be present.Items must be in order
				'**************************************************************************
				If (CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:TotalAdminCharge").text,".",",")) > 0 And _
				   CBool(xmlConfigNode__.SelectSingleNode("//AdminChargeSeparate").text)) Or _
				   (CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:TotalAdminCharge").text,".",",")) > 0 And _
				   UCase(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DefTempl").text) = "NO") Then
				   debug.WriteLine invoice__ & "  Optional GL Line1"
				   listItemsOUT__.Item(invoice__).ItemCondition = "0"
					'Build GL Line1
					'DocDate -> line item (1)
					listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
					'PostingDate -> line item (2)
					listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
					'Currency -> line item (3)
					listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
					'Reference -> line items (4)
					listItemsOUT__.Item(invoice__).LineItemReference = invoice__
					'GL -> line item (5)
					listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("AdminCharge").SelectSingleNode("GL").text
					'PARMA -> line item (6)
					listItemsOUT__.Item(invoice__).LineItemPARMA = ""
					'SpecialGL -> line item (7)
					listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
					'TotalAmount -> line item (8)
					listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:TotalAdminCharge").text,".",","))
					'TaxCode -> line item (9)
					listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("AdminCharge").SelectSingleNode("TaxCode").text
					'TaxAmount -> line item (10)
					listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
					'CostCenter -> line item (11)
					listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("AdminCharge").SelectSingleNode("CC").text
					'ProfitCenter -> line item (12)
					listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("AdminCharge").SelectSingleNode("PC").text
					'Order -> line item (13)
					listItemsOUT__.Item(invoice__).LineItemOrder = ""
					'SerialChassiNumber -> line item (14)
					listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
					'ProductVariant -> line item (15)
					listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
					'DueDate -> line item (16)
					listItemsOUT__.Item(invoice__).LineItemDueDate = ""
					'Quantity -> line item (17)
					listItemsOUT__.Item(invoice__).LineItemQuantity = ""
					'LineText -> line item (18)
					listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("AdminCharge").SelectSingleNode("LineText").text
					'NumberOfDays -> line item (19)
					listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
					'PaymentTerms -> line item (20)
					listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
					'PaymentBlock -> line item (21)
					listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
					'PaymentMethod -> line item (22)
					listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
					'Allocation -> line item (23)
					listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
					'TradingPartner -> line item (24)
					listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
					'ExchangeRate -> line item (25)
					listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
					'AmountInLocCurrency -> line item (26)
					listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
					'TaxAmountInLocCur -> line item (27)
					listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
				End If
				
				'******************************************************
				' Additional GL Lines, everything after Optional Line1
				'******************************************************
				
				'*************************************************************************************************
				' Condition (1) TotalAdminCharge = 0 & SubTotal = 0 & ChargedForeignVAT = 0 & NonVATChargable > 0
				'*************************************************************************************************
				If CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:TotalAdminCharge").text,".",",")) = 0 And _
				   CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text,".",",")) = 0 And _
				   CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:ChargedForeignVAT").text,".",",")) = 0 And _
				   CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:NonVATChargeable").text,".",",")) > 0 Then
				   debug.WriteLine invoice__ & "  Condition -> TotalAdminCharge = 0 & SubTotal = 0 & ChargedForeignVAT = 0 & NonVATChargeable > 0"
				   listItemsOUT__.Item(invoice__).ItemCondition = "1"
				    '***************************************
				   	' (1a) SCCoverage = No. Produces 1 line
				   	'***************************************
				   	If UCase(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverage").text) = "NO" Then
				   		debug.WriteLine invoice__ & "   Nested condition -> SCCoverage = No"
				   		listItemsOUT__.Item(invoice__).ItemCondition = "a"
					   	'DocDate -> line item (1)
						listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
						'PostingDate -> line item (2)
						listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
						'Currency -> line item (3)
						listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
						'Reference -> line items (4)
						listItemsOUT__.Item(invoice__).LineItemReference = invoice__
						'GL -> line item (5)
						listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("GL").text
						'PARMA -> line item (6)
						listItemsOUT__.Item(invoice__).LineItemPARMA = ""
						'SpecialGL -> line item (7)
						listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
						'TotalAmount -> line item (8)
						listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:NonVATChargeable").text,".",","))
						'TaxCode -> line item (9)
						listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("TaxCode1").text
						'TaxAmount -> line item (10)
						listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
						'CostCenter -> line item (11)
						If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdCC").text) Then
							listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
						Else
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("CC").text
						End If 
						'ProfitCenter -> line item (12)
						If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdPC").text) Then
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
						Else
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("PC").text
						End If 
						'Order -> line item (13)
						listItemsOUT__.Item(invoice__).LineItemOrder = ""
						'SerialChassiNumber -> line item (14)
						listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
						'ProductVariant -> line item (15)
						listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
						'DueDate -> line item (16)
						listItemsOUT__.Item(invoice__).LineItemDueDate = ""
						'Quantity -> line item (17)
						listItemsOUT__.Item(invoice__).LineItemQuantity = ""
						'LineText -> line item (18)
						listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("LineText").text
						'NumberOfDays -> line item (19)
						listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
						'PaymentTerms -> line item (20)
						listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
						'PaymentBlock -> line item (21)
						listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
						'PaymentMethod -> line item (22)
						listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
						'Allocation -> line item (23)
						listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
						'TradingPartner -> line item (24)
						listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
						'ExchangeRate -> line item (25)
						listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
						'AmountInLocCurrency -> line item (26)
						listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
						'TaxAmountInLocCur -> line item (27)
						listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
						listItemsOUT__.Item(invoice__).OK = True
					'****************************************
				   	' (1b) SCCoverage = Full. Produces 1 line
				   	'****************************************
				   	ElseIf UCase(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverage").text) = "FULL" Then
				   		debug.WriteLine invoice__ & "   Nested condition -> SCCoverage = Full"
				   		listItemsOUT__.Item(invoice__).ItemCondition = "b"
				   		'DocDate -> line item (1)
						listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
						'PostingDate -> line item (2)
						listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
						'Currency -> line item (3)
						listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
						'Reference -> line items (4)
						listItemsOUT__.Item(invoice__).LineItemReference = invoice__
						'GL -> line item (5)
						listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("GL").text
						'PARMA -> line item (6)
						listItemsOUT__.Item(invoice__).LineItemPARMA = ""
						'SpecialGL -> line item (7)
						listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
						'TotalAmount -> line item (8)
						listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:NonVATChargeable").text,".",","))
						'TaxCode -> line item (9)
						listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("TaxCode1").text
						'TaxAmount -> line item (10)
						listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
						'CostCenter -> line item (11)
						If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdCC").text) Then
							listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
						Else
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("CC").text
						End If 
						'ProfitCenter -> line item (12)
						If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdPC").text) Then
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
						Else
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("PC").text
						End If 
						'Order -> line item (13)
						listItemsOUT__.Item(invoice__).LineItemOrder = ""
						'SerialChassiNumber -> line item (14)
						listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
						'ProductVariant -> line item (15)
						listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
						'DueDate -> line item (16)
						listItemsOUT__.Item(invoice__).LineItemDueDate = ""
						'Quantity -> line item (17)
						listItemsOUT__.Item(invoice__).LineItemQuantity = ""
						'LineText -> line item (18)
						listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("LineText").text
						'NumberOfDays -> line item (19)
						listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
						'PaymentTerms -> line item (20)
						listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
						'PaymentBlock -> line item (21)
						listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
						'PaymentMethod -> line item (22)
						listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
						'Allocation -> line item (23)
						listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
						'TradingPartner -> line item (24)
						listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
						'ExchangeRate -> line item (25)
						listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
						'AmountInLocCurrency -> line item (26)
						listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
						'TaxAmountInLocCur -> line item (27)
						listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
						listItemsOUT__.Item(invoice__).OK = True
				   	'*********************************************
				   	' (1c) SCCoverage = Partial. Produces 2 lines
				   	'*********************************************
				   	ElseIf UCase(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverage").text) = "PARTIAL" Then
				   		debug.WriteLine invoice__ & "   Nested condition -> SCCoverage = Partial"
				   		listItemsOUT__.Item(invoice__).ItemCondition = "c"
				   		'**** Line 1c1 ****
				   		'DocDate -> line item (1)
						listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
						'PostingDate -> line item (2)
						listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
						'Currency -> line item (3)
						listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
						'Reference -> line items (4)
						listItemsOUT__.Item(invoice__).LineItemReference = invoice__
						'GL -> line item (5)
						listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("GL").text
						'PARMA -> line item (6)
						listItemsOUT__.Item(invoice__).LineItemPARMA = ""
						'SpecialGL -> line item (7)
						listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
						'TotalAmount -> line item (8)
						listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverageAmount").text,".",","))
						'TaxCode -> line item (9)
						listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("TaxCode1").text
						'TaxAmount -> line item (10)
						listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
						'CostCenter -> line item (11)
						If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdCC").text) Then
							listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
						Else
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("CC").text
						End If 
						'ProfitCenter -> line item (12)
						If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdPC").text) Then
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
						Else
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("PC").text
						End If 
						'Order -> line item (13)
						listItemsOUT__.Item(invoice__).LineItemOrder = ""
						'SerialChassiNumber -> line item (14)
						listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
						'ProductVariant -> line item (15)
						listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
						'DueDate -> line item (16)
						listItemsOUT__.Item(invoice__).LineItemDueDate = ""
						'Quantity -> line item (17)
						listItemsOUT__.Item(invoice__).LineItemQuantity = ""
						'LineText -> line item (18)
						listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("LineText").text
						'NumberOfDays -> line item (19)
						listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
						'PaymentTerms -> line item (20)
						listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
						'PaymentBlock -> line item (21)
						listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
						'PaymentMethod -> line item (22)
						listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
						'Allocation -> line item (23)
						listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
						'TradingPartner -> line item (24)
						listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
						'ExchangeRate -> line item (25)
						listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
						'AmountInLocCurrency -> line item (26)
						listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
						'TaxAmountInLocCur -> line item (27)
						listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
						listItemsOUT__.Item(invoice__).OK = True
						
						'**** Line 1c2 ****
						'DocDate -> line item (1)
						listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
						'PostingDate -> line item (2)
						listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
						'Currency -> line item (3)
						listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
						'Reference -> line items (4)
						listItemsOUT__.Item(invoice__).LineItemReference = invoice__
						'GL -> line item (5)
						listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("GL").text
						'PARMA -> line item (6)
						listItemsOUT__.Item(invoice__).LineItemPARMA = ""
						'SpecialGL -> line item (7)
						listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
						'TotalAmount -> line item (8)                                             
						listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:NonVATChargeable").text,".",",")) _
					    												   - CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverageAmount").text,".",","))
						'TaxCode -> line item (9)
						listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("TaxCode1").text
						'TaxAmount -> line item (10)
						listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
						'CostCenter -> line item (11)
						If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdCC").text) Then
							listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
						Else
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("CC").text
						End If 
						'ProfitCenter -> line item (12)
						If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdPC").text) Then
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
						Else
							listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("PC").text
						End If 
						'Order -> line item (13)
						listItemsOUT__.Item(invoice__).LineItemOrder = ""
						'SerialChassiNumber -> line item (14)
						listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
						'ProductVariant -> line item (15)
						listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
						'DueDate -> line item (16)
						listItemsOUT__.Item(invoice__).LineItemDueDate = ""
						'Quantity -> line item (17)
						listItemsOUT__.Item(invoice__).LineItemQuantity = ""
						'LineText -> line item (18)
						listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("LineText").text
						'NumberOfDays -> line item (19)
						listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
						'PaymentTerms -> line item (20)
						listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
						'PaymentBlock -> line item (21)
						listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
						'PaymentMethod -> line item (22)
						listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
						'Allocation -> line item (23)
						listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
						'TradingPartner -> line item (24)
						listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
						'ExchangeRate -> line item (25)
						listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
						'AmountInLocCurrency -> line item (26)
						listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
						'TaxAmountInLocCur -> line item (27)
						listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
						listItemsOUT__.Item(invoice__).OK = True
				   	End If	' SCCoverage EndIf
				ElseIf CBool(xmlConfigNode__.SelectSingleNode("//AdminChargeSeparate").text) = False And _
			       UCase(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DefTempl").text) = "YES" Then
			       debug.WriteLine invoice__ & "  Condition -> AdminChargeSeparate = False & DefTempl = Yes"
			       listItemsOUT__.Item(invoice__).ItemCondition = "2"
			       '**************************************************************************************
				   ' Condition (2) Config.AdminChargeSeparate = False & DefTempl = YES
				   '**************************************************************************************
				   '***********
				   'SCCoverage
				   '***********
				   Select Case UCase(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverage").text)
				   	   '***************************************
				   	   ' (2a) SCCoverage = No. Produces 1 line
				   	   '***************************************
					   Case "NO"
					   		listItemsOUT__.Item(invoice__).ItemCondition = "a"
						    debug.WriteLine invoice__ & "   Nested condition -> SCCoverage = No"	   
							'DocDate -> line item (1)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'PostingDate -> line item (2)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'Currency -> line item (3)
							listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
							'Reference -> line items (4)
							listItemsOUT__.Item(invoice__).LineItemReference = invoice__
							'GL -> line item (5)
							listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("GL").text
							'PARMA -> line item (6)
							listItemsOUT__.Item(invoice__).LineItemPARMA = ""
							'SpecialGL -> line item (7)
							listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
							'TotalAmount -> line item (8)
							listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:TotalAdminCharge").text,".",",")) _
					    												       + CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text,".",","))
							'TaxCode -> line item (9)
							listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("TaxCode1").text
							'TaxAmount -> line item (10)
							listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
							'CostCenter -> line item (11)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdCC").text) Then
								listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("CC").text
							End If 
							'ProfitCenter -> line item (12)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdPC").text) Then
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("PC").text
							End If 
							'Order -> line item (13)
							listItemsOUT__.Item(invoice__).LineItemOrder = ""
							'SerialChassiNumber -> line item (14)
							listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
							'ProductVariant -> line item (15)
							listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
							'DueDate -> line item (16)
							listItemsOUT__.Item(invoice__).LineItemDueDate = ""
							'Quantity -> line item (17)
							listItemsOUT__.Item(invoice__).LineItemQuantity = ""
							'LineText -> line item (18)
							listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("LineText").text
							'NumberOfDays -> line item (19)
							listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
							'PaymentTerms -> line item (20)
							listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
							'PaymentBlock -> line item (21)
							listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
							'PaymentMethod -> line item (22)
							listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
							'Allocation -> line item (23)
							listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
							'TradingPartner -> line item (24)
							listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
							'ExchangeRate -> line item (25)
							listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
							'AmountInLocCurrency -> line item (26)
							listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
							'TaxAmountInLocCur -> line item (27)
							listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
							listItemsOUT__.Item(invoice__).OK = True
						'****************************************
				   	    ' (2b) SCCoverage = Full. Produces 1 line
				   	    '****************************************
						Case "FULL"
							debug.WriteLine invoice__ & "   Nested condition -> SCCoverage = Full"
							listItemsOUT__.Item(invoice__).ItemCondition = "b"
							'DocDate -> line item (1)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'PostingDate -> line item (2)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'Currency -> line item (3)
							listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
							'Reference -> line items (4)
							listItemsOUT__.Item(invoice__).LineItemReference = invoice__
							'GL -> line item (5)
							listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("GL").text
							'PARMA -> line item (6)
							listItemsOUT__.Item(invoice__).LineItemPARMA = ""
							'SpecialGL -> line item (7)
							listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
							'TotalAmount -> line item (8)
							listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:TotalAdminCharge").text,".",",")) _
					    												       + CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text,".",","))
							'TaxCode -> line item (9)
							listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("TaxCode1").text
							'TaxAmount -> line item (10)
							listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
							'CostCenter -> line item (11)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdCC").text) Then
								listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("CC").text
							End If 
							'ProfitCenter -> line item (12)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdPC").text) Then
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("PC").text
							End If 
							'Order -> line item (13)
							listItemsOUT__.Item(invoice__).LineItemOrder = ""
							'SerialChassiNumber -> line item (14)
							listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
							'ProductVariant -> line item (15)
							listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
							'DueDate -> line item (16)
							listItemsOUT__.Item(invoice__).LineItemDueDate = ""
							'Quantity -> line item (17)
							listItemsOUT__.Item(invoice__).LineItemQuantity = ""
							'LineText -> line item (18)
							listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("LineText").text
							'NumberOfDays -> line item (19)
							listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
							'PaymentTerms -> line item (20)
							listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
							'PaymentBlock -> line item (21)
							listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
							'PaymentMethod -> line item (22)
							listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
							'Allocation -> line item (23)
							listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
							'TradingPartner -> line item (24)
							listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
							'ExchangeRate -> line item (25)
							listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
							'AmountInLocCurrency -> line item (26)
							listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
							'TaxAmountInLocCur -> line item (27)
							listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
							listItemsOUT__.Item(invoice__).OK = True
						'*********************************************
				   	    ' (2c) SCCoverage = Partial. Produces 2 lines
				      	'*********************************************
						Case "PARTIAL"
							debug.WriteLine invoice__ & "   Nested condition -> SCCoverage = Partial"
							listItemsOUT__.Item(invoice__).ItemCondition = "c"
					   		'**** Line 2c1 ****
					   		'DocDate -> line item (1)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'PostingDate -> line item (2)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'Currency -> line item (3)
							listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
							'Reference -> line items (4)
							listItemsOUT__.Item(invoice__).LineItemReference = invoice__
							'GL -> line item (5)
							listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("GL").text
							'PARMA -> line item (6)
							listItemsOUT__.Item(invoice__).LineItemPARMA = ""
							'SpecialGL -> line item (7)
							listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
							'TotalAmount -> line item (8)
							listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverageAmount").text,".",","))
							'TaxCode -> line item (9)
							listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("TaxCode1").text
							'TaxAmount -> line item (10)
							listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
							'CostCenter -> line item (11)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdCC").text) Then
								listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("CC").text
							End If 
							'ProfitCenter -> line item (12)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdPC").text) Then
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("PC").text
							End If 
							'Order -> line item (13)
							listItemsOUT__.Item(invoice__).LineItemOrder = ""
							'SerialChassiNumber -> line item (14)
							listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
							'ProductVariant -> line item (15)
							listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
							'DueDate -> line item (16)
							listItemsOUT__.Item(invoice__).LineItemDueDate = ""
							'Quantity -> line item (17)
							listItemsOUT__.Item(invoice__).LineItemQuantity = ""
							'LineText -> line item (18)
							listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("LineText").text
							'NumberOfDays -> line item (19)
							listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
							'PaymentTerms -> line item (20)
							listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
							'PaymentBlock -> line item (21)
							listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
							'PaymentMethod -> line item (22)
							listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
							'Allocation -> line item (23)
							listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
							'TradingPartner -> line item (24)
							listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
							'ExchangeRate -> line item (25)
							listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
							'AmountInLocCurrency -> line item (26)
							listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
							'TaxAmountInLocCur -> line item (27)
							listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
							listItemsOUT__.Item(invoice__).OK = True
							
							'**** Line 2c2 ****
							'DocDate -> line item (1)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'PostingDate -> line item (2)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'Currency -> line item (3)
							listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
							'Reference -> line items (4)
							listItemsOUT__.Item(invoice__).LineItemReference = invoice__
							'GL -> line item (5)
							listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("GL").text
							'PARMA -> line item (6)
							listItemsOUT__.Item(invoice__).LineItemPARMA = ""
							'SpecialGL -> line item (7)
							listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
							'TotalAmount -> line item (8)                                             
							listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:TotalAdminCharge").text,".",",")) _
					    												       + CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text,".",",")) _
						    												   - CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverageAmount").text,".",","))
							'TaxCode -> line item (9)
							listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("TaxCode1").text
							'TaxAmount -> line item (10)
							listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
							'CostCenter -> line item (11)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdCC").text) Then
								listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("CC").text
							End If 
							'ProfitCenter -> line item (12)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdPC").text) Then
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("PC").text
							End If 
							'Order -> line item (13)
							listItemsOUT__.Item(invoice__).LineItemOrder = ""
							'SerialChassiNumber -> line item (14)
							listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
							'ProductVariant -> line item (15)
							listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
							'DueDate -> line item (16)
							listItemsOUT__.Item(invoice__).LineItemDueDate = ""
							'Quantity -> line item (17)
							listItemsOUT__.Item(invoice__).LineItemQuantity = ""
							'LineText -> line item (18)
							listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("LineText").text
							'NumberOfDays -> line item (19)
							listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
							'PaymentTerms -> line item (20)
							listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
							'PaymentBlock -> line item (21)
							listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
							'PaymentMethod -> line item (22)
							listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
							'Allocation -> line item (23)
							listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
							'TradingPartner -> line item (24)
							listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
							'ExchangeRate -> line item (25)
							listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
							'AmountInLocCurrency -> line item (26)
							listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
							'TaxAmountInLocCur -> line item (27)
							listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
							listItemsOUT__.Item(invoice__).OK = True
					End Select 
				ElseIf UCase(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DefTempl").text) = "NO" Then
				'**************************************************************************************
				' Condition (3) DefTempl = No
				'**************************************************************************************
				   debug.WriteLine invoice__ & "  Condition -> DefTempl = No"
				   listItemsOUT__.Item(invoice__).ItemCondition = "3"
				   '***********
				   'SCCoverage
				   '***********
				   Select Case UCase(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverage").text)
				   	   '***************************************
				   	   ' (3a) SCCoverage = No. Produces 1 line
				   	   '***************************************
					   Case "NO"
						    debug.WriteLine invoice__ & "   Nested condition -> SCCoverage = No"
						    listItemsOUT__.Item(invoice__).ItemCondition = "a"   
							'DocDate -> line item (1)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'PostingDate -> line item (2)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'Currency -> line item (3)
							listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
							'Reference -> line items (4)
							listItemsOUT__.Item(invoice__).LineItemReference = invoice__
							'GL -> line item (5)
							listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("GL").text
							'PARMA -> line item (6)
							listItemsOUT__.Item(invoice__).LineItemPARMA = ""
							'SpecialGL -> line item (7)
							listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
							'TotalAmount -> line item (8)
							listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text,".",",")) _
					    												       + CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:ChargedForeignVAT").text,".",","))
							'TaxCode -> line item (9)
							listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("TaxCode2").text
							'TaxAmount -> line item (10)
							listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
							'CostCenter -> line item (11)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdCC").text) Then
								listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("CC").text
							End If 
							'ProfitCenter -> line item (12)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdPC").text) Then
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("PC").text
							End If 
							'Order -> line item (13)
							listItemsOUT__.Item(invoice__).LineItemOrder = ""
							'SerialChassiNumber -> line item (14)
							listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
							'ProductVariant -> line item (15)
							listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
							'DueDate -> line item (16)
							listItemsOUT__.Item(invoice__).LineItemDueDate = ""
							'Quantity -> line item (17)
							listItemsOUT__.Item(invoice__).LineItemQuantity = ""
							'LineText -> line item (18)
							listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("LineText").text
							'NumberOfDays -> line item (19)
							listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
							'PaymentTerms -> line item (20)
							listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
							'PaymentBlock -> line item (21)
							listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
							'PaymentMethod -> line item (22)
							listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
							'Allocation -> line item (23)
							listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
							'TradingPartner -> line item (24)
							listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
							'ExchangeRate -> line item (25)
							listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
							'AmountInLocCurrency -> line item (26)
							listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
							'TaxAmountInLocCur -> line item (27)
							listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
							listItemsOUT__.Item(invoice__).OK = True
						'****************************************
				   	    ' (3b) SCCoverage = Full. Produces 1 line
				   	    '****************************************
						Case "FULL"
							debug.WriteLine invoice__ & "   Nested condition -> SCCoverage = Full"
							listItemsOUT__.Item(invoice__).ItemCondition = "b"
							'DocDate -> line item (1)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'PostingDate -> line item (2)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'Currency -> line item (3)
							listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
							'Reference -> line items (4)
							listItemsOUT__.Item(invoice__).LineItemReference = invoice__
							'GL -> line item (5)
							listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("GL").text
							'PARMA -> line item (6)
							listItemsOUT__.Item(invoice__).LineItemPARMA = ""
							'SpecialGL -> line item (7)
							listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
							'TotalAmount -> line item (8)
							listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text,".",",")) _
					    												       + CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:ChargedForeignVAT").text,".",","))
							'TaxCode -> line item (9)
							listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("TaxCode2").text
							'TaxAmount -> line item (10)
							listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
							'CostCenter -> line item (11)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdCC").text) Then
								listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("CC").text
							End If 
							'ProfitCenter -> line item (12)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdPC").text) Then
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("PC").text
							End If 
							'Order -> line item (13)
							listItemsOUT__.Item(invoice__).LineItemOrder = ""
							'SerialChassiNumber -> line item (14)
							listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
							'ProductVariant -> line item (15)
							listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
							'DueDate -> line item (16)
							listItemsOUT__.Item(invoice__).LineItemDueDate = ""
							'Quantity -> line item (17)
							listItemsOUT__.Item(invoice__).LineItemQuantity = ""
							'LineText -> line item (18)
							listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("LineText").text
							'NumberOfDays -> line item (19)
							listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
							'PaymentTerms -> line item (20)
							listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
							'PaymentBlock -> line item (21)
							listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
							'PaymentMethod -> line item (22)
							listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
							'Allocation -> line item (23)
							listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
							'TradingPartner -> line item (24)
							listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
							'ExchangeRate -> line item (25)
							listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
							'AmountInLocCurrency -> line item (26)
							listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
							'TaxAmountInLocCur -> line item (27)
							listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
							listItemsOUT__.Item(invoice__).OK = True
						'*********************************************
				   	    ' (3c) SCCoverage = Partial. Produces 2 lines
				      	'*********************************************
						Case "PARTIAL"
							debug.WriteLine invoice__ & "   Nested condition -> SCCoverage = Partial"
							listItemsOUT__.Item(invoice__).ItemCondition = "c"
					   		'**** Line 3c1 ****
					   		'DocDate -> line item (1)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'PostingDate -> line item (2)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'Currency -> line item (3)
							listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
							'Reference -> line items (4)
							listItemsOUT__.Item(invoice__).LineItemReference = invoice__
							'GL -> line item (5)
							listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("GL").text
							'PARMA -> line item (6)
							listItemsOUT__.Item(invoice__).LineItemPARMA = ""
							'SpecialGL -> line item (7)
							listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
							'TotalAmount -> line item (8)
							listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverageAmount").text,".",","))
							'TaxCode -> line item (9)
							listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("TaxCode2").text
							'TaxAmount -> line item (10)
							listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
							'CostCenter -> line item (11)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdCC").text) Then
								listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("CC").text
							End If 
							'ProfitCenter -> line item (12)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("DealerIdPC").text) Then
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("PC").text
							End If 
							'Order -> line item (13)
							listItemsOUT__.Item(invoice__).LineItemOrder = ""
							'SerialChassiNumber -> line item (14)
							listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
							'ProductVariant -> line item (15)
							listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
							'DueDate -> line item (16)
							listItemsOUT__.Item(invoice__).LineItemDueDate = ""
							'Quantity -> line item (17)
							listItemsOUT__.Item(invoice__).LineItemQuantity = ""
							'LineText -> line item (18)
							listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS2").SelectSingleNode("LineText").text
							'NumberOfDays -> line item (19)
							listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
							'PaymentTerms -> line item (20)
							listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
							'PaymentBlock -> line item (21)
							listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
							'PaymentMethod -> line item (22)
							listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
							'Allocation -> line item (23)
							listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
							'TradingPartner -> line item (24)
							listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
							'ExchangeRate -> line item (25)
							listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
							'AmountInLocCurrency -> line item (26)
							listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
							'TaxAmountInLocCur -> line item (27)
							listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
							listItemsOUT__.Item(invoice__).OK = True
							
							'**** Line 3c2 ****
							'DocDate -> line item (1)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'PostingDate -> line item (2)
							listItemsOUT__.Item(invoice__).LineItemDocDate = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocDate").text
							'Currency -> line item (3)
							listItemsOUT__.Item(invoice__).LineItemCurrency = xmlConfigNode__.SelectSingleNode("//Currency").text
							'Reference -> line items (4)
							listItemsOUT__.Item(invoice__).LineItemReference = invoice__
							'GL -> line item (5)
							listItemsOUT__.Item(invoice__).LineItemGLAccount = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("GL").text
							'PARMA -> line item (6)
							listItemsOUT__.Item(invoice__).LineItemPARMA = ""
							'SpecialGL -> line item (7)
							listItemsOUT__.Item(invoice__).LineItemSpecialGL = ""
							'TotalAmount -> line item (8)                                             
							listItemsOUT__.Item(invoice__).LineItemTotalAmount = CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:ChargedForeignVAT").text,".",",")) _
					    												       + CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SubTotal").text,".",",")) _
						    												   - CDbl(Replace(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:SCCoverageAmount").text,".",","))
							'TaxCode -> line item (9)
							listItemsOUT__.Item(invoice__).LineItemTaxCode = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("TaxCode2").text
							'TaxAmount -> line item (10)
							listItemsOUT__.Item(invoice__).LineItemTaxAmount = "0,00"
							'CostCenter -> line item (11)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdCC").text) Then
								listItemsOUT__.Item(invoice__).LineItemCostCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("CC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("CC").text
							End If 
							'ProfitCenter -> line item (12)
							If CBool(xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("DealerIdPC").text) Then
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("DealerIdMatrix").SelectSingleNode(xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DealerId").text).attributes.getNamedItem("PC").text
							Else
								listItemsOUT__.Item(invoice__).LineItemProfitCenter = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("PC").text
							End If 
							'Order -> line item (13)
							listItemsOUT__.Item(invoice__).LineItemOrder = ""
							'SerialChassiNumber -> line item (14)
							listItemsOUT__.Item(invoice__).LineItemSerialChassiNumber = ""
							'ProductVariant -> line item (15)
							listItemsOUT__.Item(invoice__).LineItemProductVariant = ""
							'DueDate -> line item (16)
							listItemsOUT__.Item(invoice__).LineItemDueDate = ""
							'Quantity -> line item (17)
							listItemsOUT__.Item(invoice__).LineItemQuantity = ""
							'LineText -> line item (18)
							listItemsOUT__.Item(invoice__).LineItemLineText = xmlConfigNode__.SelectSingleNode("CaseOUT").SelectSingleNode("VAS1").SelectSingleNode("LineText").text
							'NumberOfDays -> line item (19)
							listItemsOUT__.Item(invoice__).LineItemNumberOfDays = ""
							'PaymentTerms -> line item (20)
							listItemsOUT__.Item(invoice__).LineItemPaymentTerms = ""
							'PaymentBlock -> line item (21)
							listItemsOUT__.Item(invoice__).LineItemPaymentBlock = ""
							'PaymentMethod -> line item (22)
							listItemsOUT__.Item(invoice__).LineItemPaymentMethod = ""
							'Allocation -> line item (23)
							listItemsOUT__.Item(invoice__).LineItemAllocation = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:RequestNumber").text
							'TradingPartner -> line item (24)
							listItemsOUT__.Item(invoice__).LineItemTradingPartner = ""
							'ExchangeRate -> line item (25)
							listItemsOUT__.Item(invoice__).LineItemExchangeRate = ""
							'AmountInLocCurrency -> line item (26)
							listItemsOUT__.Item(invoice__).LineItemAmInLocCur = ""
							'TaxAmountInLocCur -> line item (27)
							listItemsOUT__.Item(invoice__).LineItemTaxAmountInLocCur = ""
							listItemsOUT__.Item(invoice__).OK = True
					End Select 
				Else
					debug.WriteLine "Else Case"
				   'No match. Mark this and report it. Do not process such Item i.e do not append to the out.csv file
				   listItemsOUT__.Item(invoice__).OK = False 
				End If  
			
			Case Else
				debug.WriteLine "Else Case"
				otherItems__ = otherItems__ + 1
				
				If Not listItemsOTHER__.Exists(invoice__) Then
					listItemsOTHER__.Add invoice__, New MyItem
				End If
				
				'Item type Invoice/Credit note,sharepoint id,doctype
				listItemsOTHER__.Item(invoice__).ItemType = invoice__
				listItemsOTHER__.Item(invoice__).Id = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:ID").text
				listItemsOTHER__.Item(invoice__).DocType = xmlNode__.SelectSingleNode("content").SelectSingleNode("m:properties").SelectSingleNode("d:DocType").text 
			
		End Select 
		
	End Function
	
	
	'*************************
	'** P r o p e r t i e s **
	'*************************
	Public Property Get CCPCount
		CCPCount = listItemsCCP__.Count
	End Property 
	
	Public Property Get OUTCount
		OUTCount = listItemsOUT__.Count
	End Property
	
	Public Property Get OTHERCount
		OTHERCount = listItemsOTHER__.Count
	End Property 
	
	Public Property Get Count
		Count = listItemsCCP__.Count + listItemsOUT__.Count
	End Property 
	
	Public Property Get ItemsCCP
		Set ItemsCCP = listItemsCCP__
	End Property
	
	Public Property Get ItemsOUT
		Set ItemsOUT = listItemsOUT__
	End Property 
	
	Public Property Get ItemsOTHER
		Set ItemsOTHER = listItemsOTHER__
	End property
	 
End Class

'This class represents one invoice
Class MyItem
	
	Public 	outHeader__(27) ' Header line array
	'0 -> skip
	'1 -> DocumentDate; 2 -> PostingDate; 3 -> Currency; 4 -> Reference; 5 -> GLAccount; 6 -> PARMA; 7 -> SpecialG/L; 8 -> AmountDocumentCurrency
	'9 -> TAXCode; 10 -> TaxAmount; 11 -> CostCenter; 12 -> ProfitCenter; 13 -> Order; 14 -> Serial/ChassiNumber; 15 -> ProductVariant; 16 -> DueDate
	'17 -> Quantity; 18 -> LineText; 19 -> NumberOfDays; 20 -> PaymentTerms; 21 -> PaymentBlock; 22 -> PaymentMethod; 23 -> Allocation; 24 -> TradingPartner
	'25 -> ExchangeRate; 26 -> AmountInLocCur; 27 -> TaxAmountInLocCur
	
	Private outBuffer__ 	' GL lines string delimited with CRLF
	Private invoice__ 		' invoice string
	Private isInvoice__     ' Invoice or credit note. If true -> invoice else credite note
	Private rx__
	Private ok__            ' Initially False, set to true by Consume method indicating this invoice fits the conditions specified for GL lines
	Private ids__			' collection of sharepoint IDs associated with each invoice. CCP has multiple IDs OUT usually only one
	Private type__			' This field holds info about invoice type (CCP,OUT or other). For now we consider only CCP and OUT types
	Private conditionsMatched__ ' This field is for testing purposes. Concatenated conditions e.g 0 | 01 | 02 | 03 - items w/o condition or items only w/ condition 0 are not considered valid
	
	Private Sub Class_Initialize()
		Set ids__ = CreateObject("Scripting.Dictionary")
		Set rx__ = New RegExp
		rx__.Global = True
		ok__ = False 
		outBuffer__ = ""
		type__ = ""
	End Sub
	
	'***************
	'Properties
	'***************
	
	'Setters
	'Index 1 DocDate
	Public Property Let HeaderDocdate(d)
		rx__.Pattern = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
		outHeader__(1) = Replace(rx__.Execute(d)(0),"-","")
	End Property
	
	'Index 2 PostingDate
	Public Property Let HeaderPostingDate(d)
		rx__.Pattern = "[0-9]{4}-[0-9]{2}-[0-9]{2}"
		outHeader__(2) = Replace(rx__.Execute(d)(0),"-","")
	End Property
	
	'Index 3 Currency
	Public Property Let HeaderCurrency(c)
		outHeader__(3) = c
	End Property 
	
	'Index 4 Reference
	Public Property Let HeaderReference(r)
'		rx__.Pattern = "[0-9]{1,}"
'		outHeader__(3) = rx__.Execute(r)(0)
		outHeader__(4) = r
	End Property 
	
	'Index 5 GLAccount
	Public Property Let HeaderGLAccount(a)
		outHeader__(5) = a
	End Property 
	
	'Index 6 Parma
	Public Property Let HeaderParma(p)
		outHeader__(6) = p
	End Property 
	
	'Index 7 SpecialG/L
	Public Property Let HeaderSpecialGL(g)
		outHeader__(7) = g
	End Property 
	
	'Index 8 TotalAmount
	Public Property Let HeaderTotalAmount(n)
		outHeader__(8) = outHeader__(8) + CDbl(Replace(n,".",","))
	End Property 
	
	'Index 9 TaxCode
	Public Property Let HeaderTaxCode(t)
		outHeader__(9) = t
	End Property 
	
	'Index 10 TaxAmount
	Public Property Let HeaderTaxAmount(t)
		outHeader__(10) = t
	End Property
	
	'Index 11 CostCenter
	Public Property Let HeaderCostCenter(c)
		outHeader__(11) = c
	End Property 
	
	'Index 12 ProfitCenter
	Public Property Let HeaderProfitCenter(p)
		outHeader__(12) = p
	End Property
	
	'Index 13 Order
	Public Property Let HeaderOrder(o)
		outHeader__(13) = o
	End Property 
	
	'Index 14 Serial/ChassiNumber
	Public Property Let HeaderSerialChassiNumber(s)
		outHeader__(14) = s
	End Property
	
	'Index 15 ProductVariant
	Public Property Let HeaderProductVariant(v)
		outHeader__(15) = v
	End Property
	
	'Index 16 DueDate
	Public Property Let HeaderDueDate(d)
		outHeader__(16) = d
	End Property 
	
	'Index 17 Quantity
	Public Property Let HeaderQuantity(q)
		outHeader__(17) = q
	End Property
	
	'Index 18 LineText
	Public Property Let HeaderLineText(l)
		outHeader__(18) = l
	End Property
	
	'Index 19 NumberOfDays
	Public Property Let HeaderNumberOfDays(n)
		outHeader__(19) = n
	End Property
	
	'Index 20 PaymentTerms
	Public Property Let HeaderPaymentTerms(p)
		outHeader__(20) = p
	End Property
	
	'Index 21 PaymentBlock
	Public Property Let HeaderPaymentBlock(b)
		outHeader__(21) = b
	End Property 
	
	'Index 22 PaymentMethod
	Public Property Let HeaderPaymentMethod(m)
		outHeader__(22) = m
	End Property
	
	'Index 23 Allocation
	Public Property Let HeaderAllocation(a)
		outHeader__(23) = a
	End Property
	
	'Index 24 TradingPartner
	Public Property Let HeaderTradingPartner(p)
		outHeader__(24) = p
	End Property
	
	'Index 25 ExchangeRate
	Public Property Let HeaderExchangeRate(r)
		outHeader__(25) = r
	End Property 
	
	'Index 26 AmountInLocCurr
	Public Property Let HeaderAmInLocCur(a)
		outHeader__(26) = a
	End Property
	
	'Index 27 TaxAmount
	Public Property Let HeaderTaxAmountInLocCur(a)
		outHeader__(27) = a
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
'		rx__.Pattern = "[0-9]{1,}"
'		outBuffer__ = outBuffer__ & rx__.Execute(r)(0) & ";"
		outBuffer__ = outBuffer__ & r & ";"
	End Property
	
	Public Property Let LineItemGLAccount(a)
		outBuffer__ = outBuffer__ & a & ";"
	End Property
	
	Public Property Let LineItemSpecialGL(g)
		outBuffer__ = outBuffer__ & g & ";"
	End Property 
	
	Public Property Let LineItemPARMA(p)
		outBuffer__ = outBuffer__ & p & ";"
	End Property
	
	Public Property Let LineItemTotalAmount(a)
		
		If isInvoice__ Then
			If InStr(a,",") > 0 Then
				outBuffer__ = outBuffer__ & Replace(a,",","") & ";"
			Else
				outBuffer__ = outBuffer__ & a & "00;"
			End If 
		Else
			If InStr(a,",") > 0 Then
				outBuffer__ = outBuffer__ & Replace(a,",","") & "-;"
			Else
				outBuffer__ = outBuffer__ & a & "00-;"
			End If 
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
		outBuffer__ = outBuffer__ & p & ";"
	End Property
	
	Public Property Let LineItemOrder(o)
		outBuffer__ = outBuffer__ & o & ";"
	End Property
	
	Public Property Let LineItemSerialChassiNumber(s)
		outBuffer__ = outBuffer__ & s & ";"
	End Property
	
	Public Property Let LineItemProductVariant(v)
		outBuffer__ = outBuffer__ & v & ";"
	End Property
	
	Public Property Let LineItemDueDate(d)
		outBuffer__ = outBuffer__ & d & ";"
	End Property 
	 
	Public Property Let LineItemQuantity(q)
		outBuffer__ = outBuffer__ & q & ";"
	End Property
	
	Public Property Let LineItemLineText(t)
		outBuffer__ = outBuffer__ & t & ";"
	End property
	
	Public Property Let LineItemNumberOfDays(d)
		outBuffer__ = outBuffer__ & d & ";"
	End Property
	
	Public Property Let LineItemPaymentTerms(t)
		outBuffer__ = outBuffer__ & t & ";"
	End Property 
	
	Public Property Let LineItemPaymentBlock(b)
		outBuffer__ = outBuffer__ & b & ";"
	End Property
	
	Public Property Let LineItemPaymentMethod(m)
		outBuffer__ = outBuffer__ & m & ";"
	End Property 
	
	Public Property Let LineItemAllocation(a)
		outBuffer__ = outBuffer__ & a & ";;;"
	End Property 
	
	Public Property Let LineItemTradingPartner(p)
		outBuffer__ = outBuffer__ & p & ";"
	End Property 
	
	Public Property Let LineItemExchangeRate(e)
		outBuffer__ = outBuffer__ & e & ";"
	End Property
	
	Public Property Let LineItemAmInLocCur(a)
		outBuffer__ = outBuffer__ & a & ";" 
	End Property 
	
	Public Property Let LineItemTaxAmountInLocCur(a)
		outBuffer__ = outBuffer__ & a & ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf
	End Property
	
	Public Property Let OK(b)
		ok__ = b
	End Property
	
	Public Property Let DocType(t)
		type__ = t
	End Property 
	
	Public Property Let Id(i)
		ids__.Add i,""
	End Property
	
	Public Property Let ItemCondition(c)
		conditionsMatched__ = conditionsMatched__ & c
	End Property 
	
	
	'Getters
	Public Property Get GetDocType
		GetDocType = type__
	End Property 
	
	Public Property Get GetIds
		GetIds = ids__.Keys
	End Property 
	
	'Returns header string
	Public Property Get GetHeader
		GetHeader = outHeader__(1) & ";" & outHeader__(2) & ";" & outHeader__(3) & ";" & outHeader__(4) & ";" & outHeader__(5) _
		          & ";" & outHeader__(6) & ";" & outHeader__(7) & ";" & GetHeaderTotalAmount & ";" & outHeader__(9) & ";" & outHeader__(10) _
		          & ";" & outHeader__(11) & ";" & outHeader__(12) & ";" & outHeader__(13) & ";" & outHeader__(14) & ";" & outHeader__(15) _
		          & ";" & outHeader__(16) & ";" & outHeader__(17) & ";" & outHeader__(18) & ";" & outHeader__(19) & ";" & outHeader__(20) _
		          & ";" & outHeader__(21) & ";" & outHeader__(22) & ";" & outHeader__(23) & ";" & outHeader__(24) & ";" & outHeader__(25) _
		          & ";" & outHeader__(26) & ";" & outHeader__(27) & ";;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;;" & vbCrLf
	End Property 
	
	'Return line items string
	Public Property Get GetLineItems
		GetLineItems = outBuffer__
	End Property 
	
	'Returns True if item is invoice, otherwise false
	Public Property Get IsInvoice
		IsInvoice = isInvoice__
	End Property 
	
	Private Property Get GetHeaderTotalAmount
		If isInvoice__ Then
			If InStr(CStr(outHeader__(8)),",") > 0 Then
				GetHeaderTotalAmount = Replace(CStr(outHeader__(8)),",","") & "-"
			Else
				GetHeaderTotalAmount = CStr(outHeader__(8)) & "00-"
			End If 
		Else
			If InStr(CStr(outHeader__(8)),",") > 0 Then
				GetHeaderTotalAmount = Replace(CStr(outHeader__(8)),",","")
			Else
				GetHeaderTotalAmount = CStr(outHeader__(8)) & "00"
			End If 
		End If 
	End Property
	
	Public Property Get IsOK
		IsOK = ok__
	End Property
	
	Public Property Get GetInvoice
		GetInvoice = invoice__
	End Property 
	
	Public Property Get GetItemCondition
		GetItemCondition = conditionsMatched__
	End property
End Class


Class SharePointLite
	
	Private oRX 
	Private oXML
	Private oSTREAM
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
		Set oSTREAM = CreateObject("ADODB.Stream")
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
	
	
	
	Function PatchSingleItem(sSite,sList,sId,sJson)
		'"{""PostingAP"":""Processed"",""UploadFileAP"":""FILE_NAME""}"
		
		With oHTTP
			.open "PATCH", sSite & "_api/web/lists/GetByTitle('" & sList & "')/items(" & sId & ")", False
			.setRequestHeader "Accept","application/json;odata=verbose"
			.setRequestHeader "Content-Type","application/json"
			.setRequestHeader "Authorization","Bearer " & strSecurityToken
			.setRequestHeader "If-Match","*"
			.setRequestHeader "Content-Length", Len(sJson)
			.send sJson
		End With
		
		PatchSingleItem = oHTTP.status
	End Function
	
	Function UploadSingleFile(sSiteUrl,sLibraryName,sFilePath,bOverwrite,sJsonMetadata)
	'						  ^^^^^^^^ ^^^^^^^^^^^^ ^^^^^^^^^ ^^^^^^^^^^ 
	'						  Site      Library      File      Overwrite 
		Dim tmp,file,buffer
		oRX.Pattern = "\/sites\/.+\/"
		
		'Build library relative
		If IsNull(sSiteUrl) Then 'Use site url stored within instance
			sLibraryName = oRX.Execute(strSiteURL)(0) & sLibraryName
			sSiteUrl = strSiteURL
		Else
			If Not Right(sSiteUrl,1) = "/" Then
				sSiteUrl = sSiteUrl & "/"
			End If 
			
			tmp = oRX.Execute(sSiteUrl)(0)
			sLibraryName = tmp & sLibraryName
		End If 
		
	
		
		oSTREAM.Open
		oSTREAM.LoadFromFile sFilePath
		oSTREAM.Type = 2
		oSTREAM.Charset = "utf-8"
		oSTREAM.Position = 2
		
		oRX.Pattern = "[\w-]+\.[a-zA-Z]*"
		tmp = oRX.Execute(sFilePath)(0) ' file name		
		
		With oHTTP
			.open "POST", sSiteUrl & "_api/web/GetFolderByServerRelativeUrl('" & sLibraryName & "')/Files/add(url='" & tmp & "',overwrite=" & LCase(CStr(bOverwrite)) & ")", False 
			.setRequestHeader "Authorization", "Bearer " & AccessToken
			'.setRequestHeader "Content-Type", "application/octet-stream"
			.setRequestHeader "Content-Type", "text/csv"
			.setRequestHeader "X-RequestDigest", XDigest
			.setRequestHeader "Content-Length", oSTREAM.Size - 2
			.send oSTREAM.ReadText(oSTREAM.Size)
		End With 	
		
		
		'.open "GET", sSite & "_api/" & "Web/GetFileByServerRelativePath(decodedurl='/sites/unit-rc-sk-bs-it/Load4Me/test.csv')/listItemAllFields", False
		If Not IsNull(sJsonMetadata) Or sJsonMetadata <> "" Then ' Update metadata for this file
			
			oXML.loadXML oHTTP.responseText 'Load response returned when file is created
			
			With oHTTP
				.open "GET", oXML.selectSingleNode("entry").selectSingleNode("id").text & "/listItemAllFields", False ' Get item ID. Not returned when file is created
				.setRequestHeader "Authorization", "Bearer " & strSecurityToken
				.send
			End With 
			
			oXML.loadXML oHTTP.responseText
			
			With oHTTP
				.open "PATCH", oXML.selectSingleNode("entry").attributes.getNamedItem("xml:base").text & oXML.selectSingleNode("//entry/link[@rel=""edit""]").attributes.getNamedItem("href").text, False
				.setRequestHeader "Accept","application/json;odata=verbose"
				.setRequestHeader "Content-Type","application/json"
				.setRequestHeader "Authorization","Bearer " & strSecurityToken
				.setRequestHeader "If-Match","*"
				.setRequestHeader "Content-Length", Len(sJsonMetadata)
				.send sJsonMetadata
			End With
			
			UploadSingleFile = oHTTP.status
			
			
		Else
		
			UploadSingleFile = oHTTP.status
			
		End If 
		
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
	
	'Properties
	Public Property Get SiteUrl
		SiteUrl = strSiteURL
	End Property 
	
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

