<%@ Language=VBScript %>
<%	Option Explicit %>
<!--#include virtual="/_PROJECT.asp" -->
<%
	'- Head -'
	Page_Head "Account.DplExport"
	Dim oSecurityState: Set oSecurityState = m_oInstinct.Load("SecurityState")
	Dim oDbase: Set oDbase = m_oInstinct.Load("DBase")
	Dim oExport: Set oExport = m_oInstinct.Load("Export")

	'+
	Dim cSendDplId: cSendDplId = ""
	Dim dDPL_Calc_Markup
	Dim bDPL_Calc_CSMD
	Dim rAccountContact: Set rAccountContact = oDbase("").Interface("AccountContact", "Item~n=Key", oSecurityState.AccountContactKey, Null)
	If (rAccountContact.Eof = False) Then
		cSendDplId = LCase(rAccountContact("SendDplId") & "")
		bDPL_Calc_CSMD = rAccountContact("IsDPL_CostCalc_CSMD")
		If rAccountContact("DPL_CostCalc_Markup") & vbnullstring = vbnullstring Then
			dDPL_Calc_Markup = 1
		Else
			dDPL_Calc_Markup = Cdbl(rAccountContact("DPL_CostCalc_Markup")) / 100
		End If
	End If
	rAccountContact.Close: Set rAccountContact = Nothing
	'+
	Dim cDplStoreId : cDplStoreId = ""
	Dim cRedemptionSoftware : cRedemptionSoftware = ""
	Dim cDplExclusions : cDplExclusions = ""
	Dim rAccount : Set rAccount = oDbase("").Interface("Account", "Item~n=Key", oSecurityState.AccountKey, Null)
	If (rAccount.Eof = False) Then
		cDplExclusions = rAccount("DplExclusions") & ""
		cDplStoreId = rAccount("DplStoreId") & ""
		cRedemptionSoftware = rAccount("RedemptionSoftware") & ""
	End If
	rAccount.Close() : Set rAccount = Nothing
	cDplExclusions = RTrim(cDplExclusions)
	'+
	Dim r: Set r = oDbase("").Interface("OrderHistory", "Hash:DplReport~n=AccountContactKey,x=OrderHistoryKey", oSecurityState.AccountContactKey, Request.QueryString("OrderHistoryId"))
	'+ filename
	Dim cFilename: cFilename = "Invoice.DPL"
	If (r.Eof = False) Then
		if(cSendDplId = "embedepl") then 
			cFilename = r("InvoiceId") & ".EPL" 
		elseif(cSendDplId = "semnox") then 
			cFilename = r("InvoiceId") & ".CSV"
		elseif(cSendDplId = "foodtrak" or cSendDplId = "ideal") then 
			cFilename = "DPL" & r("InvoiceId") & ".DPL"
		else 			
			cFilename = r("InvoiceId") & ".DPL" 
		end if
		'cFilename = m_oInstinct.Iif(cSendDplId = "ideal", "DPL", "") & r("InvoiceId") & ".DPL"
		'cFilename = m_oInstinct.Iif(cSendDplId = "foodtrak", "DPL", "") & r("InvoiceId") & ".DPL"
		'cFilename = m_oInstinct.Iif(cSendDplId = "embedepl", "EPL", "") & r("InvoiceId") & ".EPL"
	End If
	'+ override for SACOA RedemptionSoftware types
	If (cRedemptionSoftware = "SACOA") Then
		cSendDplId = "sacoa"
	End If
	If (cRedemptionSoftware = "Sacoa Smart Scan") Then
		cSendDplId = "sacoa smart scan"
	End If
	If (cRedemptionSoftware = "Intercard") Then
		cSendDplId = "intercard"
	End If
	'+ content type
	Response.ContentType = "application/csv"
	Response.AddHeader "Content-Disposition", "attachment; filename=" & cFilename
	If ((r.Eof = False) And ((cSendDplId <> "") And (cSendDplId <> "none"))) Then
		'*****************************************************************
		'+ NOTE: KEEP IN SYNC WITH DPL OUTPUT IN /Agent/InvoiceNotify.aspx
		'+ STARTING HERE >>
		'*****************************************************************
		Dim sDPLData
		
		'+ field header
		Select Case (cSendDplId)
		Case "fcs"
			sDPLData = "FCSDPL,R,,," & r("ShippingPhone") & "," & r("OrderId")
		Case "core"
			sDPLData = "RP," & r("CustomerId") & "," & r("OrderId")
		Case "ideal"
			sDPLData = "DPL" & r("OrderId") & ",R,,," & r("CustomerId") & "," & r("OrderId")
		Case "foodtrak"
			sDPLData = ""
		Case "sacoa", "sacoa smart scan", "centeredge"
			sDPLData = cDplStoreId & "," & r("OrderId")
		Case "semnox"
			sDPLData = "DPLFILERI," & cDplStoreId & ",Redemption Plus," & r("InvoiceId") & "," & r("InvoiceDate") & ",,"		
		Case Else
			sDPLData = Replace(Replace(Replace(Replace(r("ShippingPhone") & "", "-", ""), "(", ""), ")", ""), " ", "") & "," & r("OrderId")
		End Select
		
		'Tayler Duncan EPL update Start
		if(cSendDplId = "embedepl") then 
			Dim header 
			header = Join(Array("Barcode","Product_Code","Product_Name","Unit_Type","Item_Quantity","Ticket_Value","Items_Per_Unit","Item_Cost","Category_Id"), ",")
			sDPLData = sDPLData & vbCrLf & header	
		end if
		'Tayler Duncan EPL update End
		
		Dim cKey
		Dim oSkuMap : Set oSkuMap = m_oInstinct.Load_Hash
		Dim rSkuMap : Set rSkuMap = oDbase("").Interface("AccountSkuOverride", "Hash:View~n=OrderHistoryKey", Request.QueryString("OrderHistoryId"), Null)
		Do While (rSkuMap.Eof = False)
			cKey = rSkuMap("Id") & ""
			If (Not oSkuMap.Exists(cKey)) Then
				oSkuMap.Add cKey, rSkuMap("Sku") & ""
			End If
			rSkuMap.MoveNext
		Loop
		rSkuMap.Close() : Set rSkuMap = Nothing
		
		Dim oCurHash : Set oCurHash = m_oInstinct.Load_Hash
		Dim rCur : Set rCur = oDbase("").Interface("InvoiceNotify", "Hash.Item:DplCurrencyExchange~x=OrderId", Null, r("InvoiceId"))
		Do While (rCur.Eof = False)
			cKey = rCur("Sku") & "" & rCur("Unit")
			If ((Not oCurHash.Exists(cKey)) And (rCur("CurrencyExchange") & "" <> "")) Then
				oCurHash.Add cKey, rCur("CurrencyExchange") & ""
			End If
			rCur.MoveNext
		Loop
		rCur.Close() : Set rCur = Nothing
		
		'+ field data
		Do While (r.Eof = False)
			Dim cItemNumber : cItemNumber = r("ItemNumber") & ""
			If (oSkuMap.Exists(cItemNumber)) Then
				cItemNumber = oSkuMap(cItemNumber)
			End If
			If (InStr(cDplExclusions, r("ProductStatusId")) <= 0) Then
				'JMW: Added variables to make field logic easier to read
				Dim cFieldData: cFieldData = ""
				Dim iConvFactor : iConvFactor = r("UnitConversionFactor")
				Dim iQty : iQty = r("Quantity") * iConvFactor
				Dim dUnitPrice : dUnitPrice = m_oInstinct.Cur(r("Price"), 0) * dDPL_Calc_Markup
				Dim dPriceEa : dPriceEa = dUnitPrice / iConvFactor
				Dim iTickets : iTickets = FormatString(r("CurrencyExchange") & vbNullString, "0")
				Dim sUPC : sUPC = r("Upc") & vbNullString
				Dim iKitPcs : iKitPcs = r("KitQtyPcs")
				Dim sStatusId : sStatusId = UCase(r("ProductStatusId"))
				Dim sUnitOfMeas : sUnitOfMeas = UCase(GetDplUnitOfMeasure(cSendDplId, r("UnitOfMeasure")))
				Dim sCategory : sCategory = ""'SanitizeForSacoaDpl(r("Category"))
				Dim sSubCategory : sSubCategory = ""'SanitizeForSacoaDpl(r("SubCategory"))
				Dim sTheme : sTheme = r("theme")
				
				cKey = cItemNumber & REPLACE(UCASE(sUnitOfMeas), "INNER", "INNR")
				If (oCurHash.Exists(cKey)) Then
					iTickets = m_oInstinct.Lng(oCurHash(cKey), 0)
				End If
				
				Dim iConvFactorDisplay : iConvFactorDisplay = iConvFactor
				'JMW: Added kit pricing for CS/MD items sold as "EACH" if the preference is set on AccountContact
				If iKitPcs > 1 And sUnitOfMeas = "EACH" Then
					'Force UOM CASE
					sUnitOfMeas = "CASE"
					dPriceEa = dUnitPrice / iKitPcs
					iQty = iQty * iKitPcs
					iTickets = m_oInstinct.Lng(iTickets, 0)
					'Force Pcs/unit = KitQtyPcs
					iConvFactorDisplay = iKitPcs
				End If
				
				'JMW: Insert new line only if DPL already contains data
				If sDPLData <> "" Then sDPLData = sDPLData & vbCrLf

				'JMW: "na" and "n/a" values cause problems with certain vendors
				IF LCase(sUPC) = "na" OR LCase(sUPC) = "n/a" THEN sUPC = ""

				Select Case (cSendDplId)
				Case "fcs"
					iQty = (iQty / iConvFactor)
					cFieldData = Join(Array(cItemNumber & "", SanitizeForDpl(r("ItemDescription") & ""), sUnitOfMeas, iQty, FormatString(dUnitPrice, "0.00"), iTickets, "3", iConvFactorDisplay, FormatString(dPriceEa, "0.00")), ",")
				Case "ideal"
					iQty = (iQty / iConvFactor)
					cFieldData = Join(Array(cItemNumber & "", SanitizeForDpl(r("ItemDescription") & ""), sUnitOfMeas, iQty, FormatString(dUnitPrice, "0.00"), iTickets, "3", iConvFactorDisplay, FormatString(dPriceEa, "0.00"), "", "", "", "", sUPC & ""), ",")
				Case "foodtrak"
					cFieldData = Join(Array(r("CustomerId") & "", r("OrderId") & "", r("InvoiceDate") & "", cItemNumber & "", FormatString(iQty, "0"), FormatString(dPriceEa*iQty, "0.00")), ",")
				Case "mainevent"
					cFieldData = Join(Array(cItemNumber & "", SanitizeForDpl(r("ItemDescription") & ""), sUnitOfMeas, iQty, FormatString(dUnitPrice, "0.0000"), iTickets, iConvFactorDisplay, FormatString(dPriceEa, "0.0000")), ",")
				Case "core"
					cFieldData = Join(Array(cItemNumber & "", SanitizeForDpl(r("ItemDescription") & ""), sUnitOfMeas, iQty, FormatString(dUnitPrice, "0.00"), iTickets, iConvFactorDisplay, FormatString(dPriceEa, "0.00"), sUPC), ",")
				case "sacoa smart scan"
					cFieldData = Join(Array(m_oInstinct.Iif(sUPC & "" = "", "000" & cItemNumber, sUPC), cItemNumber & "", iTickets, SanitizeForSacoaDpl(r("ItemDescription") & ""), SanitizeForSacoaDpl(sCategory), SanitizeForSacoaDpl(sSubCategory), FormatString(dUnitPrice, "0.00"), FormatString(dPriceEa, "0.00"), iConvFactor, iQty / iConvFactor, iQty), ",")
				case "intercard"
					cFieldData = Join(Array(cItemNumber & "", SanitizeForDpl(r("ItemDescription") & ""), sUnitOfMeas, iQty, FormatString(dUnitPrice, "0.00"), iTickets, iConvFactorDisplay, FormatString(dPriceEa, "0.00"), m_oInstinct.Iif(sUPC & "" = "", "000" & cItemNumber, sUPC)), ",")
				case "embedepl"																								
					cFieldData = Join(Array(m_oInstinct.Iif(sUPC & "" = "", "" & cItemNumber, sUPC), "" & cItemNumber, SanitizeForDpl(r("ItemDescription") & ""), sUnitOfMeas, iQty, iTickets, iConvFactorDisplay, FormatString(dPriceEa, "0.00"), "" & sTheme), ",") 
				case "centeredge"
					cFieldData = Join(Array(m_oInstinct.Iif(sUPC & "" = "", "" & cItemNumber, sUPC), SanitizeForDpl(r("ItemDescription") & ""), sUnitOfMeas, iQty, FormatString(dUnitPrice, "0.00"), iTickets, iConvFactorDisplay, FormatString(dPriceEa, "0.00"),"", cItemNumber), ",")
				case "Standard-UPC"
					cFieldData = Join(Array(m_oInstinct.Iif(sUPC & "" = "", "" & cItemNumber, sUPC), cItemNumber & "", SanitizeForDpl(r("ItemDescription") & ""), sUnitOfMeas, iQty, FormatString(dUnitPrice, "0.00"), iTickets, iConvFactorDisplay, FormatString(dPriceEa, "0.00")), ",")
				case "semnox"
					cFieldData = Join(Array(m_oInstinct.Iif(sUPC & "" = "", "" & cItemNumber, sUPC), SanitizeForDpl(r("ItemDescription") & ""), r("Quantity"), sUnitOfMeas, iConvFactorDisplay, FormatString(dPriceEa, "0.00"), iTickets), ",")
				Case Else
					cFieldData = Join(Array(cItemNumber & "", SanitizeForDpl(r("ItemDescription") & ""), sUnitOfMeas, iQty, FormatString(dUnitPrice, "0.00"), iTickets, iConvFactorDisplay, FormatString(dPriceEa, "0.00")), ",")
				End Select				
				sDPLData = sDPLData & cFieldData
				
			End If
			r.MoveNext
		Loop
		r.Close: Set r = Nothing
		'*****************************************************************
		'+ NOTE: KEEP IN SYNC WITH DPL OUTPUT IN /Agent/InvoiceNotify.aspx
		'+ << ENDING HERE
		'*****************************************************************
		Response.Write sDPLData	
		Response.End
	End If

	Function GetDplUnitOfMeasure(ByVal cSendDplId, ByVal cUnitOfMeasure)
		If ((cSendDplId & "" <> "") And (cUnitOfMeasure & "" <> "")) Then
			Select Case (LCase(cSendDplId))
			Case "fcs", "ideal"
				If (LCase(cUnitOfMeasure) = "inner") Then
					GetDplUnitOfMeasure = "INNR"
					Exit Function
				End If
			Case Else
				If (LCase(cUnitOfMeasure) = "innr") Then
					GetDplUnitOfMeasure = "INNER"
					Exit Function
				End If
			End Select
		End If
		GetDplUnitOfMeasure = UCase(cUnitOfMeasure & "")
	End Function

	Function SanitizeForDpl(ByRef cValue)
		If (cValue & "" = "") Then
			SanitizeForDpl = ""
			Exit Function
		End If

		Dim cReturnValue : cReturnValue = cValue
		Dim i : For i = 1 To Len(cReturnValue)
			If (i > Len(cReturnValue)) Then
				Exit For
			End If
			Dim iChar : iChar = asc(mid(cReturnValue, i, 1))
			If ((iChar < 32 Or iChar > 122) And (iChar <> 10) And (iChar <> 13)) Then
				i = i - 1
				cReturnValue = Left(cReturnValue, i) & Right(cReturnValue, Len(cReturnValue) - i - 1)
			End If
		Next
		SanitizeForDpl = Replace(Replace(cReturnValue, "&", "and"), """", "in")
	End Function
	
	Function SanitizeForSacoaDpl(ByRef cValue)
		SanitizeForSacoaDpl = Replace(Replace(Replace(Replace(SanitizeForDpl(cValue), ",", ""), "'", ""), "-", ""), "/", "")
	End Function
	
	'- FormatString -'
	Function FormatString(ByVal c, ByVal cFormat)
		FormatString = ""
		c = c & ""
		If (c = "") Then
			Exit Function
		End If
		Select Case (cFormat)
		Case "0.00"
			If (IsNumeric(c) = True) Then
				FormatString = FormatNumber(c, 2, 0) & ""
			End If
		Case "0.0000"
			If (IsNumeric(c) = True) Then
				FormatString = FormatNumber(c, 4, 0) & ""
			End If
		Case "0"
			If (IsNumeric(c) = True) Then
				FormatString = Fix(c) & ""
			End If
		End Select
	End Function

	'- Foot -'
	m_oInstinct.Free oExport
	m_oInstinct.Free oDbase
	m_oInstinct.Free oSecurityState
	Page_Foot
%>