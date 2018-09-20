<%
	'- Index_Process -'
	Public Function Index_Process()
		Index_Process = False
		Dim cRefArray: cRefArray = Split(m_oForm.Ref & "||", "|", 3)
		Select Case (m_oForm.Action)
		Case "Update"
			Dim cProductCategoryId : cProductCategoryId = m_oInstinct.Str(cRefArray(0), "")
			m_oInstinct.Redirect CreateProductUrl(cProductCategoryId, m_oForm.Request("ProductSubCategoryId"), Null, Null, m_oForm.Request("Gender"), m_oForm.Request("ProductAgeReferenceId"))
		Case "AddToCartBackorder"
			m_oCart.AddToBackorderCart()
		Case "DeleteBackorderStudentItem"
			m_oCart.DeleteBackorderStudentItem
		Case "SetCustomAge"
			m_oSession("CustomProductAge") = m_oForm.Request("CustomAge")
		Case Else
			Index_Process = True
		End Select
		cProductCategoryId = m_oInstinct.Str(cRefArray(0), "")
		Select Case (m_oForm.Task)
		Case "Search"
			m_oPage.Site.Frame.MetaTitle = "Search Results"
		Case "OnSale"
			m_oPage.Site.Frame.MetaTitle = "Redemption and Incentive Sale Items"
			m_oPage.Site.Frame.MetaKeyword = "redemption toys, toys and novelties, crane toys, redemption prizes, redemption merchandise, crane plush toys"
			m_oPage.Site.Frame.MetaDescription = "Redemption toys and novelties that are on sale."
		Case "New"
			m_oPage.Site.Frame.MetaTitle = "New Redemption and Incentive Products"
			m_oPage.Site.Frame.MetaKeyword = "redemption toys, toys and novelties, crane toys, redemption prizes, redemption merchandise, crane plush toys"
			m_oPage.Site.Frame.MetaDescription = "The latest redemption toys and novelties, crane toys, and plush novelties."
		Case "HotProduct"
			m_oPage.Site.Frame.MetaTitle = "Top Sellers in Redemption and Incentive Merchandise"
			m_oPage.Site.Frame.MetaKeyword = "redemption toys, toys and novelties, crane toys, redemption prizes, redemption merchandise, crane plush toys"
			m_oPage.Site.Frame.MetaDescription = "Top selling redemption prizes, redemption toys and novelties and crane toys."
		Case Else
			Dim rProductCategory : Set rProductCategory = m_oDbase("").Interface("ProductCategory", "Item~x=Id", Null, m_oInstinct.Str(m_oInstinct.Str(Request.QueryString("SubId"), m_oForm.Request("ProductSubCategoryId")), cProductCategoryId))
			If (rProductCategory.Eof = False) Then
				m_oPage.Site.Frame.MetaTitle = m_oInstinct.Str(rProductCategory("MetaTitle"), rProductCategory("Name").Value)
				m_oPage.Site.Frame.MetaKeyword = m_oInstinct.Str(rProductCategory("MetaKeyword"), rProductCategory("Name").Value)
				m_oPage.Site.Frame.MetaDescription = m_oInstinct.Str(rProductCategory("MetaDescription"), "")
			End If
			rProductCategory.Close() : Set rProductCategory = Nothing
		End Select
	End Function

	'- Index_Render -'
	Public Sub Index_Render(ByRef oHtml)
		Dim bIsLoggedIn : bIsLoggedIn = m_oInstinct.Iif(m_oInstinct.Item("SecurityState").AccountKey > 0, True, False)
		With oHtml
		x_ oHtml, "Head"
		x_Error oHtml, ""
		Select Case (m_oForm.Task)
		Case "OnSale"
			.x_ "<h1 style=""margin:0;padding:0;""><img src=""/Product/SaleItems.jpg"" alt=""Redemption Merchandise on Sale"" /></h1>"
		Case "New"
			.x_ "<h1 style=""margin:0;padding:0;""><img src=""/Product/NewProducts.jpg"" alt=""New Redemption Merchandise"" /></h1>"
		Case "HotProduct"
			.x_ "<h1 style=""margin:0;padding:0;""><img src=""/Product/TopSellers.jpg"" alt=""Top Selling Redemption Merchandise"" /></h1>"
			.x_ "<br /><span style=""font-size:12px;color:#000;"">These are the <font size=""2"">top selling items</font> from the previous day.  This list is updated each day so you’ll always be up to date on what our top movers are.</p>"
			.x_ "<br /><br /><a href=""/Catalog/Product-Spotlight/"">Click here to view a list of redemption and incentive merchandise trends.</a></span>"
		End Select
		Dim cDescription : cDescription = ""
		Dim cRefArray: cRefArray = Split(m_oForm.Ref & "||", "|", 3)
		Dim cProductCategoryId: cProductCategoryId = ""
		Dim cProductSubCategoryId: cProductSubCategoryId = m_oForm.Field("ProductSubCategoryId").Text
		Dim nProductSubCategoryKey : nProductSubCategoryKey = ""
		Dim nProductCategoryKey : nProductCategoryKey = 0
		If (m_oForm.Task <> "Search") Then
			cProductCategoryId = m_oInstinct.Str(cRefArray(0), "")
		End If
		'+
		If (cProductCategoryId & "" <> "") Then
			Dim cProductCategoryName: cProductCategoryName = ""
			Dim rProductCategory: Set rProductCategory = m_oDbase("").Interface("ProductCategory", "Item:View~x=Id|SubCategoryId", Null, cProductCategoryId & "|" & cProductSubCategoryId)
			Dim bIsAdminOnlyCategory : bIsAdminOnlyCategory = False
			If (rProductCategory.Eof = False) Then
				bIsAdminOnlyCategory = m_oInstinct.Bit(rProductCategory("AdminOnly"), False)
				nProductCategoryKey = m_oInstinct.Lng(rProductCategory("Key"), 0)
				cProductCategoryName = rProductCategory("Name")

'JMW 24Sep2012: Remove Grey filter boxed as per Sales & Marketing
'				Dim bIsHideCategoryFilters: bIsHideCategoryFilters = m_oInstinct.Bit(rProductCategory("IsHideCategoryFilters"), False)
'				If (bIsHideCategoryFilters = False) Then
'		.o_Tr ""
'			.o_Td Null, ""
'				m_oFrame.Render oHtml, "GrayEngineHeadHead"
'				m_oFrame.Render oHtml, "GrayEngineHeadAction"
'				If (rProductCategory("HeaderImage_O") & "" <> "") Then
'				.x_Img Config_FileLibrary & "ProductCategory/" & nProductCategoryKey & "/" & rProductCategory("HeaderImage_O"), ""
'				Else
'				.x_ cProductCategoryName
'				End If
'				m_oFrame.Render oHtml, "GrayEngineHeadFoot"
'				m_oFrame.Render oHtml, "EngineBodyHead"
'				.o_Tr ""
'					.o_Td Null, "align=center"
'						.x_Br
'						.x_ "Narrow the list of products in the """ & cProductCategoryName & """ category with 3 easy clicks:": .x_Br
'						.x_Br
'						.o_Table "cellpadding=5"
'						.o_Tr ""
							'JMW 26Sep2012: Even though we don't show the category filters we still need to track the sub-category parameter for paging.
							'+ subcategory
							Dim cProductSubCategory: cProductSubCategory = ""
							Dim rProductSubCategory: Set rProductSubCategory = m_oDbase("").Interface("ProductCategory", "Hash:UsedView~n=Key,x=AccountKey|ProductGenderId|ProductAgeReferenceId", nProductCategoryKey, m_oSecurityState.AccountKey & "|" & m_oForm.Field("Gender").Text & "|" & m_oForm.Field("ProductAgeReferenceId").Text)
							If (rProductSubCategory.Eof = False) Then
'							.o_Td Null, ""
'								.x_ m_oForm.Field("ProductSubCategoryId").Render_Label: .x_Br
'								.x_ m_oForm.Field("ProductSubCategoryId").Render_Input
'								m_oForm.Option_Add Null, "-- all subcategories --"
'								Do While (rProductSubCategory.Eof = False)
'								m_oForm.Option_Add rProductSubCategory("Id"), rProductSubCategory("Name")
'								rProductSubCategory.MoveNext
'								Loop
'								.x_ m_oForm.Field("ProductSubCategoryId").Render_Special
								If (cProductSubCategoryId <> "") Then
									rProductSubCategory.Filter = "(Id = '" & Replace(cProductSubCategoryId, "'", "''") & "')"
									If (rProductSubCategory.Eof = False) Then
										nProductSubCategoryKey = m_oInstinct.Lng(rProductSubCategory("Key"), 0)
										cProductSubCategory = rProductSubCategory("Name")
									End If
								End If
							End If
							rProductSubCategory.Close: Set rProductSubCategory = Nothing
'							'+ gender
'							Dim cGender: cGender = ""
'							Dim rProductGender: Set rProductGender = m_oDbase("").Interface("ProductGender", "Hash:UsedView~n=ProductCategoryKey,x=ProductAgeReferenceId", m_oInstinct.Lng(nProductSubCategoryKey, nProductCategoryKey), m_oForm.Field("ProductAgeReferenceId").Text)
'							If (rProductGender.Eof = False) Then
'							.o_Td Null, ""
'								.x_ m_oForm.Field("Gender").Render_Label: .x_Br
'								.x_ m_oForm.Field("Gender").Render_Input
'								m_oForm.Option_Add Null, "-- all genders --"
'								Do While (rProductGender.Eof = False)
'								m_oForm.Option_Add rProductGender("Id"), rProductGender("Name")
'								rProductGender.MoveNext
'								Loop
'								.x_ m_oForm.Field("Gender").Render_Special
'								If (m_oForm.Field("Gender").Text <> "") Then
'									rProductGender.Filter = "(Id = '" & Replace(m_oForm.Field("Gender").Value, "'", "''") & "')"
'									If (rProductGender.Eof = False) Then
'										cGender = rProductGender("Name")
'									End If
'								End If
'							End If
'							rProductGender.Close: Set rProductGender = Nothing
'							'+ age
'							Dim cProductAge: cProductAge = ""
'							Dim rProductAge: Set rProductAge = m_oDbase("").Interface("ProductAge", "Hash:UsedView~n=ProductCategoryKey,x=ProductGenderId", m_oInstinct.Lng(nProductSubCategoryKey, nProductCategoryKey), m_oForm.Field("Gender").Text)
'							If (rProductAge.Eof = False) Then
'							.o_Td Null, ""
'								.x_ m_oForm.Field("ProductAgeReferenceId").Render_Label: .x_Br
'								.x_ m_oForm.Field("ProductAgeReferenceId").Render_Input
'								m_oForm.Option_Add Null, "-- all ages --"
'								Do While (rProductAge.Eof = False)
'								m_oForm.Option_Add rProductAge("ReferenceId"), rProductAge("Name")
'								rProductAge.MoveNext
'								Loop
'								.x_ m_oForm.Field("ProductAgeReferenceId").Render_Special
'								If (m_oForm.Field("ProductAgeReferenceId").Text <> "") Then
'									rProductAge.Filter = "(ReferenceId = '" & Replace(m_oForm.Field("ProductAgeReferenceId").Value, "'", "''") & "')"
'									If (rProductAge.Eof = False) Then
'										cProductAge = rProductAge("Name")
'									End If
'								End If
'							End If
'							rProductAge.Close: Set rProductAge = Nothing
'						.x_Table: .Table = "TD"
'				m_oFrame.Render oHtml, "EngineBodyFoot"
'				m_oFrame.Render oHtml, "EngineFoot"
'				.x_Br
'				Else
				.x_ m_oForm.Field("ProductSubCategoryId").Render_Hidden
'				.x_ m_oForm.Field("Gender").Render_Hidden
'				.x_ m_oForm.Field("ProductAgeReferenceId").Render_Hidden
'				End If
			End If
			rProductCategory.Close: Set rProductCategory = Nothing
		End If
		'+
		If (bIsAdminOnlyCategory = True And m_oSecurityState.SecurityAccountKey = 0) Then
			x_ oHtml, "Foot"
			Exit Sub
		End If

		'+
		cDescription = m_oDbase("").Str_("ProductCategory", "Calc:Description~n=Key", m_oInstinct.Lng(nProductSubCategoryKey, nProductCategoryKey), Null, "Description", "")
		If (cDescription & "" <> "") Then
			.x_ cDescription & "<br/><br/>"
		End If
		If (m_oForm.Task = "New") Then
			Dim oNewContent : Set oNewContent = m_oInstinct.Load("Engine_Content")
			oNewContent.Reference = "NewProducts"
			If (Not(oNewContent.Item("Block") Is Nothing)) Then
				.x_ oNewContent.Item("Block").Attrib("Body")
			End If
			m_oInstinct.Free oNewContent
			'+
			.o_Div "ageFilter", "style=text-align:center;"
				.x_ m_oForm.Field("CustomAge").Render_Label
				.x_ m_oForm.Field("CustomAge").Render_Input
					m_oForm.Option_Add "15", "Last 15 Days"
					m_oForm.Option_Add "30", "Last 30 Days"
					m_oForm.Option_Add "45", "Last 45 Days"
					m_oForm.Option_Add "60", "Last 60 Days"
				.x_ m_oForm.Field("CustomAge").Render_Special
			.x_Div
			.x_ "<scr" & "ipt type=""text/javascript"">" & vbCrLf
			.x_ "$(document).ready(function() {" & vbCrLf
			.x_ "	$('#CustomAge').bind('change', function() {" & vbCrLf
			.x_ "		var form = document.Form_Engine;" & vbCrLf
			.x_ "		form.Action.value = 'SetCustomAge';" & vbCrLf
			.x_ "		form.submit();" & vbCrLf
			.x_ "	});" & vbCrLf
			.x_ "});" & vbCrLf
			.x_ "</scr" & "ipt>" & vbCrLf
		End If
		'+
		Dim cRequestHash: Set cRequestHash = m_oInstinct.Load_Hash
		If (m_nCartAdd_Quantity = 0) Then
			Dim hRequestForm: For Each hRequestForm In Request.Form
				If (Left(hRequestForm, 9) = "Quantity_") Then
					cRequestHash(Mid(hRequestForm, 10)) = m_oInstinct.Lng(Request.Form(hRequestForm), Null)
				ElseIf (Left(hRequestForm, 8) = "Balance_") Then
					cRequestHash("B" & Mid(hRequestForm, 9)) = m_oInstinct.Lng(Request.Form(hRequestForm), Null)
				End If
			Next
		End If
		Dim cResultType
		Select Case (m_oForm.Task)
'		Case "SkuSearch"
'			cResultType = "SkuSearch"
		Case "Backorder"
			cResultType = "Backorder"
		Case "HotProduct"
			cResultType = "Index"
			m_oSelectPager.AddTH "HotProductRank", "Rank", "hidden"
		Case "Search"
			cResultType = "Index"
			m_oSelectPager.AddTH "SearchRank", "Rank", "hidden"
		Case "New", "OnSale"
			m_oSelectPager.AddTH "CreateDate", "CreateDate", "hidden"
			cResultType = "Index"
		Case Else
			cResultType = "Index"
		End Select
		m_oSelectPager.AddTH "IsNew", "", "nosort"
		m_oSelectPager.AddTH Null, "", "nosort"
		m_oSelectPager.AddTH Null, "", "nosort"
		m_oSelectPager.AddTH "Name", "name", ""
		m_oSelectPager.AddTH "Id", "SKU", "align=center"
		If (bIsLoggedIn = True) Then
		m_oSelectPager.AddTH "InStockSequence", "in stock", "align=center;reversesort;"
		End If
		If (m_oProduct.AccountCurrency <> "") Then
		m_oSelectPager.AddTH "CurrencyExchange", LCase(m_oProduct.AccountCurrency), ""
		End If
		If (bIsLoggedIn = True) Then
			m_oSelectPager.AddTH "EachUnitPrice", "each price", "align=center"
		End If
		m_oSelectPager.AddTH "Quantity", "qty/units", "nosort;align=right"
		If (bIsLoggedIn = True) Then
			m_oSelectPager.AddTH "EachPrice", "unit price", "nosort"
			m_oSelectPager.AddTH Null, "qty", "nosort;align=center"
		End If

		Dim cSubheader
		Dim cText : cText = Left(Trim(m_oInstinct.Str(m_oSession("Search.Keyword"), "")), 300)
		Select Case (m_oForm.Task)
		Case "Search"
			cSubheader = "SearchResults.gif"
			m_oSelectPager.AddSearchAttribute "AccountKey", m_oSecurityState.AccountKey
			If ((Len(cText) = 6) And (m_oInstinct.Lng(cText, 0) <> 0)) Then
				m_oSelectPager.AddSearchAttribute "Sku", cText
				m_oSelectPager.ExecuteContract "ProductGroup", "Hash:SkuSearch~cXml"
			Else
				m_oSelectPager.AddSearchAttribute "PriceFrom", m_oInstinct.Str(m_oSession("Search.PriceFrom"), "")
				m_oSelectPager.AddSearchAttribute "PriceTo", m_oInstinct.Str(m_oSession("Search.PriceTo"), "")
				m_oSelectPager.AddSearchAttribute "PriceType", m_oInstinct.Str(m_oSession("Search.PriceType"), "")
				m_oSelectPager.AddSearchAttribute "IsActive", "1"
				m_oSelectPager.AddSearchAttribute "KeyRummage_Request", cRefArray(1)
				m_oSelectPager.AddSearchAttribute "Method", "Hash:ViewMatch~n=KeyRummage_Request,x=AccountKey|PriceFrom|PriceTo|PriceType|IsActive"
				Select Case (m_oSelectPager.SortField)
				Case "Name", "Id"
					m_oSelectPager.SortPrefix = "ProductGroup"
				End Select
				m_oSelectPager.ExecuteContract "ProductGroup", "Hash:Search~cXml"
			End If

		Case Else
			m_oSelectPager.AddSearchAttribute "Task", m_oForm.Task
			m_oSelectPager.AddSearchAttribute "AccountKey", m_oSecurityState.AccountKey
			Select Case (m_oForm.Task)
			Case "SkuSearch"
				cSubheader = "SearchResults.gif"
				m_oSelectPager.AddSearchAttribute "Sku", cText
			Case "Index"
				cSubheader = "Product.gif"
				m_oSelectPager.AddSearchAttribute "ProductCategoryId", m_oInstinct.Str(m_oForm.Field("ProductSubCategoryId").Value, cProductCategoryId)
				m_oSelectPager.AddSearchAttribute "ProductGenderId", m_oForm.Field("Gender").Text
				m_oSelectPager.AddSearchAttribute "ProductAgeReferenceId", m_oForm.Field("ProductAgeReferenceId").Text
			Case "Backorder"
				cSubheader = "BackorderItem.gif"
				m_oSelectPager.AddSearchAttribute "AccountContactKey", m_oSecurityState.AccountContactKey
				m_oSelectPager.AddSearchAttribute "CartKey", m_oInstinct.Item("CartState").CartKey
			Case "Favorite"
				cSubheader = "FavoriteProduct.gif"
			Case "Online"
				cSubheader = "OnlineExclusiveProduct.gif"
			Case "New"
				cSubheader = "NewProduct.gif"
				m_oSelectPager.AddSearchAttribute "CustomAge", m_oForm.Field("CustomAge").Value
			Case "HotProduct"
				cSubheader = "HotProduct.gif"
			Case "OnSale"
				cSubheader = "OnSale.gif"
			Case "Bin"
				cSubheader = "BinItems.gif"
			End Select
			Select Case (m_oSelectPager.SortField)
			Case "Name", "Id", "CreateDate"
				m_oSelectPager.SortPrefix = "ProductGroup"
			End Select
			m_oSelectPager.ExecuteContract "ProductGroup", "Hash:View~cXml"
		End Select
		Dim cSelectNavigation: cSelectNavigation = m_oSelectPager.RenderNavigation
		.o_Tr ""
			.o_Td Null, "align=left"
				.o_Table "width=100%"
				.o_Tr ""
					.o_Td Null, "style=color:#666666~font-size: 14px"
					If (m_oForm.Task <> "Search") Then
						'JMW 24Sep2012: Remove Grey filter boxed as per Sales & Marketing
						.x_ m_oInstinct.AxB(m_oInstinct.xAy("""", cProductCategoryName, """"), " and ", m_oInstinct.xAy("""", cProductSubCategory, """")): .x_Br
						'.x_ m_oInstinct.xAy("""", cProductCategoryName, """"): .x_Br
                  .x_ m_oSelectPager.RecordCount & " item" & m_oInstinct.Iif(m_oSelectPager.RecordCount = 1, "", "s") : .x_Br
					ElseIf (m_cSearchText <> "") Then
						.x_ "Search results for """ & m_cSearchText & """": .x_Br
                  .x_ m_oSelectPager.RecordCount & " item" & m_oInstinct.Iif(m_oSelectPager.RecordCount = 1, "", "s") & " found" : .x_Br
					Else
						.x_ "Search results": .x_Br
                  .x_ m_oSelectPager.RecordCount & " item" & m_oInstinct.Iif(m_oSelectPager.RecordCount = 1, "", "s") & " found" : .x_Br
					End If
						.x_ "Page " & m_oSelectPager.PageKey & " of " & m_oInstinct.Iif(m_oSelectPager.IsShowAll = True, 1, m_oSelectPager.PageMax): .x_Br
					.o_Td Null, "valign=bottom;align=right"
					If (bIsLoggedIn = True) Then
						.x_ m_oForm.Render_Button("Image_Button", "/_PROJECT/_Atom/Product/AddToCart_O.gif", "o.Action.value='AddToCart" & m_oinstinct.Iif(m_oForm.Task = "Backorder", "Backorder", "") & "';o.submit();")
					End If
				.x_Table
				.x_Br
				.Table = "TD"
		.o_Tr ""
			.o_Td Null, ""
				'+ engine
				m_oFrame.Render oHtml, "EngineHeadHead"
			If (Right(cSubHeader, 4) = ".gif") Then
				.x_Img m_cRoot & cSubHeader, ""
			Else
				.x_ cSubHeader
			End If
				m_oFrame.Render oHtml, "EngineHeadAction"
				.x_ cSelectNavigation
				m_oFrame.Render oHtml, "EngineHeadFoot"
				'+ select
				m_oFrame.Padding = 3
				m_oFrame.Render oHtml, "SelectHead"
				m_oFrame.Render oHtml, "SelectTR"
					.x_ m_oSelectPager.RenderHeader
		Dim bIsEof : bIsEof = m_oSelectPager.Eof
		If (m_oSelectPager.Eof = True) Then
				m_oFrame.Render oHtml, "SelectFoot"
		.o_Tr ""
			.o_Td Null, "align=center"
				.x_Br
				.x_ "<b>No products are available for your selection.</b>": .x_Br
				.x_ "<b>This may be due to your account settings and configuration.</b>"
			If (m_oSecurityState.AccountContactKey > 0) Then
				.x_Br
			Select Case m_oSecurityState.AccountContactRole
			Case "Executive", "Manager"
				.o_A "/Account/Account.asp" & m_oInstinct.Item("Web").Url("&Method=Account&~="), ""
				.x_ "<b>Click here to update your account information.</b>": .x_A
			Case "User"
				.x_ "<b>Please contact the person responsible for your account.</b>"
			End Select
			End If
		Else
				Do While (m_oSelectPager.Eof = False)
				m_oFrame.Render oHtml, "SelectTR"
			'		If (m_cErrorHash.Exists(m_oInstinct.Lng(m_oSelectPager("Key"), 0)) = True) Then
			'	.o_Tr ""
			'		.o_Td m_oFrame.CurrentSelectTdStyle, "colspan=10;align=center"
			'			.o_ "z-iError", m_cErrorHash(m_oInstinct.Lng(m_oSelectPager("Key"), 0))
			'	.o_Tr ""
			'		End If
					m_oProduct.RenderProductRow oHtml, m_oFrame, m_oForm, m_oSelectPager, cRequestHash, m_oFrame.CurrentSelectTdStyle, cResultType
					m_oSelectPager.MoveNext
				Loop
				m_oFrame.Render oHtml, "SelectFoot"
		End If
				m_oSelectPager.Close
			'	m_oFrame.Render oHtml, "EngineFoot"
		If ((m_oForm.Task = "Backorder") And (bIsEof = False)) Then
		.o_Tr ""
			.o_Td Null, "align=left"
				.o_A "javascript:selectAllCheckboxes(document.Form_" & m_oForm.Name & ");", ""
					.x_ "select all"
				.x_A
				.x_ " | "
				.o_A "javascript:clearAllCheckboxes(document.Form_" & m_oForm.Name & ");", ""
					.x_ "clear all"
				.x_A
		End If
		.o_Tr ""
			.o_Td Null, "align=left"
				.x_Br
				.o_Table "width=100%;class=EngineFoot"
				.o_Tr ""
					.o_Td Null, ""
						.x_ cSelectNavigation
					.o_Td Null, "align=right"
					If (bIsLoggedIn = True) Then
						.x_ m_oForm.Render_Button("Image_Button", "/_PROJECT/_Atom/Product/AddToCart_O.gif", "o.Action.value='AddToCart" & m_oinstinct.Iif(m_oForm.Task = "Backorder", "Backorder", "") & "';o.submit();")
					End If
				.x_Table
				'+ persist
				Dim hRequest: For Each hRequest In cRequestHash
					If (cRequestHash(hRequest) <> "") Then
					If (m_oCartState.IsOverrideBalance = False) Then
					.x_Hidden "Quantity_" & hRequest, cRequestHash(hRequest)
					Else
					.x_Hidden "Balance_" & Mid(hRequest, 2), cRequestHash(hRequest)
					End If
					End If
				Next
				.Table = "TD"
		If ((m_oForm.Task = "Backorder") And (bIsEof = False)) Then
			.x_ m_oForm.Render_Button("Image", "/_Project/_Block/Action/Delete_O.gif", "if (confirm('Are you sure you want to remove these backorder item(s)?') == true) {o.Action.value='DeleteBackorderItems';o.submit();} return false;")
		End If
		m_oInstinct.Kill_Hash cRequestHash
		cRefArray = Null
		x_ oHtml, "Foot"
		End With
	End Sub

	'- Index_Retrieve -'
	Public Sub Index_Retrieve()
		m_oSelectPager.IsPersist = True
		m_oSelectPager.Search "Name"
		If ((m_oForm.Task = "Search") And (m_oSession("Search.Keyword") & "" <> "")) Then
			m_oSelectPager.InitialSort = "~SearchRank"
		ElseIf (m_oForm.Task = "HotProduct") Then
			m_oSelectPager.InitialSort = "HotProductRank"
		ElseIf (m_oForm.Task = "New" Or m_oForm.Task = "OnSale") Then
			m_oSelectPager.InitialSort = "~CreateDate"
			m_oSelectPager.DefaultSort = "~CreateDate"
		Else
			m_oSelectPager.DefaultSort = "Name"
		End If
		If (m_oForm.Task = "Favorite") Then
			m_oSelectPager.PageLength = 500
		Else
			m_oSelectPager.PageLength = 10
		End If
		m_oForm.Field_Add "ProductSubCategoryId", "by subcategory", Null, "Cbo", "30"
			m_oForm.Field("ProductSubCategoryId").Extra = "onchange=var o=document.Form_" & m_oForm.Name & "~o.Action.value='Update'~o.submit()~return(false)~"
		m_oForm.Field_Add "Gender", "by gender", Null, "Cbo", "20"
			m_oForm.Field("Gender").Extra = "onchange=var o=document.Form_" & m_oForm.Name & "~o.Action.value='Update'~o.submit()~return(false)~"
		m_oForm.Field_Add "ProductAgeReferenceId", "by age range", Null, "Cbo", "20"
			m_oForm.Field("ProductAgeReferenceId").Extra = "onchange=var o=document.Form_" & m_oForm.Name & "~o.Action.value='Update'~o.submit()~return(false)~"
		m_oForm.Field_Add "CustomAge", "<div style=""margin-top:15px;font-weight:bold;font-size:13px;"">Show new products from</div> ", Null, "Cbo", "20"
		m_oForm.Field("CustomAge").Extra = "id=CustomAge"
		m_oForm.Field_Retrieve
		If (m_oForm.Task = "Backorder") Then
			Dim oMas200: Set oMas200 = m_oInstinct.LoadLibrary("Atom_Mas200")
			oMas200.UpdateBackorderInventory m_oSecurityState.AccountKey, m_oSecurityState.AccountContactKey
			m_oInstinct.FreeLibrary oMas200
		End If
		If (m_oForm.Action = "Init") Then
			If (Request.QueryString("SubId") & "" <> "") Then
				Dim cSubId : cSubId = m_oInstinct.Str(Request.QueryString("SubId"), "")
				m_oForm.Field("ProductSubCategoryId").Value = m_oInstinct.Iif(cSubId = "All", "", cSubId)
			End If
			If (Request.QueryString("Gender") & "" <> "") Then
				m_oForm.Field("Gender").Value = m_oInstinct.Str(Request.QueryString("Gender"), "")
			End If
			If (Request.QueryString("Age") & "" <> "") Then
				m_oForm.Field("ProductAgeReferenceId").Value = m_oInstinct.Str(Request.QueryString("Age"), "")
			End If
			m_oForm.Field("CustomAge").Value = m_oInstinct.Str(m_oSession("CustomProductAge"), "30")
		End If
	End Sub

	'- Index_Commit -'
	Private Sub Index_Commit()
		m_oForm.Action = "Next"
	End Sub
%>