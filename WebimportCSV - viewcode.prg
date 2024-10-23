PUBLIC odataimport

odataimport=NEWOBJECT("dataimport")
odataimport.Show
RETURN


	**************************************************
*-- Form:         dataimport (f:\sb\bml\webcsv\webimpcsv.scx)
*-- ParentClass:  form
*-- BaseClass:    form
*-- Time Stamp:   10/21/24 04:19:02 PM
*
DEFINE CLASS dataimport AS form


	DataSession = 1
	Height = 719
	Width = 1514
	ShowWindow = 0
	DoCreate = .T.
	ShowTips = .F.
	BufferMode = 0
	AutoCenter = .T.
	Picture = "..\web\graphics\background.jpg"
	Caption = "Data Import"
	WindowType = 0
	WindowState = 2
	LockScreen = .F.
	BackColor = RGB(236,233,216)
	errornum = .F.
	errortxt = .F.
	reject = .F.
	endorttl = .F.
	lfindintpoliid = .F.
	lmultipoliid = .F.
	processed = .F.
	note = ""
	cimpfiledt = .F.
	Name = "DataImport"
	DIMENSION laendorse[1,1]


	ADD OBJECT cmddataassist AS commandbutton WITH ;
		AutoSize = .T., ;
		Top = 2, ;
		Left = 921, ;
		Height = 25, ;
		Width = 72, ;
		FontSize = 8, ;
		Caption = "Data Assist", ;
		Enabled = .F., ;
		TabIndex = 7, ;
		ToolTipText = "Automatic editing and matching before import", ;
		Name = "cmdDataAssist"


	ADD OBJECT cmdgetimportfile AS commandbutton WITH ;
		AutoSize = .T., ;
		Top = 2, ;
		Left = 436, ;
		Height = 25, ;
		Width = 59, ;
		FontSize = 8, ;
		Caption = "Get file...", ;
		Enabled = .F., ;
		TabIndex = 4, ;
		ToolTipText = "Get a CSV file to import", ;
		Name = "cmdGetImportFile"


	ADD OBJECT txtimportfile AS textbox WITH ;
		Value = "", ;
		Height = 23, ;
		Left = 507, ;
		ReadOnly = .T., ;
		TabIndex = 5, ;
		Top = 3, ;
		Width = 270, ;
		Name = "txtImportFile"


	ADD OBJECT edtnote AS editbox WITH ;
		Anchor = 10, ;
		Height = 94, ;
		Left = 5, ;
		ReadOnly = .T., ;
		TabIndex = 14, ;
		Top = 370, ;
		Width = 1500, ;
		ControlSource = "webimp.notes", ;
		Name = "edtNote"


	ADD OBJECT grdpi_ssn AS grid WITH ;
		ColumnCount = 27, ;
		Anchor = 10, ;
		DeleteMark = .F., ;
		Height = 74, ;
		Left = 5, ;
		Panel = 1, ;
		ReadOnly = .T., ;
		RecordSource = "PI_SSN", ;
		TabIndex = 15, ;
		ToolTipText = "Match = SSN4/EIN + First (left 2)+ Last (left 3) + State", ;
		Top = 468, ;
		Width = 1500, ;
		Optimize = .T., ;
		Name = "grdPI_SSN", ;
		Column1.ColumnOrder = 5, ;
		Column1.ControlSource = "PI_SSN.polinumb", ;
		Column1.Width = 105, ;
		Column1.ReadOnly = .T., ;
		Column1.Name = "clmPoliNumb", ;
		Column2.ColumnOrder = 6, ;
		Column2.ControlSource = "PI_SSN.poliid", ;
		Column2.ReadOnly = .T., ;
		Column2.Name = "clmPoliID", ;
		Column3.ColumnOrder = 7, ;
		Column3.ControlSource = "PI_SSN.firstname", ;
		Column3.ReadOnly = .T., ;
		Column3.Name = "clmFirstName", ;
		Column4.ColumnOrder = 9, ;
		Column4.ControlSource = "PI_SSN.company", ;
		Column4.ReadOnly = .T., ;
		Column4.Name = "clmCompany", ;
		Column5.ColumnOrder = 4, ;
		Column5.ControlSource = "PI_SSN.tipe", ;
		Column5.Width = 30, ;
		Column5.ReadOnly = .T., ;
		Column5.Name = "clmTipe", ;
		Column6.ColumnOrder = 3, ;
		Column6.ControlSource = "PI_SSN.license", ;
		Column6.ReadOnly = .T., ;
		Column6.Name = "clmLicense", ;
		Column7.ColumnOrder = 10, ;
		Column7.ControlSource = "PI_SSN.address1", ;
		Column7.ReadOnly = .T., ;
		Column7.Name = "clmAddress1", ;
		Column8.ColumnOrder = 11, ;
		Column8.ControlSource = "PI_SSN.city", ;
		Column8.ReadOnly = .T., ;
		Column8.Name = "clmCity", ;
		Column9.Alignment = 2, ;
		Column9.ColumnOrder = 12, ;
		Column9.ControlSource = "PI_SSN.state", ;
		Column9.Width = 30, ;
		Column9.ReadOnly = .T., ;
		Column9.Name = "clmState", ;
		Column10.ColumnOrder = 13, ;
		Column10.ControlSource = "PI_SSN.zip", ;
		Column10.Width = 40, ;
		Column10.ReadOnly = .T., ;
		Column10.Name = "clmZIP", ;
		Column11.ColumnOrder = 26, ;
		Column11.ControlSource = "PI_SSN.county", ;
		Column11.ReadOnly = .T., ;
		Column11.Name = "clmCounty", ;
		Column12.ColumnOrder = 14, ;
		Column12.ControlSource = "PI_SSN.phone1", ;
		Column12.Width = 90, ;
		Column12.ReadOnly = .T., ;
		Column12.Name = "clmPhone1", ;
		Column13.ColumnOrder = 15, ;
		Column13.ControlSource = "PI_SSN.phone2", ;
		Column13.Width = 90, ;
		Column13.ReadOnly = .T., ;
		Column13.Name = "clmPhone2", ;
		Column14.ColumnOrder = 16, ;
		Column14.ControlSource = "PI_SSN.fax", ;
		Column14.Width = 90, ;
		Column14.ReadOnly = .T., ;
		Column14.Name = "clmFAX", ;
		Column15.ColumnOrder = 17, ;
		Column15.ControlSource = "PI_SSN.origdate", ;
		Column15.Width = 66, ;
		Column15.ReadOnly = .T., ;
		Column15.Name = "clmOrigDate", ;
		Column16.ColumnOrder = 18, ;
		Column16.ControlSource = "PI_SSN.effective", ;
		Column16.Width = 68, ;
		Column16.ReadOnly = .T., ;
		Column16.Name = "clmEffective", ;
		Column17.ColumnOrder = 19, ;
		Column17.ControlSource = "PI_SSN.end", ;
		Column17.Width = 68, ;
		Column17.ReadOnly = .T., ;
		Column17.Name = "clmEnd", ;
		Column18.ColumnOrder = 20, ;
		Column18.ControlSource = "PI_SSN.prem", ;
		Column18.ReadOnly = .T., ;
		Column18.Name = "clmPrem", ;
		Column19.ColumnOrder = 21, ;
		Column19.ControlSource = "PI_SSN.taxes", ;
		Column19.Width = 40, ;
		Column19.ReadOnly = .T., ;
		Column19.Name = "clmTaxes", ;
		Column20.ColumnOrder = 22, ;
		Column20.ControlSource = "PI_SSN.entrydate", ;
		Column20.Width = 68, ;
		Column20.ReadOnly = .T., ;
		Column20.Name = "clmEntryDate", ;
		Column21.ColumnOrder = 23, ;
		Column21.ControlSource = "PI_SSN.entryby", ;
		Column21.ReadOnly = .T., ;
		Column21.Name = "clmEntryBy", ;
		Column22.ColumnOrder = 24, ;
		Column22.ControlSource = "PI_SSN.entrytime", ;
		Column22.Width = 55, ;
		Column22.ReadOnly = .T., ;
		Column22.Name = "clmEntryTime", ;
		Column23.ColumnOrder = 25, ;
		Column23.ControlSource = "PI_SSN.email", ;
		Column23.ReadOnly = .T., ;
		Column23.Name = "clmEmail", ;
		Column24.ColumnOrder = 8, ;
		Column24.ControlSource = "PI_SSN.lastname", ;
		Column24.ReadOnly = .T., ;
		Column24.Name = "clmLastName", ;
		Column25.ColumnOrder = 27, ;
		Column25.ControlSource = "PI_SSN.address2", ;
		Column25.ReadOnly = .T., ;
		Column25.Name = "clmAddress2", ;
		Column26.Alignment = 2, ;
		Column26.ColumnOrder = 1, ;
		Column26.ControlSource = "PI_SSN.cssn4", ;
		Column26.Width = 37, ;
		Column26.ReadOnly = .T., ;
		Column26.Name = "clmCSSN4", ;
		Column27.ColumnOrder = 2, ;
		Column27.ControlSource = "PI_SSN.cein", ;
		Column27.ReadOnly = .T., ;
		Column27.Name = "clmCEIN"


	ADD OBJECT dataimport.grdpi_ssn.clmpolinumb.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Polinumb", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmpolinumb.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.polinumb", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmpoliid.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "PoliID", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmpoliid.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.poliid", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmfirstname.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "FirstName", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmfirstname.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.lastname", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmcompany.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Company", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmcompany.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.company", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmtipe.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Tipe", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmtipe.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.tipe", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmlicense.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "License", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmlicense.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.license", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmaddress1.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Address1", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmaddress1.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.address1", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmcity.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "City", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmcity.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.city", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmstate.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "State", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmstate.text1 AS textbox WITH ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.state", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmzip.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "ZIP", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmzip.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.zip", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmcounty.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "County", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmcounty.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.county", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmphone1.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Phone1", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmphone1.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.phone1", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmphone2.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Phone2", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmphone2.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.phone2", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmfax.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "FAX", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmfax.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.fax", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmorigdate.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "OrigDate", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmorigdate.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.origdate", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmeffective.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Effective", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmeffective.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.effective", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmend.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "End", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmend.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.end", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmprem.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "PremAmt", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmprem.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.prem", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmtaxes.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Taxes", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmtaxes.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.taxes", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmentrydate.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "EntryDate", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmentrydate.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.entrydate", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmentryby.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "EntryBy", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmentryby.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.entryby", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmentrytime.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "EntryTime", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmentrytime.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.entrytime", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmemail.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "eMail", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmemail.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.email", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmlastname.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "LastName", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmlastname.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.lastname", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmaddress2.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Address2", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmaddress2.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.address2", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmcssn4.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 9, ;
		FontUnderline = .T., ;
		Alignment = 2, ;
		Caption = "SSN4", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmcssn4.text1 AS textbox WITH ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.cssn4", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_ssn.clmcein.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 9, ;
		FontUnderline = .T., ;
		Alignment = 2, ;
		Caption = "EIN", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_ssn.clmcein.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_SSN.cein", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT grdpi_lic AS grid WITH ;
		ColumnCount = 27, ;
		Anchor = 10, ;
		DeleteMark = .F., ;
		Height = 74, ;
		Left = 5, ;
		Panel = 1, ;
		ReadOnly = .T., ;
		RecordSource = "PI_Lic", ;
		TabIndex = 16, ;
		ToolTipText = "Match = State + Policy LIcense#", ;
		Top = 545, ;
		Width = 1500, ;
		Optimize = .T., ;
		Name = "grdPI_Lic", ;
		Column1.ColumnOrder = 5, ;
		Column1.ControlSource = "PI_Lic.polinumb", ;
		Column1.Width = 105, ;
		Column1.ReadOnly = .T., ;
		Column1.Name = "clmPoliNumb", ;
		Column2.ColumnOrder = 6, ;
		Column2.ControlSource = "PI_Lic.poliid", ;
		Column2.ReadOnly = .T., ;
		Column2.Name = "clmPoliID", ;
		Column3.ColumnOrder = 7, ;
		Column3.ControlSource = "PI_Lic.firstname", ;
		Column3.ReadOnly = .T., ;
		Column3.Name = "clmFirstName", ;
		Column4.ColumnOrder = 9, ;
		Column4.ControlSource = "PI_Lic.company", ;
		Column4.ReadOnly = .T., ;
		Column4.Name = "clmCompany", ;
		Column5.ColumnOrder = 2, ;
		Column5.ControlSource = "PI_Lic.tipe", ;
		Column5.Width = 30, ;
		Column5.ReadOnly = .T., ;
		Column5.Name = "clmTipe", ;
		Column6.ColumnOrder = 1, ;
		Column6.ControlSource = "PI_Lic.license", ;
		Column6.ReadOnly = .T., ;
		Column6.Name = "clmLicense", ;
		Column7.ColumnOrder = 10, ;
		Column7.ControlSource = "PI_Lic.address1", ;
		Column7.ReadOnly = .T., ;
		Column7.Name = "clmAddress1", ;
		Column8.ColumnOrder = 11, ;
		Column8.ControlSource = "PI_Lic.city", ;
		Column8.ReadOnly = .T., ;
		Column8.Name = "clmCity", ;
		Column9.Alignment = 2, ;
		Column9.ColumnOrder = 12, ;
		Column9.ControlSource = "PI_Lic.state", ;
		Column9.Width = 30, ;
		Column9.ReadOnly = .T., ;
		Column9.Name = "clmState", ;
		Column10.ColumnOrder = 13, ;
		Column10.ControlSource = "PI_Lic.zip", ;
		Column10.Width = 40, ;
		Column10.ReadOnly = .T., ;
		Column10.Name = "clmZIP", ;
		Column11.ColumnOrder = 27, ;
		Column11.ControlSource = "PI_Lic.county", ;
		Column11.ReadOnly = .T., ;
		Column11.Name = "clmCounty", ;
		Column12.ColumnOrder = 14, ;
		Column12.ControlSource = "PI_Lic.phone1", ;
		Column12.Width = 90, ;
		Column12.ReadOnly = .T., ;
		Column12.Name = "clmPhone1", ;
		Column13.ColumnOrder = 15, ;
		Column13.ControlSource = "PI_Lic.phone2", ;
		Column13.Width = 90, ;
		Column13.ReadOnly = .T., ;
		Column13.Name = "clmPhone2", ;
		Column14.ColumnOrder = 16, ;
		Column14.ControlSource = "PI_Lic.fax", ;
		Column14.Width = 90, ;
		Column14.ReadOnly = .T., ;
		Column14.Name = "clmFAX", ;
		Column15.ColumnOrder = 17, ;
		Column15.ControlSource = "PI_Lic.origdate", ;
		Column15.Width = 68, ;
		Column15.ReadOnly = .T., ;
		Column15.Name = "clmOrigDate", ;
		Column16.ColumnOrder = 18, ;
		Column16.ControlSource = "PI_Lic.effective", ;
		Column16.Width = 68, ;
		Column16.ReadOnly = .T., ;
		Column16.Name = "clmEffective", ;
		Column17.ColumnOrder = 19, ;
		Column17.ControlSource = "PI_Lic.end", ;
		Column17.Width = 68, ;
		Column17.ReadOnly = .T., ;
		Column17.Name = "clmEnd", ;
		Column18.ColumnOrder = 20, ;
		Column18.ControlSource = "PI_Lic.end", ;
		Column18.ReadOnly = .T., ;
		Column18.Name = "clmPrem", ;
		Column19.ColumnOrder = 21, ;
		Column19.ControlSource = "PI_Lic.taxes", ;
		Column19.Width = 40, ;
		Column19.ReadOnly = .T., ;
		Column19.Name = "clmTaxes", ;
		Column20.ColumnOrder = 22, ;
		Column20.ControlSource = "PI_Lic.entrydate", ;
		Column20.Width = 68, ;
		Column20.ReadOnly = .T., ;
		Column20.Name = "clmEntryDate", ;
		Column21.ColumnOrder = 23, ;
		Column21.ControlSource = "PI_Lic.entryby", ;
		Column21.ReadOnly = .T., ;
		Column21.Name = "clmEntryBy", ;
		Column22.ColumnOrder = 24, ;
		Column22.ControlSource = "PI_Lic.entrytime", ;
		Column22.Width = 55, ;
		Column22.ReadOnly = .T., ;
		Column22.Name = "clmEntryTime", ;
		Column23.ColumnOrder = 25, ;
		Column23.ControlSource = "PI_Lic.email", ;
		Column23.ReadOnly = .T., ;
		Column23.Name = "clmEmail", ;
		Column24.ColumnOrder = 8, ;
		Column24.ControlSource = "PI_Lic.lastname", ;
		Column24.ReadOnly = .T., ;
		Column24.Name = "clmLastName", ;
		Column25.ColumnOrder = 26, ;
		Column25.ControlSource = "PI_Lic.address2", ;
		Column25.ReadOnly = .T., ;
		Column25.Name = "clmAddress2", ;
		Column26.Alignment = 2, ;
		Column26.ColumnOrder = 3, ;
		Column26.ControlSource = "PI_Lic.cssn4", ;
		Column26.Width = 38, ;
		Column26.ReadOnly = .T., ;
		Column26.Name = "clmCSSN4", ;
		Column27.Alignment = 2, ;
		Column27.ColumnOrder = 4, ;
		Column27.ControlSource = "PI_Lic.cein", ;
		Column27.ReadOnly = .T., ;
		Column27.Name = "clmEIN"


	ADD OBJECT dataimport.grdpi_lic.clmpolinumb.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Polinumb", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmpolinumb.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.polinumb", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmpoliid.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "PoliID", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmpoliid.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.poliid", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmfirstname.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "FirstName", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmfirstname.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.lastname", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmcompany.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Company", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmcompany.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.company", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmtipe.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 9, ;
		FontUnderline = .T., ;
		Caption = "Tipe", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmtipe.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.tipe", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmlicense.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 9, ;
		FontUnderline = .T., ;
		Alignment = 2, ;
		Caption = "License", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmlicense.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.license", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmaddress1.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Address1", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmaddress1.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.address1", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmcity.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "City", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmcity.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.city", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmstate.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "State", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmstate.text1 AS textbox WITH ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.state", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmzip.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "ZIP", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmzip.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.zip", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmcounty.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "County", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmcounty.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.county", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmphone1.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Phone1", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmphone1.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.phone1", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmphone2.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Phone2", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmphone2.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.phone2", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmfax.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "FAX", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmfax.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.fax", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmorigdate.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "OrigDate", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmorigdate.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.origdate", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmeffective.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Effective", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmeffective.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.effective", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmend.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "End", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmend.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.end", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmprem.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "PremAmt", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmprem.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.prem", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmtaxes.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Taxes", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmtaxes.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.taxes", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmentrydate.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "EntryDate", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmentrydate.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.entrydate", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmentryby.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "EntryBy", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmentryby.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.entryby", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmentrytime.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "EntryTime", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmentrytime.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.entrytime", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmemail.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "eMail", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmemail.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.email", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmlastname.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "LastName", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmlastname.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.lastname", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmaddress2.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Address2", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmaddress2.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.address2", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmcssn4.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "SSN4", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmcssn4.text1 AS textbox WITH ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.cssn4", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdpi_lic.clmein.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "EIN", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdpi_lic.clmein.text1 AS textbox WITH ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.cein", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT cmdimport AS commandbutton WITH ;
		AutoSize = .T., ;
		Top = 2, ;
		Left = 1029, ;
		Height = 25, ;
		Width = 45, ;
		FontBold = .F., ;
		FontSize = 8, ;
		Caption = "Import", ;
		Enabled = .F., ;
		TabIndex = 8, ;
		ToolTipText = "Import records into BML data.", ;
		Name = "cmdImport"


	ADD OBJECT txtuser AS textbox WITH ;
		Height = 23, ;
		Left = 154, ;
		TabIndex = 2, ;
		Top = 3, ;
		Width = 100, ;
		Name = "txtUser"


	ADD OBJECT txtdate AS textbox WITH ;
		Enabled = .F., ;
		Height = 23, ;
		Left = 296, ;
		TabIndex = 3, ;
		Top = 3, ;
		Width = 100, ;
		Name = "txtDate"


	ADD OBJECT lbluser AS label WITH ;
		AutoSize = .T., ;
		BackStyle = 1, ;
		Caption = "User:", ;
		Height = 17, ;
		Left = 121, ;
		Top = 6, ;
		Width = 32, ;
		TabIndex = 18, ;
		ToolTipText = "User entering this data", ;
		Name = "lblUser"


	ADD OBJECT lbldate AS label WITH ;
		AutoSize = .T., ;
		Caption = "Date:", ;
		Height = 17, ;
		Left = 262, ;
		Top = 6, ;
		Width = 31, ;
		TabIndex = 19, ;
		ToolTipText = "Date for this import", ;
		Name = "lblDate"


	ADD OBJECT cmdexit AS commandbutton WITH ;
		AutoSize = .T., ;
		Top = 2, ;
		Left = 1308, ;
		Height = 25, ;
		Width = 33, ;
		FontBold = .F., ;
		FontSize = 8, ;
		Caption = "Exit", ;
		TabIndex = 12, ;
		ToolTipText = "Exit the program.", ;
		Name = "cmdExit"


	ADD OBJECT grdwebimp AS grid WITH ;
		ColumnCount = 88, ;
		Anchor = 11, ;
		DeleteMark = .T., ;
		GridLines = 2, ;
		GridLineWidth = 1, ;
		HeaderHeight = 20, ;
		Height = 270, ;
		Left = 5, ;
		Panel = 1, ;
		RecordMark = .F., ;
		RecordSource = "webimp", ;
		RecordSourceType = 1, ;
		ScrollBars = 3, ;
		TabIndex = 13, ;
		Top = 36, ;
		Width = 1500, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		GridLineColor = RGB(0,128,128), ;
		HighlightBackColor = RGB(255,255,128), ;
		HighlightForeColor = RGB(0,0,0), ;
		HighlightStyle = 2, ;
		Optimize = .T., ;
		Name = "grdWebImp", ;
		Column1.FontBold = .T., ;
		Column1.FontName = "MS Sans Serif", ;
		Column1.FontSize = 8, ;
		Column1.Alignment = 0, ;
		Column1.ColumnOrder = 2, ;
		Column1.ControlSource = "webimp.polinumb", ;
		Column1.Width = 110, ;
		Column1.Visible = .T., ;
		Column1.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column1.ForeColor = RGB(0,0,0), ;
		Column1.BackColor = RGB(236,233,216), ;
		Column1.Name = "clmPoliNumb", ;
		Column2.FontBold = .T., ;
		Column2.FontName = "MS Sans Serif", ;
		Column2.FontSize = 8, ;
		Column2.Alignment = 0, ;
		Column2.ColumnOrder = 3, ;
		Column2.ControlSource = "webimp.pdanum", ;
		Column2.Width = 90, ;
		Column2.Visible = .T., ;
		Column2.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column2.ForeColor = RGB(0,0,0), ;
		Column2.BackColor = RGB(236,233,216), ;
		Column2.Name = "clmPDANum", ;
		Column3.FontBold = .T., ;
		Column3.FontName = "MS Sans Serif", ;
		Column3.FontSize = 8, ;
		Column3.Alignment = 0, ;
		Column3.ColumnOrder = 4, ;
		Column3.ControlSource = "webimp.poliid", ;
		Column3.Width = 65, ;
		Column3.Visible = .T., ;
		Column3.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column3.ForeColor = RGB(0,0,0), ;
		Column3.BackColor = RGB(236,233,216), ;
		Column3.Name = "clmPoliID", ;
		Column4.FontBold = .T., ;
		Column4.FontName = "MS Sans Serif", ;
		Column4.FontSize = 8, ;
		Column4.Alignment = 0, ;
		Column4.ColumnOrder = 5, ;
		Column4.ControlSource = "webimp.firstname", ;
		Column4.Width = 90, ;
		Column4.Visible = .T., ;
		Column4.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column4.ForeColor = RGB(0,0,0), ;
		Column4.BackColor = RGB(236,233,216), ;
		Column4.Name = "clmFirstname", ;
		Column5.FontBold = .T., ;
		Column5.FontName = "MS Sans Serif", ;
		Column5.FontSize = 8, ;
		Column5.Alignment = 0, ;
		Column5.ColumnOrder = 7, ;
		Column5.ControlSource = "webimp.lastname", ;
		Column5.Width = 90, ;
		Column5.Visible = .T., ;
		Column5.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column5.ForeColor = RGB(0,0,0), ;
		Column5.BackColor = RGB(236,233,216), ;
		Column5.Name = "clmLastname", ;
		Column6.FontBold = .T., ;
		Column6.FontName = "MS Sans Serif", ;
		Column6.FontSize = 8, ;
		Column6.Alignment = 0, ;
		Column6.ColumnOrder = 8, ;
		Column6.ControlSource = "webimp.company", ;
		Column6.Width = 90, ;
		Column6.Visible = .T., ;
		Column6.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column6.ForeColor = RGB(0,0,0), ;
		Column6.BackColor = RGB(236,233,216), ;
		Column6.Name = "clmCompany", ;
		Column7.FontBold = .T., ;
		Column7.FontName = "MS Sans Serif", ;
		Column7.FontSize = 8, ;
		Column7.Alignment = 0, ;
		Column7.ColumnOrder = 12, ;
		Column7.ControlSource = "webimp.address1", ;
		Column7.Width = 90, ;
		Column7.Visible = .T., ;
		Column7.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column7.ForeColor = RGB(0,0,0), ;
		Column7.BackColor = RGB(236,233,216), ;
		Column7.Name = "clmAddress1", ;
		Column8.FontBold = .T., ;
		Column8.FontName = "MS Sans Serif", ;
		Column8.FontSize = 8, ;
		Column8.Alignment = 0, ;
		Column8.ColumnOrder = 14, ;
		Column8.ControlSource = "webimp.city", ;
		Column8.Width = 90, ;
		Column8.Visible = .T., ;
		Column8.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column8.ForeColor = RGB(0,0,0), ;
		Column8.BackColor = RGB(236,233,216), ;
		Column8.Name = "clmCity", ;
		Column9.FontBold = .T., ;
		Column9.FontName = "MS Sans Serif", ;
		Column9.FontSize = 8, ;
		Column9.Alignment = 2, ;
		Column9.ColumnOrder = 15, ;
		Column9.ControlSource = "webimp.state", ;
		Column9.Width = 26, ;
		Column9.Visible = .T., ;
		Column9.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column9.ForeColor = RGB(0,0,0), ;
		Column9.BackColor = RGB(236,233,216), ;
		Column9.Name = "clmState", ;
		Column10.FontBold = .T., ;
		Column10.FontName = "MS Sans Serif", ;
		Column10.FontSize = 8, ;
		Column10.Alignment = 0, ;
		Column10.ColumnOrder = 16, ;
		Column10.ControlSource = "webimp.zip", ;
		Column10.Width = 40, ;
		Column10.Visible = .T., ;
		Column10.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column10.ForeColor = RGB(0,0,0), ;
		Column10.BackColor = RGB(236,233,216), ;
		Column10.Name = "clmZIP", ;
		Column11.FontBold = .T., ;
		Column11.FontName = "MS Sans Serif", ;
		Column11.FontSize = 8, ;
		Column11.Alignment = 0, ;
		Column11.ColumnOrder = 18, ;
		Column11.ControlSource = "webimp.county", ;
		Column11.Width = 90, ;
		Column11.Visible = .T., ;
		Column11.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column11.ForeColor = RGB(0,0,0), ;
		Column11.BackColor = RGB(236,233,216), ;
		Column11.Name = "clmCounty", ;
		Column12.FontBold = .T., ;
		Column12.FontName = "MS Sans Serif", ;
		Column12.FontSize = 8, ;
		Column12.Alignment = 0, ;
		Column12.ColumnOrder = 19, ;
		Column12.ControlSource = "webimp.phone1", ;
		Column12.Width = 90, ;
		Column12.Visible = .T., ;
		Column12.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column12.ForeColor = RGB(0,0,0), ;
		Column12.BackColor = RGB(236,233,216), ;
		Column12.Name = "clmPhone1", ;
		Column13.FontBold = .T., ;
		Column13.FontName = "MS Sans Serif", ;
		Column13.FontSize = 8, ;
		Column13.Alignment = 0, ;
		Column13.ColumnOrder = 20, ;
		Column13.ControlSource = "webimp.phone1ext", ;
		Column13.Width = 45, ;
		Column13.Visible = .T., ;
		Column13.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column13.ForeColor = RGB(0,0,0), ;
		Column13.BackColor = RGB(236,233,216), ;
		Column13.Name = "clmPhone1Ext", ;
		Column14.FontBold = .T., ;
		Column14.FontName = "MS Sans Serif", ;
		Column14.FontSize = 8, ;
		Column14.Alignment = 0, ;
		Column14.ColumnOrder = 23, ;
		Column14.ControlSource = "webimp.fax", ;
		Column14.Width = 90, ;
		Column14.Visible = .T., ;
		Column14.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column14.ForeColor = RGB(0,0,0), ;
		Column14.BackColor = RGB(236,233,216), ;
		Column14.Name = "clmFAX", ;
		Column15.FontBold = .T., ;
		Column15.FontName = "MS Sans Serif", ;
		Column15.FontSize = 8, ;
		Column15.Alignment = 0, ;
		Column15.ColumnOrder = 26, ;
		Column15.ControlSource = "webimp.origdate", ;
		Column15.Width = 75, ;
		Column15.Visible = .T., ;
		Column15.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column15.ForeColor = RGB(0,0,0), ;
		Column15.BackColor = RGB(236,233,216), ;
		Column15.Name = "clmOrigdate", ;
		Column16.FontBold = .T., ;
		Column16.FontName = "MS Sans Serif", ;
		Column16.FontSize = 8, ;
		Column16.Alignment = 0, ;
		Column16.ColumnOrder = 24, ;
		Column16.ControlSource = "webimp.effective", ;
		Column16.Width = 75, ;
		Column16.Visible = .T., ;
		Column16.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column16.ForeColor = RGB(0,0,0), ;
		Column16.BackColor = RGB(236,233,216), ;
		Column16.Name = "clmEffective", ;
		Column17.FontBold = .T., ;
		Column17.FontName = "MS Sans Serif", ;
		Column17.FontSize = 8, ;
		Column17.Alignment = 0, ;
		Column17.ColumnOrder = 25, ;
		Column17.ControlSource = "webimp.end", ;
		Column17.Width = 75, ;
		Column17.Visible = .T., ;
		Column17.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column17.ForeColor = RGB(0,0,0), ;
		Column17.BackColor = RGB(236,233,216), ;
		Column17.Name = "clmEnd", ;
		Column18.FontBold = .T., ;
		Column18.FontName = "MS Sans Serif", ;
		Column18.FontSize = 8, ;
		Column18.Alignment = 1, ;
		Column18.ColumnOrder = 27, ;
		Column18.ControlSource = "webimp.premimport", ;
		Column18.Width = 72, ;
		Column18.Visible = .T., ;
		Column18.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column18.ForeColor = RGB(0,0,0), ;
		Column18.BackColor = RGB(236,233,216), ;
		Column18.Name = "clmPremimport", ;
		Column19.FontBold = .T., ;
		Column19.FontName = "MS Sans Serif", ;
		Column19.FontSize = 8, ;
		Column19.Alignment = 1, ;
		Column19.ColumnOrder = 28, ;
		Column19.ControlSource = "webimp.prem", ;
		Column19.Width = 75, ;
		Column19.Visible = .T., ;
		Column19.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column19.ForeColor = RGB(0,0,0), ;
		Column19.BackColor = RGB(236,233,216), ;
		Column19.Name = "clmPrem", ;
		Column20.FontBold = .T., ;
		Column20.FontName = "MS Sans Serif", ;
		Column20.FontSize = 8, ;
		Column20.Alignment = 1, ;
		Column20.ColumnOrder = 29, ;
		Column20.ControlSource = "webimp.taxes", ;
		Column20.Width = 40, ;
		Column20.Visible = .T., ;
		Column20.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column20.ForeColor = RGB(0,0,0), ;
		Column20.BackColor = RGB(236,233,216), ;
		Column20.Name = "clmTaxes", ;
		Column21.FontBold = .T., ;
		Column21.FontName = "MS Sans Serif", ;
		Column21.FontSize = 8, ;
		Column21.Alignment = 0, ;
		Column21.ColumnOrder = 30, ;
		Column21.ControlSource = "webimp.entrydate", ;
		Column21.Width = 75, ;
		Column21.Visible = .T., ;
		Column21.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column21.ForeColor = RGB(0,0,0), ;
		Column21.BackColor = RGB(236,233,216), ;
		Column21.Name = "clmEntrydate", ;
		Column22.FontBold = .T., ;
		Column22.FontName = "MS Sans Serif", ;
		Column22.FontSize = 8, ;
		Column22.Alignment = 0, ;
		Column22.ColumnOrder = 31, ;
		Column22.ControlSource = "webimp.entryby", ;
		Column22.Width = 77, ;
		Column22.Visible = .T., ;
		Column22.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column22.ForeColor = RGB(0,0,0), ;
		Column22.BackColor = RGB(236,233,216), ;
		Column22.Name = "clmEntryby", ;
		Column23.FontBold = .T., ;
		Column23.FontName = "MS Sans Serif", ;
		Column23.FontSize = 8, ;
		Column23.Alignment = 0, ;
		Column23.ColumnOrder = 32, ;
		Column23.ControlSource = "webimp.entrytime", ;
		Column23.Width = 61, ;
		Column23.Visible = .T., ;
		Column23.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column23.ForeColor = RGB(0,0,0), ;
		Column23.BackColor = RGB(236,233,216), ;
		Column23.Name = "clmEntrytime", ;
		Column24.FontBold = .T., ;
		Column24.FontName = "MS Sans Serif", ;
		Column24.FontSize = 8, ;
		Column24.Alignment = 0, ;
		Column24.ColumnOrder = 33, ;
		Column24.ControlSource = "webimp.email", ;
		Column24.Width = 90, ;
		Column24.Visible = .T., ;
		Column24.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column24.ForeColor = RGB(0,0,0), ;
		Column24.BackColor = RGB(236,233,216), ;
		Column24.Name = "clmEmail", ;
		Column25.FontBold = .F., ;
		Column25.FontName = "MS Sans Serif", ;
		Column25.FontSize = 8, ;
		Column25.Alignment = 2, ;
		Column25.ColumnOrder = 34, ;
		Column25.ControlSource = "webimp.claimalert", ;
		Column25.Width = 20, ;
		Column25.Sparse = .F., ;
		Column25.Visible = .T., ;
		Column25.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column25.ForeColor = RGB(0,0,0), ;
		Column25.BackColor = RGB(236,233,216), ;
		Column25.ToolTipText = "Had previous claim?", ;
		Column25.Name = "clmClaimalert", ;
		Column26.FontBold = .F., ;
		Column26.FontName = "MS Sans Serif", ;
		Column26.FontSize = 8, ;
		Column26.Alignment = 2, ;
		Column26.ColumnOrder = 47, ;
		Column26.ControlSource = "webimp.higherlim1", ;
		Column26.Width = 20, ;
		Column26.Sparse = .F., ;
		Column26.Visible = .T., ;
		Column26.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column26.ForeColor = RGB(0,0,0), ;
		Column26.BackColor = RGB(236,233,216), ;
		Column26.Name = "clmHigherlim1", ;
		Column27.FontBold = .F., ;
		Column27.FontName = "MS Sans Serif", ;
		Column27.FontSize = 8, ;
		Column27.Alignment = 2, ;
		Column27.ColumnOrder = 48, ;
		Column27.ControlSource = "webimp.higherlim2", ;
		Column27.Width = 20, ;
		Column27.Sparse = .F., ;
		Column27.Visible = .T., ;
		Column27.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column27.ForeColor = RGB(0,0,0), ;
		Column27.BackColor = RGB(236,233,216), ;
		Column27.Name = "clmHigherlim2", ;
		Column28.FontBold = .F., ;
		Column28.FontName = "MS Sans Serif", ;
		Column28.FontSize = 8, ;
		Column28.Alignment = 2, ;
		Column28.ColumnOrder = 37, ;
		Column28.ControlSource = "webimp.appraisal", ;
		Column28.Width = 25, ;
		Column28.Sparse = .F., ;
		Column28.Visible = .T., ;
		Column28.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column28.ForeColor = RGB(0,0,0), ;
		Column28.BackColor = RGB(236,233,216), ;
		Column28.Name = "clmAppraisal", ;
		Column29.FontBold = .F., ;
		Column29.FontName = "MS Sans Serif", ;
		Column29.FontSize = 8, ;
		Column29.Alignment = 2, ;
		Column29.ColumnOrder = 54, ;
		Column29.ControlSource = "webimp.propmgr", ;
		Column29.Width = 25, ;
		Column29.Sparse = .F., ;
		Column29.Visible = .T., ;
		Column29.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column29.ForeColor = RGB(0,0,0), ;
		Column29.BackColor = RGB(236,233,216), ;
		Column29.Name = "clmPropmgr", ;
		Column30.FontBold = .F., ;
		Column30.FontName = "MS Sans Serif", ;
		Column30.FontSize = 8, ;
		Column30.Alignment = 2, ;
		Column30.ColumnOrder = 42, ;
		Column30.ControlSource = "webimp.environx", ;
		Column30.Width = 25, ;
		Column30.Sparse = .F., ;
		Column30.Visible = .T., ;
		Column30.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column30.ForeColor = RGB(0,0,0), ;
		Column30.BackColor = RGB(236,233,216), ;
		Column30.Name = "clmEnvironx", ;
		Column31.FontBold = .F., ;
		Column31.FontName = "MS Sans Serif", ;
		Column31.FontSize = 8, ;
		Column31.Alignment = 2, ;
		Column31.ColumnOrder = 44, ;
		Column31.ControlSource = "webimp.fairhouse", ;
		Column31.Width = 25, ;
		Column31.Sparse = .F., ;
		Column31.Visible = .T., ;
		Column31.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column31.ForeColor = RGB(0,0,0), ;
		Column31.BackColor = RGB(236,233,216), ;
		Column31.Name = "clmFairhouse", ;
		Column32.FontBold = .F., ;
		Column32.FontName = "MS Sans Serif", ;
		Column32.FontSize = 8, ;
		Column32.Alignment = 2, ;
		Column32.ColumnOrder = 55, ;
		Column32.ControlSource = "webimp.complaints", ;
		Column32.Width = 25, ;
		Column32.Sparse = .F., ;
		Column32.Visible = .T., ;
		Column32.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column32.ForeColor = RGB(0,0,0), ;
		Column32.BackColor = RGB(236,233,216), ;
		Column32.Name = "clmComplaints", ;
		Column33.FontBold = .F., ;
		Column33.FontName = "MS Sans Serif", ;
		Column33.FontSize = 8, ;
		Column33.Alignment = 2, ;
		Column33.ColumnOrder = 62, ;
		Column33.ControlSource = "webimp.onlinechg", ;
		Column33.Width = 20, ;
		Column33.Sparse = .F., ;
		Column33.Visible = .T., ;
		Column33.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column33.ForeColor = RGB(0,0,0), ;
		Column33.BackColor = RGB(236,233,216), ;
		Column33.Name = "clmOnlinechg", ;
		Column34.FontBold = .F., ;
		Column34.FontName = "MS Sans Serif", ;
		Column34.FontSize = 8, ;
		Column34.Alignment = 2, ;
		Column34.ColumnOrder = 40, ;
		Column34.ControlSource = "webimp.earnmo", ;
		Column34.Width = 25, ;
		Column34.Sparse = .F., ;
		Column34.Visible = .T., ;
		Column34.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column34.ForeColor = RGB(0,0,0), ;
		Column34.BackColor = RGB(236,233,216), ;
		Column34.Name = "clmEarnmo", ;
		Column35.FontBold = .F., ;
		Column35.FontName = "MS Sans Serif", ;
		Column35.FontSize = 8, ;
		Column35.Alignment = 2, ;
		Column35.ColumnOrder = 43, ;
		Column35.ControlSource = "webimp.lcenviro", ;
		Column35.Width = 25, ;
		Column35.Sparse = .F., ;
		Column35.Visible = .T., ;
		Column35.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column35.ForeColor = RGB(0,0,0), ;
		Column35.BackColor = RGB(236,233,216), ;
		Column35.Name = "clmLcenviro", ;
		Column36.FontBold = .F., ;
		Column36.FontName = "MS Sans Serif", ;
		Column36.FontSize = 8, ;
		Column36.Alignment = 2, ;
		Column36.ColumnOrder = 45, ;
		Column36.ControlSource = "webimp.fairhausa", ;
		Column36.Width = 25, ;
		Column36.Sparse = .F., ;
		Column36.Visible = .T., ;
		Column36.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column36.ForeColor = RGB(0,0,0), ;
		Column36.BackColor = RGB(236,233,216), ;
		Column36.Name = "clmFairhausa", ;
		Column37.FontBold = .F., ;
		Column37.FontName = "MS Sans Serif", ;
		Column37.FontSize = 8, ;
		Column37.Alignment = 2, ;
		Column37.ColumnOrder = 46, ;
		Column37.ControlSource = "webimp.fairhausb", ;
		Column37.Width = 25, ;
		Column37.Sparse = .F., ;
		Column37.Visible = .T., ;
		Column37.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column37.ForeColor = RGB(0,0,0), ;
		Column37.BackColor = RGB(236,233,216), ;
		Column37.Name = "clmFairhausb", ;
		Column38.FontBold = .F., ;
		Column38.FontName = "MS Sans Serif", ;
		Column38.FontSize = 8, ;
		Column38.Alignment = 2, ;
		Column38.ColumnOrder = 53, ;
		Column38.ControlSource = "webimp.primeres", ;
		Column38.Width = 25, ;
		Column38.Sparse = .F., ;
		Column38.Visible = .T., ;
		Column38.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column38.ForeColor = RGB(0,0,0), ;
		Column38.BackColor = RGB(236,233,216), ;
		Column38.Name = "clmPrimeres", ;
		Column39.FontBold = .F., ;
		Column39.FontName = "MS Sans Serif", ;
		Column39.FontSize = 8, ;
		Column39.Alignment = 2, ;
		Column39.ColumnOrder = 49, ;
		Column39.ControlSource = "webimp.hilimmil1", ;
		Column39.Width = 28, ;
		Column39.Sparse = .F., ;
		Column39.Visible = .T., ;
		Column39.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column39.ForeColor = RGB(0,0,0), ;
		Column39.BackColor = RGB(236,233,216), ;
		Column39.Name = "clmHilimmil1", ;
		Column40.FontBold = .F., ;
		Column40.FontName = "MS Sans Serif", ;
		Column40.FontSize = 8, ;
		Column40.Alignment = 2, ;
		Column40.ColumnOrder = 50, ;
		Column40.ControlSource = "webimp.hilimmil2", ;
		Column40.Width = 28, ;
		Column40.Sparse = .F., ;
		Column40.Visible = .T., ;
		Column40.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column40.ForeColor = RGB(0,0,0), ;
		Column40.BackColor = RGB(236,233,216), ;
		Column40.Name = "clmHilimmil2", ;
		Column41.FontBold = .T., ;
		Column41.FontName = "MS Sans Serif", ;
		Column41.FontSize = 8, ;
		Column41.Alignment = 2, ;
		Column41.ColumnOrder = 63, ;
		Column41.ControlSource = "webimp.govstate", ;
		Column41.Width = 25, ;
		Column41.Sparse = .F., ;
		Column41.Visible = .T., ;
		Column41.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column41.ForeColor = RGB(0,0,0), ;
		Column41.BackColor = RGB(236,233,216), ;
		Column41.Name = "clmGovstate", ;
		Column42.FontBold = .T., ;
		Column42.FontName = "MS Sans Serif", ;
		Column42.FontSize = 8, ;
		Column42.Alignment = 0, ;
		Column42.ColumnOrder = 64, ;
		Column42.ControlSource = "webimp.batchnum", ;
		Column42.Width = 40, ;
		Column42.Visible = .T., ;
		Column42.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column42.ForeColor = RGB(0,0,0), ;
		Column42.BackColor = RGB(236,233,216), ;
		Column42.Name = "clmBatchnum", ;
		Column43.FontBold = .T., ;
		Column43.FontName = "MS Sans Serif", ;
		Column43.FontSize = 8, ;
		Column43.Alignment = 0, ;
		Column43.ColumnOrder = 13, ;
		Column43.ControlSource = "webimp.address2", ;
		Column43.Width = 56, ;
		Column43.Visible = .T., ;
		Column43.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column43.ForeColor = RGB(0,0,0), ;
		Column43.BackColor = RGB(236,233,216), ;
		Column43.Name = "clmAddress2", ;
		Column44.FontBold = .F., ;
		Column44.Alignment = 2, ;
		Column44.ColumnOrder = 51, ;
		Column44.ControlSource = "webimp.premiuma", ;
		Column44.Width = 25, ;
		Column44.Sparse = .F., ;
		Column44.Visible = .T., ;
		Column44.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column44.ForeColor = RGB(0,0,0), ;
		Column44.BackColor = RGB(236,233,216), ;
		Column44.Name = "clmPremiumA", ;
		Column45.FontBold = .F., ;
		Column45.Alignment = 2, ;
		Column45.ColumnOrder = 52, ;
		Column45.ControlSource = "webimp.premiumb", ;
		Column45.Width = 25, ;
		Column45.Sparse = .F., ;
		Column45.Visible = .T., ;
		Column45.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column45.ForeColor = RGB(0,0,0), ;
		Column45.BackColor = RGB(236,233,216), ;
		Column45.Name = "clmPremiumB", ;
		Column46.Alignment = 2, ;
		Column46.ColumnOrder = 1, ;
		Column46.ControlSource = "webimp.lhold", ;
		Column46.Width = 26, ;
		Column46.ReadOnly = .T., ;
		Column46.Sparse = .F., ;
		Column46.ForeColor = RGB(0,0,0), ;
		Column46.BackColor = RGB(236,233,216), ;
		Column46.Name = "clmHold", ;
		Column47.ColumnOrder = 65, ;
		Column47.ControlSource = "webimp.tPurchase", ;
		Column47.Width = 140, ;
		Column47.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column47.ForeColor = RGB(0,0,0), ;
		Column47.BackColor = RGB(236,233,216), ;
		Column47.Name = "clmTPurchase", ;
		Column48.Alignment = 2, ;
		Column48.ColumnOrder = 56, ;
		Column48.ControlSource = "webimp.respersint", ;
		Column48.Width = 25, ;
		Column48.Sparse = .F., ;
		Column48.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column48.ForeColor = RGB(0,0,0), ;
		Column48.BackColor = RGB(236,233,216), ;
		Column48.Name = "clmResPersInt", ;
		Column49.Alignment = 2, ;
		Column49.ColumnOrder = 41, ;
		Column49.ControlSource = "webimp.bndl", ;
		Column49.Width = 25, ;
		Column49.Sparse = .F., ;
		Column49.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column49.ForeColor = RGB(0,0,0), ;
		Column49.BackColor = RGB(236,233,216), ;
		Column49.Name = "clmBndl", ;
		Column50.Alignment = 2, ;
		Column50.ColumnOrder = 38, ;
		Column50.ControlSource = "webimp.bodyinj", ;
		Column50.Width = 25, ;
		Column50.Sparse = .F., ;
		Column50.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column50.ForeColor = RGB(0,0,0), ;
		Column50.BackColor = RGB(236,233,216), ;
		Column50.Name = "clmBodyInj", ;
		Column51.FontBold = .T., ;
		Column51.FontName = "MS Sans Serif", ;
		Column51.Alignment = 2, ;
		Column51.ColumnOrder = 9, ;
		Column51.ControlSource = "webimp.cssn4", ;
		Column51.Width = 38, ;
		Column51.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column51.ForeColor = RGB(0,0,0), ;
		Column51.BackColor = RGB(236,233,216), ;
		Column51.Name = "clmCSSN4", ;
		Column52.FontBold = .T., ;
		Column52.FontName = "MS Sans Serif", ;
		Column52.ColumnOrder = 10, ;
		Column52.ControlSource = "webimp.cein", ;
		Column52.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column52.ForeColor = RGB(0,0,0), ;
		Column52.BackColor = RGB(236,233,216), ;
		Column52.Name = "clmCEIN", ;
		Column53.FontBold = .T., ;
		Column53.Alignment = 0, ;
		Column53.ColumnOrder = 6, ;
		Column53.ControlSource = "webimp.middlename", ;
		Column53.Width = 64, ;
		Column53.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column53.ForeColor = RGB(0,0,0), ;
		Column53.BackColor = RGB(236,233,216), ;
		Column53.Name = "clmMiddleName", ;
		Column54.ColumnOrder = 11, ;
		Column54.ControlSource = "webimp.cPoliLic", ;
		Column54.Width = 79, ;
		Column54.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column54.ForeColor = RGB(0,0,0), ;
		Column54.BackColor = RGB(236,233,216), ;
		Column54.Name = "clmCPoliLic", ;
		Column55.ColumnOrder = 39, ;
		Column55.ControlSource = "webimp.cConfLic", ;
		Column55.Width = 104, ;
		Column55.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column55.ForeColor = RGB(0,0,0), ;
		Column55.BackColor = RGB(236,233,216), ;
		Column55.Name = "clmcConfLic", ;
		Column56.ColumnOrder = 17, ;
		Column56.ControlSource = "webimp.zip4", ;
		Column56.Width = 37, ;
		Column56.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column56.ForeColor = RGB(0,0,0), ;
		Column56.BackColor = RGB(236,233,216), ;
		Column56.Name = "clmZIP4", ;
		Column57.ColumnOrder = 66, ;
		Column57.ControlSource = "webimp.cPolicyTyp", ;
		Column57.Width = 50, ;
		Column57.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column57.ForeColor = RGB(0,0,0), ;
		Column57.BackColor = RGB(236,233,216), ;
		Column57.Name = "clmcPolicyTyp", ;
		Column58.Alignment = 2, ;
		Column58.ColumnOrder = 67, ;
		Column58.ControlSource = "webimp.cCovType", ;
		Column58.Width = 47, ;
		Column58.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column58.ForeColor = RGB(0,0,0), ;
		Column58.BackColor = RGB(236,233,216), ;
		Column58.Name = "clmCovType", ;
		Column59.ColumnOrder = 21, ;
		Column59.ControlSource = "webimp.cCellPhn", ;
		Column59.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column59.ForeColor = RGB(0,0,0), ;
		Column59.BackColor = RGB(236,233,216), ;
		Column59.Name = "clmcCellPhn", ;
		Column60.Alignment = 2, ;
		Column60.ColumnOrder = 68, ;
		Column60.ControlSource = "webimp.lCommerc", ;
		Column60.Width = 45, ;
		Column60.Sparse = .F., ;
		Column60.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column60.ForeColor = RGB(0,0,0), ;
		Column60.BackColor = RGB(236,233,216), ;
		Column60.Name = "clmlCommerc", ;
		Column61.Alignment = 2, ;
		Column61.ColumnOrder = 36, ;
		Column61.ControlSource = "webimp.lClm5yr", ;
		Column61.Width = 25, ;
		Column61.Sparse = .F., ;
		Column61.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column61.ForeColor = RGB(0,0,0), ;
		Column61.BackColor = RGB(236,233,216), ;
		Column61.Name = "clmlClm5yr", ;
		Column62.Alignment = 2, ;
		Column62.ColumnOrder = 35, ;
		Column62.ControlSource = "webimp.lCurrClm", ;
		Column62.Width = 20, ;
		Column62.Sparse = .F., ;
		Column62.ForeColor = RGB(0,0,0), ;
		Column62.BackColor = RGB(236,233,216), ;
		Column62.Name = "clmlCurrClm", ;
		Column63.ColumnOrder = 69, ;
		Column63.ControlSource = "webimp.capprlic", ;
		Column63.Width = 63, ;
		Column63.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column63.ForeColor = RGB(0,0,0), ;
		Column63.BackColor = RGB(236,233,216), ;
		Column63.Name = "clmcApprLic", ;
		Column64.FontBold = .F., ;
		Column64.FontSize = 8, ;
		Column64.Alignment = 2, ;
		Column64.ColumnOrder = 22, ;
		Column64.ControlSource = "webimp.lTextMsg", ;
		Column64.Width = 20, ;
		Column64.Sparse = .F., ;
		Column64.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column64.ForeColor = RGB(0,0,0), ;
		Column64.BackColor = RGB(236,233,216), ;
		Column64.Name = "clmlTextMsg", ;
		Column65.ColumnOrder = 70, ;
		Column65.ControlSource = "webimp.cApprTrain", ;
		Column65.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column65.ForeColor = RGB(0,0,0), ;
		Column65.BackColor = RGB(236,233,216), ;
		Column65.Name = "clmcApprTrain", ;
		Column66.ColumnOrder = 71, ;
		Column66.ControlSource = "webimp.cTransact", ;
		Column66.Width = 56, ;
		Column66.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column66.ForeColor = RGB(0,0,0), ;
		Column66.BackColor = RGB(236,233,216), ;
		Column66.Name = "clmcTransact", ;
		Column67.ColumnOrder = 72, ;
		Column67.ControlSource = "webimp.cWebUser", ;
		Column67.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column67.ForeColor = RGB(0,0,0), ;
		Column67.BackColor = RGB(236,233,216), ;
		Column67.Name = "clmcWebUser", ;
		Column68.ColumnOrder = 73, ;
		Column68.ControlSource = "webimp.cPayeeName", ;
		Column68.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column68.ForeColor = RGB(0,0,0), ;
		Column68.BackColor = RGB(236,233,216), ;
		Column68.Name = "clmcPayeeName", ;
		Column69.ColumnOrder = 74, ;
		Column69.ControlSource = "webimp.cpmtid", ;
		Column69.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column69.ForeColor = RGB(0,0,0), ;
		Column69.BackColor = RGB(236,233,216), ;
		Column69.Name = "clmcPmtID", ;
		Column70.ColumnOrder = 75, ;
		Column70.ControlSource = "webimp.cpbname", ;
		Column70.Width = 83, ;
		Column70.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column70.ForeColor = RGB(0,0,0), ;
		Column70.BackColor = RGB(236,233,216), ;
		Column70.Name = "clmcPBname", ;
		Column71.ColumnOrder = 76, ;
		Column71.ControlSource = "webimp.cpbaddress", ;
		Column71.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column71.ForeColor = RGB(0,0,0), ;
		Column71.BackColor = RGB(236,233,216), ;
		Column71.Name = "clmcPBaddress", ;
		Column72.ColumnOrder = 77, ;
		Column72.ControlSource = "webimp.cpbcity", ;
		Column72.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column72.ForeColor = RGB(0,0,0), ;
		Column72.BackColor = RGB(236,233,216), ;
		Column72.Name = "clmcPBcity", ;
		Column73.ColumnOrder = 78, ;
		Column73.ControlSource = "webimp.cpbstate", ;
		Column73.Width = 29, ;
		Column73.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column73.ForeColor = RGB(0,0,0), ;
		Column73.BackColor = RGB(236,233,216), ;
		Column73.Name = "clmcPBstate", ;
		Column74.ColumnOrder = 79, ;
		Column74.ControlSource = "webimp.cpbzip", ;
		Column74.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column74.ForeColor = RGB(0,0,0), ;
		Column74.BackColor = RGB(236,233,216), ;
		Column74.Name = "clmcPBzip", ;
		Column75.ColumnOrder = 80, ;
		Column75.ControlSource = "webimp.ctxctauth", ;
		Column75.Width = 91, ;
		Column75.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column75.ForeColor = RGB(0,0,0), ;
		Column75.BackColor = RGB(236,233,216), ;
		Column75.Name = "Column9", ;
		Column76.ColumnOrder = 81, ;
		Column76.ControlSource = "webimp.ctxctcode", ;
		Column76.Width = 47, ;
		Column76.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column76.ForeColor = RGB(0,0,0), ;
		Column76.BackColor = RGB(236,233,216), ;
		Column76.Name = "clmctxctcode", ;
		Column77.ColumnOrder = 82, ;
		Column77.ControlSource = "webimp.ntxctrate", ;
		Column77.Width = 44, ;
		Column77.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column77.ForeColor = RGB(0,0,0), ;
		Column77.BackColor = RGB(236,233,216), ;
		Column77.Name = "clmntxctrate", ;
		Column78.ColumnOrder = 83, ;
		Column78.ControlSource = "webimp.ntxctamt", ;
		Column78.Width = 41, ;
		Column78.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column78.ForeColor = RGB(0,0,0), ;
		Column78.BackColor = RGB(236,233,216), ;
		Column78.Name = "clmntxctamt", ;
		Column79.ColumnOrder = 84, ;
		Column79.ControlSource = "webimp.ctxcoauth", ;
		Column79.Width = 83, ;
		Column79.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column79.ForeColor = RGB(0,0,0), ;
		Column79.BackColor = RGB(236,233,216), ;
		Column79.Name = "clmctxcoauth", ;
		Column80.ColumnOrder = 85, ;
		Column80.ControlSource = "webimp.ctxcocode", ;
		Column80.Width = 63, ;
		Column80.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column80.ForeColor = RGB(0,0,0), ;
		Column80.BackColor = RGB(236,233,216), ;
		Column80.Name = "clmctxcocode", ;
		Column81.ColumnOrder = 86, ;
		Column81.ControlSource = "webimp.ntxcorate", ;
		Column81.Width = 60, ;
		Column81.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column81.ForeColor = RGB(0,0,0), ;
		Column81.BackColor = RGB(236,233,216), ;
		Column81.Name = "clmntxcorate", ;
		Column82.ColumnOrder = 87, ;
		Column82.ControlSource = "webimp.ntxcoamt", ;
		Column82.Width = 57, ;
		Column82.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column82.ForeColor = RGB(0,0,0), ;
		Column82.BackColor = RGB(236,233,216), ;
		Column82.Name = "clmntxcoamt", ;
		Column83.ColumnOrder = 88, ;
		Column83.ControlSource = "webimp.surcharge", ;
		Column83.Width = 55, ;
		Column83.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column83.ForeColor = RGB(0,0,0), ;
		Column83.BackColor = RGB(236,233,216), ;
		Column83.Name = "clmsurcharge", ;
		Column84.FontName = "MS Sans Serif", ;
		Column84.Alignment = 2, ;
		Column84.ColumnOrder = 57, ;
		Column84.ControlSource = "webimp.devcon", ;
		Column84.Width = 25, ;
		Column84.Sparse = .F., ;
		Column84.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column84.ForeColor = RGB(0,0,0), ;
		Column84.BackColor = RGB(236,233,216), ;
		Column84.Name = "clmdevcon", ;
		Column85.FontName = "MS Sans Serif", ;
		Column85.Alignment = 2, ;
		Column85.ColumnOrder = 58, ;
		Column85.ControlSource = "webimp.subpoena", ;
		Column85.Width = 25, ;
		Column85.Sparse = .F., ;
		Column85.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column85.ForeColor = RGB(0,0,0), ;
		Column85.BackColor = RGB(236,233,216), ;
		Column85.Name = "clmsubpoena", ;
		Column86.ColumnOrder = 59, ;
		Column86.ControlSource = "webimp.rm10", ;
		Column86.Width = 25, ;
		Column86.Sparse = .F., ;
		Column86.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column86.ForeColor = RGB(0,0,0), ;
		Column86.BackColor = RGB(236,233,216), ;
		Column86.Name = "clmRM10", ;
		Column87.ColumnOrder = 60, ;
		Column87.ControlSource = "webimp.rm20", ;
		Column87.Width = 25, ;
		Column87.Sparse = .F., ;
		Column87.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column87.ForeColor = RGB(0,0,0), ;
		Column87.BackColor = RGB(236,233,216), ;
		Column87.Name = "clmRM20", ;
		Column88.ColumnOrder = 61, ;
		Column88.ControlSource = "webimp.lockbox", ;
		Column88.Width = 25, ;
		Column88.Sparse = .F., ;
		Column88.DynamicBackColor = "IIF(MOD(RECNO(),2)=1,RGB(255,255,255),RGB(192,220,192))", ;
		Column88.ForeColor = RGB(0,0,0), ;
		Column88.BackColor = RGB(236,233,216), ;
		Column88.Name = "clmLockBox"


	ADD OBJECT dataimport.grdwebimp.clmpolinumb.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontShadow = .F., ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "*PoliNumb*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Policy#:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmpolinumb.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.polinumb", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmpdanum.header1 AS header WITH ;
		FontBold = .T., ;
		FontItalic = .F., ;
		FontName = "MS Sans Serif", ;
		FontOutline = .F., ;
		FontSize = 8, ;
		FontUnderline = .F., ;
		FontCondense = .F., ;
		Alignment = 2, ;
		Caption = "*PDANum*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "PDA#:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmpdanum.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.pdanum", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmpoliid.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "*PoliID*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Policy ID:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmpoliid.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.poliid", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmfirstname.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "*First*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "First name:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmfirstname.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.firstname", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmlastname.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "*Last*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Last (Firm) name:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmlastname.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.lastname", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcompany.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "*Company*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Company (Firm) name:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcompany.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.company", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmaddress1.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "*Address1*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Address 1:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmaddress1.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.address1", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcity.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "*City*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "City:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcity.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.city", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmstate.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "*St*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "State:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmstate.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.state", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmzip.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 1, ;
		Caption = "*ZIP*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "ZIP Code 5:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmzip.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.zip", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcounty.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "County", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "County", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcounty.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.county", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmphone1.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "*Phone#*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Phone#:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmphone1.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.phone1", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmphone1ext.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "Ext#", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Phone Extension#", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmphone1ext.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.phone1ext", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmfax.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "FAX#", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "FAX#", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmfax.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.fax", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmorigdate.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "OrigDate", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Original date of continuous coverage", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmorigdate.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.origdate", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmeffective.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "Effective", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Effective date of policy", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmeffective.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.effective", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmend.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "End", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "End date of policy", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmend.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.end", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmpremimport.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 1, ;
		Caption = "PremImport", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Total amount:  Premium + Endorrsements +OnlineChg", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmpremimport.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 1, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.premimport", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmprem.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 1, ;
		Caption = "Premium", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Premium Policy Amount", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmprem.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 1, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.prem", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmtaxes.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 1, ;
		Caption = "Taxes", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Tax on total", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmtaxes.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 1, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.taxes", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmentrydate.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "EntryDate", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Record entered on <date>", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmentrydate.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.entrydate", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmentryby.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "EntryBy", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Record entered by <named>", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmentryby.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.entryby", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmentrytime.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "EntryTime", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Record entered at <time>", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmentrytime.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.entrytime", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmemail.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "*Email->*", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Email:  [Dbl+Click]=order, [Click]=no order, scroll-right for more", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmemail.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.email", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmclaimalert.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		FontCharSet = 0, ;
		Alignment = 2, ;
		Caption = "Clm", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Had previous claim?", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmclaimalert.check1 AS checkbox WITH ;
		Top = 63, ;
		Left = 16, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		FontSize = 8, ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.claimalert", ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmhigherlim1.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "HL1", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "HIGHER LIMIT - 1", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmhigherlim1.check1 AS checkbox WITH ;
		Top = 99, ;
		Left = 25, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.higherlim1", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmhigherlim2.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "HL2", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "HIGHER LIMIT - 2", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmhigherlim2.check1 AS checkbox WITH ;
		Top = 99, ;
		Left = 22, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.higherlim2", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmappraisal.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		FontCharSet = 0, ;
		Alignment = 2, ;
		Caption = "Appr", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "APPRAISAL", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmappraisal.check1 AS checkbox WITH ;
		Top = 75, ;
		Left = 18, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.appraisal", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmpropmgr.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "Prop", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "LEASING & PROPERTY MGMT", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmpropmgr.check1 AS checkbox WITH ;
		Top = 111, ;
		Left = 19, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.propmgr", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmenvironx.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "EnvD", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "ENVIRONMENTAL DMG[/CE] (Damages [Claims Expenses])", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmenvironx.check1 AS checkbox WITH ;
		Top = 75, ;
		Left = 26, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.environx", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmfairhouse.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "FairHs", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "FAIR HOUSING", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmfairhouse.check1 AS checkbox WITH ;
		Top = 87, ;
		Left = 19, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.fairhouse", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmcomplaints.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "RegC", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "REGULATORY COMPLAINT", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcomplaints.check1 AS checkbox WITH ;
		Top = 99, ;
		Left = 31, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.complaints", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmonlinechg.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "Chg", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "ONLINE CHARGE", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmonlinechg.check1 AS checkbox WITH ;
		Top = 87, ;
		Left = 16, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.onlinechg", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmearnmo.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "EarnMo", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "EARNEST MONEY DISPUTE", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmearnmo.check1 AS checkbox WITH ;
		Top = 123, ;
		Left = 12, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.earnmo", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmlcenviro.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "EnvC", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "ENVIRONMENTAL CE (Claims Expenses)", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmlcenviro.check1 AS checkbox WITH ;
		Top = 99, ;
		Left = 8, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.lcenviro", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmfairhausa.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "FairA", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "FAIR HOUSING (A)", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmfairhausa.check1 AS checkbox WITH ;
		Top = 99, ;
		Left = 22, ;
		Height = 17, ;
		Width = 60, ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.fairhausa", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmfairhausb.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "FairB", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "FAIR HOUSING (B)", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmfairhausb.check1 AS checkbox WITH ;
		Top = 87, ;
		Left = 17, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.fairhausb", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmprimeres.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "P-Res", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "PRIMARY RESIDENCE", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmprimeres.check1 AS checkbox WITH ;
		Top = 99, ;
		Left = 15, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.primeres", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmhilimmil1.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "HLM1", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "HIGHER LIMIT - 1M", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmhilimmil1.check1 AS checkbox WITH ;
		Top = 99, ;
		Left = 23, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.hilimmil1", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmhilimmil2.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "HLM2", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "HIGHER LIMIT - 2M", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmhilimmil2.check1 AS checkbox WITH ;
		Top = 87, ;
		Left = 22, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.hilimmil2", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmgovstate.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "Arial", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "Gov", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Governing State", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmgovstate.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.govstate", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmbatchnum.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "Arial", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "Batch#", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Batch Number", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmbatchnum.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.batchnum", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmaddress2.header1 AS header WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		Caption = "Address2", ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(192,220,192), ;
		ToolTipText = "Address 2", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmaddress2.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.address2", ;
		Margin = 0, ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmpremiuma.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "PrmA", ;
		ToolTipText = "PREMIUM A", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmpremiuma.check1 AS checkbox WITH ;
		Top = 75, ;
		Left = 12, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.premiuma", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmpremiumb.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "PrmB", ;
		ToolTipText = "PREMIUM B", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmpremiumb.check1 AS checkbox WITH ;
		Top = 87, ;
		Left = 24, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.premiumb", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmhold.hold AS header WITH ;
		FontBold = .F., ;
		FontItalic = .T., ;
		FontSize = 8, ;
		Caption = "Hold", ;
		ToolTipText = "Will not import when placed on Hold.", ;
		Name = "Hold"


	ADD OBJECT dataimport.grdwebimp.clmhold.check1 AS checkbox WITH ;
		Top = 35, ;
		Left = 7, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.lhold", ;
		BackColor = RGB(236,233,216), ;
		ReadOnly = .F., ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmtpurchase.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "Arial", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "Purchased", ;
		ToolTipText = "Enrollment Date/Time     ", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmtpurchase.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "webimp.tpurchase", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmrespersint.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "ResP", ;
		ToolTipText = "RESIDENTIAL PERSONAL INTEREST", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmrespersint.check1 AS checkbox WITH ;
		Top = 75, ;
		Left = 34, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.respersint", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmbndl.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "Bndl", ;
		ToolTipText = "ENDORSEMENT BUNDLE", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmbndl.check1 AS checkbox WITH ;
		Top = 75, ;
		Left = 43, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.bndl", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmbodyinj.header1 AS header WITH ;
		FontBold = .F., ;
		FontItalic = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		FontCharSet = 0, ;
		Alignment = 2, ;
		Caption = "Body", ;
		ToolTipText = "BODILY INJURY/PROPERTY DAMAGE", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmbodyinj.check1 AS checkbox WITH ;
		Top = 99, ;
		Left = 27, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "webimp.bodyinj", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmcssn4.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "*SS#4*", ;
		ToolTipText = "Social Security last-4:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcssn4.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.cssn4", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcein.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "*EIN*", ;
		ToolTipText = "Employer Identification Number:  [Dbl+Click]=order, [Click]=no order", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcein.text1 AS textbox WITH ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		BorderStyle = 0, ;
		ControlSource = "webimp.cein", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmmiddlename.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 8, ;
		Caption = "Middle", ;
		ToolTipText = "Middle name", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmmiddlename.text1 AS textbox WITH ;
		FontBold = .T., ;
		Alignment = 0, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.middlename", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcpolilic.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 8, ;
		Caption = "Policy Lic#s->", ;
		ToolTipText = "Policy Licenses: <State>.<Type>.<Number>;", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcpolilic.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "webimp.cPoliLic", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcconflic.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "Arial", ;
		FontSize = 8, ;
		FontCharSet = 0, ;
		Caption = "Conformity Lic#s->", ;
		ToolTipText = "Conformity Licenses: <State>.<Type>.<Number>;", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcconflic.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmzip4.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 8, ;
		Caption = "+4", ;
		ToolTipText = "ZIP Code +4", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmzip4.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "webimp.zip4", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcpolicytyp.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "Arial", ;
		FontSize = 8, ;
		Caption = "PoliType", ;
		ToolTipText = "Policy Type:  Individual, Firm", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcpolicytyp.text1 AS textbox WITH ;
		FontSize = 8, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.cpolicytyp", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcovtype.header1 AS header WITH ;
		FontName = "Arial", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "CovType", ;
		ToolTipText = "Coverage Type: EO, AP", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcovtype.text1 AS textbox WITH ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "webimp.cCovType", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmccellphn.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "Cell#", ;
		ToolTipText = "Cell phone#", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmccellphn.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "webimp.cCellPhn", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmlcommerc.header1 AS header WITH ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "Comm75", ;
		ToolTipText = "Commercial 75%+ of sales?", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmlcommerc.check1 AS checkbox WITH ;
		Top = 63, ;
		Left = 25, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmlclm5yr.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		FontCharSet = 0, ;
		Alignment = 2, ;
		Caption = "Clm5", ;
		ToolTipText = "Had claim in last 5 yrs?", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmlclm5yr.check1 AS checkbox WITH ;
		Top = 75, ;
		Left = 27, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmlcurrclm.header1 AS header WITH ;
		FontBold = .F., ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		FontCharSet = 0, ;
		Alignment = 2, ;
		Caption = "Curr", ;
		ToolTipText = "Have current claim?", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmlcurrclm.check1 AS checkbox WITH ;
		Top = 75, ;
		Left = 18, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmcapprlic.header1 AS header WITH ;
		FontName = "Arial", ;
		FontSize = 8, ;
		Caption = "ApprLic#s->", ;
		ToolTipText = "Appraisor Licenses: <State>.<Type>.<Number>;", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcapprlic.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "webimp.capprlic", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmltextmsg.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 8, ;
		Caption = "Txt", ;
		ToolTipText = "Accept Cell Text Messages?", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmltextmsg.check1 AS checkbox WITH ;
		Top = 63, ;
		Left = 46, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		FontSize = 8, ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		BackColor = RGB(236,233,216), ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmcapprtrain.header1 AS header WITH ;
		FontName = "Arial", ;
		FontSize = 8, ;
		Caption = "ApprTrainee->", ;
		ToolTipText = "Appraiser Trainees:  <Full Name>;", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcapprtrain.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "webimp.cApprTrain", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmctransact.header1 AS header WITH ;
		FontName = "Arial", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "TransactID", ;
		ToolTipText = "Transaction ID from Authorize.NET", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmctransact.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcwebuser.header1 AS header WITH ;
		FontName = "Arial", ;
		FontSize = 8, ;
		Caption = "WebUser->", ;
		ToolTipText = "Web Account Username (email address)", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcwebuser.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcpayeename.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Payee", ;
		ToolTipText = "Payee:  Name on credit card", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcpayeename.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "webimp.cPayeeName", ;
		Margin = 0, ;
		ReadOnly = .F., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcpmtid.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "PaymentID", ;
		ToolTipText = "Payment ID", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcpmtid.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "webimp.cpmtid", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcpbname.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Principle Broker", ;
		BackColor = RGB(236,233,216), ;
		ToolTipText = "Principle Broker Name", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcpbname.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcpbaddress.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Address", ;
		ToolTipText = "Principle Broker Address", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcpbaddress.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcpbcity.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "City", ;
		ToolTipText = "Principle Broker City", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcpbcity.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcpbstate.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "State", ;
		ToolTipText = "Principle Broker State", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcpbstate.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmcpbzip.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "ZIP", ;
		ToolTipText = "Principle Broker ZIP", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmcpbzip.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.column9.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "City Name", ;
		ToolTipText = "Taxing City Name", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.column9.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmctxctcode.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "CityCode", ;
		ToolTipText = "Taxing City Code", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmctxctcode.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmntxctrate.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 1, ;
		Caption = "CityRate", ;
		ToolTipText = "Taxing City Rate", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmntxctrate.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmntxctamt.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 1, ;
		Caption = "CityAmt", ;
		ToolTipText = "Taxing City Amount", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmntxctamt.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmctxcoauth.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "County Name", ;
		ToolTipText = "Taxing County Name", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmctxcoauth.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmctxcocode.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "CountyCode", ;
		ToolTipText = "Taxing County Code", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmctxcocode.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmntxcorate.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "CountyRate", ;
		ToolTipText = "Taxing County Rate", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmntxcorate.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmntxcoamt.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "CountyAmt", ;
		ToolTipText = "Taxing County Amount", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmntxcoamt.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmsurcharge.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 1, ;
		Caption = "Surcharge", ;
		ToolTipText = "Surcharge Amount", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmsurcharge.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(236,233,216), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdwebimp.clmdevcon.header1 AS header WITH ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "DevC", ;
		ToolTipText = "DEVELOPED/CONSTRUCTED", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmdevcon.check1 AS checkbox WITH ;
		Top = 71, ;
		Left = 26, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmsubpoena.header1 AS header WITH ;
		FontName = "MS Sans Serif", ;
		FontSize = 9, ;
		Alignment = 2, ;
		Caption = "SubP", ;
		ToolTipText = "SUBPOENA", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmsubpoena.check1 AS checkbox WITH ;
		Top = 35, ;
		Left = 34, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmrm10.header1 AS header WITH ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "R10", ;
		ToolTipText = "REVERSE MORTGAGE LOAN TRANSACTION - DMG - 10K DED           ", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmrm10.check1 AS checkbox WITH ;
		Top = 83, ;
		Left = 21, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmrm20.header1 AS header WITH ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "R20", ;
		ToolTipText = "REVERSE MORTGAGE LOAN TRANSACTION - DMG - 20K DED           ", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmrm20.check1 AS checkbox WITH ;
		Top = 83, ;
		Left = 29, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		Name = "Check1"


	ADD OBJECT dataimport.grdwebimp.clmlockbox.header1 AS header WITH ;
		FontName = "MS Sans Serif", ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "LkBx", ;
		ToolTipText = "LOCK BOX LIABILITY - DMG/CE- 5K/10K                         ", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdwebimp.clmlockbox.check1 AS checkbox WITH ;
		Top = 71, ;
		Left = 25, ;
		Height = 17, ;
		Width = 60, ;
		FontBold = .T., ;
		FontName = "MS Sans Serif", ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		Name = "Check1"


	ADD OBJECT cmdholdsonly AS commandbutton WITH ;
		AutoSize = .T., ;
		Top = 2, ;
		Left = 5, ;
		Height = 25, ;
		Width = 66, ;
		FontSize = 8, ;
		Caption = "Holds only", ;
		TabIndex = 1, ;
		ToolTipText = "View only records on Hold.", ;
		Visible = .F., ;
		Name = "cmdHoldsOnly"


	ADD OBJECT cmdrejects AS commandbutton WITH ;
		AutoSize = .T., ;
		Top = 2, ;
		Left = 1237, ;
		Height = 25, ;
		Width = 52, ;
		FontSize = 8, ;
		Caption = "Rejects", ;
		TabIndex = 11, ;
		ToolTipText = "View/edit previously reject import records", ;
		Name = "cmdRejects"


	ADD OBJECT shape1 AS shape WITH ;
		Top = 4, ;
		Left = 93, ;
		Height = 20, ;
		Width = 20, ;
		Curvature = 99, ;
		BackColor = RGB(192,192,192), ;
		Name = "Shape1"


	ADD OBJECT label1 AS label WITH ;
		AutoSize = .T., ;
		FontBold = .T., ;
		Caption = "1", ;
		Height = 17, ;
		Left = 100, ;
		Top = 6, ;
		Width = 9, ;
		TabIndex = 20, ;
		ForeColor = RGB(255,255,255), ;
		BackColor = RGB(192,192,192), ;
		Name = "Label1"


	ADD OBJECT shape2 AS shape WITH ;
		Top = 4, ;
		Left = 413, ;
		Height = 20, ;
		Width = 20, ;
		Curvature = 99, ;
		BackColor = RGB(192,192,192), ;
		Name = "Shape2"


	ADD OBJECT label2 AS label WITH ;
		AutoSize = .T., ;
		FontBold = .T., ;
		Caption = "2", ;
		Height = 17, ;
		Left = 420, ;
		Top = 6, ;
		Width = 9, ;
		TabIndex = 21, ;
		ForeColor = RGB(255,255,255), ;
		BackColor = RGB(192,192,192), ;
		Name = "Label2"


	ADD OBJECT shape3 AS shape WITH ;
		Top = 4, ;
		Left = 898, ;
		Height = 20, ;
		Width = 20, ;
		Curvature = 99, ;
		BackColor = RGB(192,192,192), ;
		Name = "Shape3"


	ADD OBJECT label3 AS label WITH ;
		AutoSize = .T., ;
		FontBold = .T., ;
		Caption = "3", ;
		Height = 17, ;
		Left = 905, ;
		Top = 6, ;
		Width = 9, ;
		TabIndex = 22, ;
		ForeColor = RGB(255,255,255), ;
		BackColor = RGB(192,192,192), ;
		Name = "Label3"


	ADD OBJECT shape4 AS shape WITH ;
		Top = 4, ;
		Left = 1005, ;
		Height = 20, ;
		Width = 20, ;
		Curvature = 99, ;
		BackColor = RGB(192,192,192), ;
		Name = "Shape4"


	ADD OBJECT label4 AS label WITH ;
		AutoSize = .T., ;
		FontBold = .T., ;
		Caption = "4", ;
		Height = 17, ;
		Left = 1012, ;
		Top = 6, ;
		Width = 9, ;
		TabIndex = 23, ;
		ForeColor = RGB(255,255,255), ;
		BackColor = RGB(192,192,192), ;
		Name = "Label4"


	ADD OBJECT cmdclear AS commandbutton WITH ;
		AutoSize = .T., ;
		Top = 2, ;
		Left = 1131, ;
		Height = 25, ;
		Width = 41, ;
		FontSize = 8, ;
		Caption = "Clear", ;
		Enabled = .F., ;
		TabIndex = 9, ;
		ToolTipText = "Clear the current imported records.", ;
		Name = "cmdClear"


	ADD OBJECT btnnoteedit AS commandbutton WITH ;
		AutoSize = .T., ;
		Top = 2, ;
		Left = 801, ;
		Height = 25, ;
		Width = 57, ;
		FontSize = 8, ;
		Caption = "AltNote?", ;
		Enabled = .F., ;
		TabIndex = 6, ;
		ToolTipText = "Enter a note for placement in each imported record's Notes.", ;
		Name = "btnNoteEdit"


	ADD OBJECT grdnameref AS grid WITH ;
		ColumnCount = 30, ;
		Anchor = 15, ;
		DeleteMark = .F., ;
		Height = 93, ;
		Left = 5, ;
		Panel = 1, ;
		ReadOnly = .F., ;
		RecordSource = "NameRef", ;
		TabIndex = 17, ;
		ToolTipText = "Match = GovState + Last (left 2) + First (left 3)", ;
		Top = 621, ;
		Width = 1500, ;
		Optimize = .T., ;
		Name = "grdNameRef", ;
		Column1.ColumnOrder = 2, ;
		Column1.ControlSource = "NameRef.polinumb", ;
		Column1.Width = 110, ;
		Column1.ReadOnly = .F., ;
		Column1.Visible = .T., ;
		Column1.Name = "clmPoliNumb", ;
		Column2.ColumnOrder = 4, ;
		Column2.ControlSource = "NameRef.poliid", ;
		Column2.Width = 65, ;
		Column2.ReadOnly = .T., ;
		Column2.Visible = .T., ;
		Column2.Name = "clmPoliID", ;
		Column3.ColumnOrder = 5, ;
		Column3.ControlSource = "NameRef.firstname", ;
		Column3.Width = 154, ;
		Column3.ReadOnly = .T., ;
		Column3.Visible = .T., ;
		Column3.Name = "clmFirstName", ;
		Column4.ColumnOrder = 7, ;
		Column4.ControlSource = "NameRef.company", ;
		Column4.Width = 90, ;
		Column4.ReadOnly = .T., ;
		Column4.Visible = .T., ;
		Column4.Name = "clmCompany", ;
		Column5.ColumnOrder = 11, ;
		Column5.ControlSource = "NameRef.tipe", ;
		Column5.Width = 30, ;
		Column5.ReadOnly = .T., ;
		Column5.Visible = .T., ;
		Column5.Name = "clmTipe", ;
		Column6.ColumnOrder = 10, ;
		Column6.ControlSource = "NameRef.license", ;
		Column6.Width = 60, ;
		Column6.ReadOnly = .T., ;
		Column6.Visible = .T., ;
		Column6.Name = "clmLicense", ;
		Column7.ColumnOrder = 12, ;
		Column7.ControlSource = "NameRef.address1", ;
		Column7.ReadOnly = .T., ;
		Column7.Visible = .T., ;
		Column7.Name = "clmAddress1", ;
		Column8.ColumnOrder = 14, ;
		Column8.ControlSource = "NameRef.city", ;
		Column8.ReadOnly = .T., ;
		Column8.Visible = .T., ;
		Column8.Name = "clmCity", ;
		Column9.Alignment = 2, ;
		Column9.ColumnOrder = 15, ;
		Column9.ControlSource = "NameRef.state", ;
		Column9.Width = 30, ;
		Column9.ReadOnly = .T., ;
		Column9.Visible = .T., ;
		Column9.Name = "clmState", ;
		Column10.ColumnOrder = 16, ;
		Column10.ControlSource = "NameRef.zip", ;
		Column10.Width = 40, ;
		Column10.ReadOnly = .T., ;
		Column10.Visible = .T., ;
		Column10.Name = "clmZIP", ;
		Column11.ColumnOrder = 17, ;
		Column11.ControlSource = "NameRef.county", ;
		Column11.ReadOnly = .T., ;
		Column11.Visible = .T., ;
		Column11.Name = "clmCounty", ;
		Column12.ColumnOrder = 18, ;
		Column12.ControlSource = "NameRef.phone1", ;
		Column12.Width = 90, ;
		Column12.ReadOnly = .T., ;
		Column12.Visible = .T., ;
		Column12.Name = "clmPhone1", ;
		Column13.ColumnOrder = 20, ;
		Column13.ControlSource = "NameRef.ccellphn", ;
		Column13.Width = 90, ;
		Column13.ReadOnly = .T., ;
		Column13.Visible = .T., ;
		Column13.Name = "clmPhone2", ;
		Column14.ColumnOrder = 21, ;
		Column14.ControlSource = "NameRef.fax", ;
		Column14.Width = 90, ;
		Column14.ReadOnly = .T., ;
		Column14.Visible = .T., ;
		Column14.Name = "clmFAX", ;
		Column15.ColumnOrder = 24, ;
		Column15.ControlSource = "NameRef.origdate", ;
		Column15.Width = 68, ;
		Column15.ReadOnly = .T., ;
		Column15.Visible = .T., ;
		Column15.Name = "clmOrigDate", ;
		Column16.ColumnOrder = 22, ;
		Column16.ControlSource = "NameRef.effective", ;
		Column16.Width = 68, ;
		Column16.ReadOnly = .T., ;
		Column16.Visible = .T., ;
		Column16.Name = "clmEffective", ;
		Column17.ColumnOrder = 23, ;
		Column17.ControlSource = "NameRef.end", ;
		Column17.Width = 68, ;
		Column17.ReadOnly = .T., ;
		Column17.Visible = .T., ;
		Column17.Name = "clmEnd", ;
		Column18.ColumnOrder = 25, ;
		Column18.ControlSource = "NameRef.prem", ;
		Column18.ReadOnly = .T., ;
		Column18.Visible = .T., ;
		Column18.Name = "clmPrem", ;
		Column19.ColumnOrder = 26, ;
		Column19.ControlSource = "NameRef.taxes", ;
		Column19.Width = 40, ;
		Column19.ReadOnly = .T., ;
		Column19.Visible = .T., ;
		Column19.Name = "clmTaxes", ;
		Column20.ColumnOrder = 27, ;
		Column20.ControlSource = "NameRef.entrydate", ;
		Column20.Width = 68, ;
		Column20.ReadOnly = .T., ;
		Column20.Visible = .T., ;
		Column20.Name = "clmEntryDate", ;
		Column21.ColumnOrder = 28, ;
		Column21.ControlSource = "NameRef.entryby", ;
		Column21.ReadOnly = .T., ;
		Column21.Visible = .T., ;
		Column21.Name = "clmEntryBy", ;
		Column22.ColumnOrder = 29, ;
		Column22.ControlSource = "NameRef.entrytime", ;
		Column22.Width = 55, ;
		Column22.ReadOnly = .T., ;
		Column22.Visible = .T., ;
		Column22.Name = "clmEntryTime", ;
		Column23.ColumnOrder = 30, ;
		Column23.ControlSource = "NameRef.email", ;
		Column23.ReadOnly = .T., ;
		Column23.Visible = .T., ;
		Column23.Name = "clmEmail", ;
		Column24.ColumnOrder = 6, ;
		Column24.ControlSource = "NameRef.lastname", ;
		Column24.Width = 90, ;
		Column24.ReadOnly = .T., ;
		Column24.Visible = .T., ;
		Column24.Name = "clmLastName", ;
		Column25.ColumnOrder = 13, ;
		Column25.ControlSource = "NameRef.address2", ;
		Column25.ReadOnly = .T., ;
		Column25.Visible = .T., ;
		Column25.Name = "clmAddress2", ;
		Column26.ColumnOrder = 1, ;
		Column26.ControlSource = "NameRef.llink", ;
		Column26.Width = 22, ;
		Column26.ReadOnly = .F., ;
		Column26.Sparse = .F., ;
		Column26.Visible = .T., ;
		Column26.Name = "clmLink", ;
		Column27.Alignment = 2, ;
		Column27.ColumnOrder = 8, ;
		Column27.ControlSource = "NameRef.cssn4", ;
		Column27.Width = 36, ;
		Column27.ReadOnly = .T., ;
		Column27.Name = "clmSSN4", ;
		Column28.Alignment = 2, ;
		Column28.ColumnOrder = 9, ;
		Column28.ControlSource = "NameRef.cein", ;
		Column28.ReadOnly = .T., ;
		Column28.Name = "Column4", ;
		Column29.ColumnOrder = 3, ;
		Column29.ControlSource = "NameRef.pdanum", ;
		Column29.Width = 90, ;
		Column29.ReadOnly = .F., ;
		Column29.Name = "clmPDAnum", ;
		Column30.ColumnOrder = 19, ;
		Column30.ControlSource = "NameRef.cphnext1", ;
		Column30.ReadOnly = .F., ;
		Column30.Name = "clmcPhnExt1"


	ADD OBJECT dataimport.grdnameref.clmpolinumb.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "Polinumb", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmpolinumb.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.polinumb", ;
		Margin = 0, ;
		ReadOnly = .F., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmpoliid.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "PoliID", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmpoliid.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "PI_Lic.poliid", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmfirstname.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 8, ;
		FontUnderline = .T., ;
		Alignment = 0, ;
		Caption = "First (Middle)", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmfirstname.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.firstname", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmcompany.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Company", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmcompany.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.company", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmtipe.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Tipe", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmtipe.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.tipe", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmlicense.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "License", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmlicense.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.license", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmaddress1.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Address1", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmaddress1.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.address1", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmcity.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "City", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmcity.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.city", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmstate.header1 AS header WITH ;
		FontBold = .T., ;
		FontShadow = .F., ;
		FontSize = 8, ;
		FontUnderline = .T., ;
		Alignment = 2, ;
		Caption = "St", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmstate.text1 AS textbox WITH ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.state", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmzip.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "ZIP", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmzip.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.zip", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmcounty.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "County", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmcounty.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.county", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmphone1.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "Phone#", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmphone1.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.phone1", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmphone2.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "Cell#", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmphone2.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.phone2", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmfax.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "FAX", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmfax.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.fax", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmorigdate.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "OrigDate", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmorigdate.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.origdate", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmeffective.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "Effective", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmeffective.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.effective", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmend.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "End", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmend.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.end", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmprem.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 1, ;
		Caption = "Premium", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmprem.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.prem", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmtaxes.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 1, ;
		Caption = "Taxes", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmtaxes.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.taxes", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmentrydate.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "EntryDate", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmentrydate.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.entrydate", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmentryby.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "EntryBy", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmentryby.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.entryby", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmentrytime.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "EntryTime", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmentrytime.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.entrytime", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmemail.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "eMail", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmemail.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.email", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmlastname.header1 AS header WITH ;
		FontBold = .T., ;
		FontSize = 8, ;
		FontUnderline = .T., ;
		Alignment = 0, ;
		Caption = "Last", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmlastname.text1 AS textbox WITH ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.lastname", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmaddress2.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Address2", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmaddress2.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.address2", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		Visible = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmlink.link AS header WITH ;
		FontItalic = .T., ;
		FontSize = 8, ;
		Caption = "Link", ;
		Name = "Link"


	ADD OBJECT dataimport.grdnameref.clmlink.check1 AS checkbox WITH ;
		Top = 23, ;
		Left = 8, ;
		Height = 17, ;
		Width = 60, ;
		Alignment = 2, ;
		Centered = .T., ;
		Caption = "", ;
		ControlSource = "NameRef.llink", ;
		BackColor = RGB(236,233,216), ;
		ReadOnly = .F., ;
		Name = "Check1"


	ADD OBJECT dataimport.grdnameref.clmssn4.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "SS#4", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmssn4.text1 AS textbox WITH ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.cssn4", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.column4.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "EIN", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.column4.text1 AS textbox WITH ;
		Alignment = 2, ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.cein", ;
		Margin = 0, ;
		ReadOnly = .T., ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmpdanum.header1 AS header WITH ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "PDANum", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmpdanum.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.pdanum", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT dataimport.grdnameref.clmcphnext1.header1 AS header WITH ;
		FontSize = 8, ;
		Caption = "Ext#", ;
		Name = "Header1"


	ADD OBJECT dataimport.grdnameref.clmcphnext1.text1 AS textbox WITH ;
		BorderStyle = 0, ;
		ControlSource = "NameRef.cphnext1", ;
		Margin = 0, ;
		ForeColor = RGB(0,0,0), ;
		BackColor = RGB(255,255,255), ;
		Name = "Text1"


	ADD OBJECT cmdbrowseedit AS commandbutton WITH ;
		AutoSize = .T., ;
		Top = 2, ;
		Left = 1187, ;
		Height = 25, ;
		Width = 33, ;
		FontSize = 8, ;
		Caption = "Edit", ;
		Enabled = .F., ;
		TabIndex = 10, ;
		ToolTipText = ['"Edit"'], ;
		Visible = .F., ;
		Name = "cmdBrowseEdit"


	ADD OBJECT lblorder AS label WITH ;
		AutoSize = .F., ;
		FontSize = 8, ;
		Alignment = 2, ;
		Caption = "*Order =", ;
		Height = 16, ;
		Left = 1352, ;
		Top = 6, ;
		Width = 120, ;
		BackColor = RGB(255,255,128), ;
		Name = "lblOrder"


	ADD OBJECT edtnewnotes AS editbox WITH ;
		Anchor = 10, ;
		Height = 56, ;
		Left = 5, ;
		TabIndex = 14, ;
		Top = 310, ;
		Width = 1500, ;
		ControlSource = "webimp.mnewnotes", ;
		Name = "edtNewNotes"


	PROCEDURE dataassist
		PRIVATE m.lcNotes, m.lcNoteTxt, m.lcPDANum, m.ldEntryDate, m.lcEntryTime, m.llDoOrigDateToo, m.ldUpDate, m.lnIntDupe,	;
			m.lcSSNPolicy, m.lcLicPolicy, m.lnMultiName, m.lnMultiSSN, m.lnMultiLic, m.lnIntSSEI, m.lnIntLic, 	;
			m.lcUser, m.lmNotes, m.lcGapNote, m.lcSetDeleted, m.lcAltNoteXLS, m.lcAltNoteWeb

		IF FILE("Step.on")
			SET STEP ON 
		ENDIF

		m.lcSetDeleted	=SET("Deleted")
		SET DELETED ON

		m.ldEntryDate	=DATE()
		m.lcEntryTime	=TIME()
		*m.lmNotes		=DTOC(DATE())+" - premium transactions downloaded ";
						+"from online payment system."
		*m.lmXLSNotes	=DTOC(DATE())+" - premium transactions imported from state-provided spreadsheet."
		m.lmNotes		=":  Premium transactions downloaded from online payment system."
		m.lmXLSNotes	=":  Premium transactions imported from state-provided spreadsheet."

		m.lcAltNoteXLS	=":  Premium transactions uploaded from Multiple Licensee Enrollment Form provided by "
		m.lcAltNoteWeb	=":  "

		*m.lcUser	="EdS"
		m.lcUser	=ALLTRIM(ThisForm.txtUser.Value)

		*ON KEY LABEL F4 WAIT WINDOW NOWAIT "F4 is no longer available"
		m.lnSelect	=SELECT()

		PRIVATE m.lnRecno
		m.lnRecno	=RECNO()

		*SELECT 0
		*USE ..\ChartA
		*SET ORDER TO STATEYR   && STATE+LEFT(POLINUMB,2)
		*SET FILTER TO ChartA.lActive=.T.

		IF !USED("IntDupe")
			SELECT 0
			USE WebImp AGAIN ALIAS IntDupe
			SET FILTER TO !IntDupe.lHold .and. ALLTRIM(IntDupe.cUsername) = ALLTRIM(Thisform.txtuser.Value)
		ENDIF

		SELECT WebImp
		WAIT WINDOW NOWAIT LEFT("Updating..." +SPACE(40), 40)

		* Append 'x' to PDANUM if file PDANUMx.txt exists	&& done by CSVxWebimp concordance now
		*IF FILE("PDANUMx.txt")
		*	REPLACE pdanum WITH ALLTRIM(pdanum)+'x' FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM(Thisform.txtUser.Value)
		*ENDIF
		*---------------------------------------------------------------*

		WAIT WINDOW NOWAIT LEFT("Updating........." +SPACE(40), 40)

		SCAN FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM(Thisform.txtUser.Value)
			* Find Policy in Chart	------------------------------------------------------------------------------------*
			SELECT ChartA
			=SEEK(ALLTRIM(WebImp.GovState)) 

			DO CASE
			* EO policy
			CASE Webimp.cCovType = "RE"
				LOCATE REST WHILE ChartA.State=ALLTRIM(WebImp.GovState)	;
					FOR SUBSTR(ChartA.PoliNumb,4,2)="EO"	;
					.AND. BETWEEN(WebImp.Effective, ChartA.Effective, ChartA.End-1)	;
					.AND. ChartA.end = Webimp.end

			* MLO policy
			CASE Webimp.cCovType = "ML"
				LOCATE REST WHILE ChartA.State=ALLTRIM(WebImp.GovState)	;
					FOR SUBSTR(ChartA.PoliNumb,4,2)="ML" ;
					.AND. BETWEEN(WebImp.Effective, ChartA.Effective, ChartA.End-1)	;
					.AND. ChartA.end = Webimp.end

			* AP policy
			OTHERWISE
				LOCATE REST WHILE ChartA.State=ALLTRIM(WebImp.GovState)	;
					FOR SUBSTR(ChartA.PoliNumb,4,2)="AP"	;
					.AND. BETWEEN(WebImp.Effective, ChartA.Effective, ChartA.End-1)	;
					.AND. ChartA.end = Webimp.end

			ENDCASE


			* Populate these policy attributes
			IF FOUND()
				REPLACE WebImp.PoliNumb		WITH ChartA.PoliNumb,	;
						WebImp.AddTl_Key	WITH ChartA.AddTl_Key,	;
						WebImp.CompOwner	WITH ChartA.cCompOwner,	;
						WebImp.End			WITH ChartA.End 
			ELSE
				REPLACE WebImp.PDANum	WITH "GovState !in ChartA"
			ENDIF

			SELECT WebImp


		ENDSCAN


		*USE IN ChartA

		*IF EMPTY(TAG())
		*	USE WebImp EXCLUSIVE
		*	INDEX ON LastName+FirstName+PoliNumb	TAG NamePoli
		*	USE WebImp
		*ELSE
		*	SET ORDER TO NamePoli	&& LASTNAME+FIRSTNAME+POLINUMB
		*ENDIF

		SET ORDER TO 

		WAIT WINDOW NOWAIT LEFT("Updating............" +SPACE(40), 40)
		*REPLACE ALL	EntryBy		WITH 'www' +WebImp.BatchNum +'-' +RIGHT('00000'+ALLTRIM(STR(RECNO())),5)

		REPLACE ALL	EntryDate	WITH m.ldEntryDate, EntryTime	WITH m.lcEntryTime FOR !WebImp.lHold
		REPLACE ALL	EntryBy		WITH IIF(LEFT(WebImp.PDANum,3)="XLS", "XLS", 'www') +WebImp.BatchNum +'-' +RIGHT('00000'+ALLTRIM(STR(RECNO())),5)  FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM(Thisform.txtUser.Value)
		REPLACE ALL PDANum		WITH "" FOR !WebImp.lHold .AND. LEFT(WebImp.PDANum,3)="XLS" .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM(Thisform.txtUser.Value)
		*REPLACE ALL PremImport	WITH Prem FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM(Thisform.txtUser.Value)

		WAIT WINDOW NOWAIT LEFT("Updating..............." +SPACE(40), 40)
		m.lcNoteTxt	=ALLTRIM(NoteTxt.NoteTxt)
		m.lcDUPEnote="******* POTENTIAL DUPE PAYMENT, ENTER NOTES IN THE ORIGINAL POLICY *******"+CHR(13)

		SCAN FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM(Thisform.txtUser.Value)
			m.lcPDANum	=ALLTRIM(WebImp.PDANum) +": "
		*	m.lcSSNPolicy	=IIF(EMPTY(WebImp.SSN) .OR. LEFT(WebImp.ssn,5)="*****", " ", PI_SSN.PoliNumb)
			m.lcSSNPolicy	=IIF(EMPTY(WebImp.cSSN4+WebImp.cEIN) .OR. VAL(WebImp.cSSN4+WebImp.cEIN)=0, " ", PI_SSN.PoliNumb)
			PRIVATE m.lcLicNumber
			m.lcLicNumber	=Thisform.GetLicense(WebImp.cPolilic,1,"NUMBER")
			m.lcLicNumber	=IIF(UPPER(m.lcLicNumber)="PENDING","",m.lcLicNumber)
			m.lcLicPolicy	=IIF(EMPTY(m.lcLicNumber), " ", PI_Lic.PoliNumb)
			m.lcGapNote	=""

			* If the SSN match is the latest...
			IF (m.lcSSNPolicy >= m.lcLicPolicy)
				PRIVATE m.llGotSSNorEIN
				m.llGotSSNorEIN=!EMPTY(WebImp.cSSN4+WebImp.cEIN) .AND. VAL(WebImp.cSSN4+WebImp.cEIN)#0

				* PDANum "Status" ----------------------------------*
				DO CASE
				* Found previous by SSN and new effective equals prior end -> (Ok)
				* Found previous by SSN4 and new effective equals prior end -> (Ok)
		*		CASE !EMPTY(WebImp.SSN) .AND. FOUND("PI_SSN") .AND. WebImp.Effective=PI_SSN.End .AND. (m.lcSSNPolicy>=m.lcLicPolicy)
				CASE m.llGotSSNorEIN .AND. FOUND("PI_SSN") .AND. WebImp.Effective=PI_SSN.End .AND. (m.lcSSNPolicy>=m.lcLicPolicy)	;
					.AND. (IIF(WebImp.cPolicyTyp="Firm" .AND. PI_SSN.nPolicyTyp=2, .T., .F.) .OR. IIF(WebImp.cPolicyTyp="Individual" .AND. PI_SSN.nPolicyTyp=1, .T., .F.))
		*			REPLACE WebImp.PoliID	WITH PI_SSN.PoliID, WebImp.OrigDate WITH PI_SSN.OrigDate
					REPLACE WebImp.PoliID	WITH PI_SSN.PoliID, WebImp.OrigDate WITH {}

				* Found previous by SSN and new effective is between prior effective and end -> "DUPE"
		*		CASE !EMPTY(WebImp.SSN) .AND. FOUND("PI_SSN") .AND. (BETWEEN(WebImp.Effective, PI_SSN.Effective, PI_SSN.End-1) .OR. WebImp.End = PI_SSN.End) .AND. (m.lcSSNPolicy>=m.lcLicPolicy)
				CASE m.llGotSSNorEIN .AND. FOUND("PI_SSN") .AND. (BETWEEN(WebImp.Effective, PI_SSN.Effective, PI_SSN.End-1) .OR. WebImp.End = PI_SSN.End) .AND. (m.lcSSNPolicy>=m.lcLicPolicy)	;
					.AND. (IIF(WebImp.cPolicyTyp="Firm" .AND. PI_SSN.nPolicyTyp=2, .T., .F.) .OR. IIF(WebImp.cPolicyTyp="Individual" .AND. PI_SSN.nPolicyTyp=1, .T., .F.))
					m.lcPDANum	=ALLTRIM(m.lcPDANum) +"DUPE"
		*			REPLACE WebImp.PoliID	WITH SPACE(LEN(WebImp.PoliID))	&&	Duplicates must not inherit PoliID#s
		*			REPLACE WebImp.OrigDate	WITH {}							&&		or OrigDates
					REPLACE WebImp.PoliID	WITH PI_SSN.PoliID, WebImp.OrigDate WITH {}

				* Found previous by SSN and new effective is later than prior end -> "GAP"
		*		CASE !EMPTY(WebImp.SSN) .AND. FOUND("PI_SSN") .AND. WebImp.Effective>PI_SSN.End .AND. (m.lcSSNPolicy>=m.lcLicPolicy)
				CASE m.llGotSSNorEIN .AND. FOUND("PI_SSN") .AND. WebImp.Effective>PI_SSN.End .AND. (m.lcSSNPolicy>=m.lcLicPolicy)	;
					.AND. (IIF(WebImp.cPolicyTyp="Firm" .AND. PI_SSN.nPolicyTyp=2, .T., .F.) .OR. IIF(WebImp.cPolicyTyp="Individual" .AND. PI_SSN.nPolicyTyp=1, .T., .F.))
					m.lcPDANum	=ALLTRIM(m.lcPDANum) +"GAP"
					m.lcGapNote	="=======GAP from " +DTOC(PI_SSN.End) +"=======" +CHR(13)
		*			REPLACE WebImp.PoliID WITH PI_SSN.PoliID, WebImp.OrigDate	WITH WebImp.Effective
					REPLACE WebImp.PoliID WITH PI_SSN.PoliID, WebImp.OrigDate	WITH {}

				* no above conditions apply
				OTHERWISE
					&&	It is Ok.
		*			REPLACE WebImp.OrigDate	WITH WebImp.Effective
					REPLACE WebImp.OrigDate	WITH {}

				ENDCASE

			ELSE	&&	If the Lic# match is the latest...

				DO CASE
				* Found previous by License and new effective equals prior end -> (Ok)
		*		CASE !EMPTY(WebImp.cPoliLic) .AND. FOUND("PI_Lic") .AND. WebImp.Effective=PI_Lic.End .AND. (m.lcLicPolicy>m.lcSSNPolicy)
				CASE !EMPTY(m.lcLicNumber) .AND. FOUND("PI_Lic") .AND. WebImp.Effective=PI_Lic.End .AND. (m.lcLicPolicy>m.lcSSNPolicy)	;
					.AND. (IIF(WebImp.cPolicyTyp="Firm" .AND. PI_Lic.nPolicyTyp=2, .T., .F.) .OR. IIF(WebImp.cPolicyTyp="Individual" .AND. PI_Lic.nPolicyTyp=1, .T., .F.))
		*			REPLACE WebImp.PoliID WITH PI_Lic.PoliID, WebImp.OrigDate WITH PI_Lic.OrigDate
					REPLACE WebImp.PoliID WITH PI_Lic.PoliID, WebImp.OrigDate WITH {}

				* Found previous by License and new effective is between prior effective and end -> "DUPE"
		*		CASE !EMPTY(WebImp.cPoliLic) .AND. FOUND("PI_Lic") .AND. (BETWEEN(WebImp.Effective, PI_Lic.Effective, PI_Lic.End-1) .OR. WebImp.End = PI_Lic.End)  .AND. (m.lcLicPolicy>m.lcSSNPolicy)
				CASE !EMPTY(m.lcLicNumber) .AND. FOUND("PI_Lic") .AND. (BETWEEN(WebImp.Effective, PI_Lic.Effective, PI_Lic.End-1) .OR. WebImp.End = PI_Lic.End)  .AND. (m.lcLicPolicy>m.lcSSNPolicy)	;
					.AND. (IIF(WebImp.cPolicyTyp="Firm" .AND. PI_Lic.nPolicyTyp=2, .T., .F.) .OR. IIF(WebImp.cPolicyTyp="Individual" .AND. PI_Lic.nPolicyTyp=1, .T., .F.))
					m.lcPDANum	=ALLTRIM(m.lcPDANum) +"DUPE"
		*			REPLACE WebImp.PoliID	WITH SPACE(LEN(WebImp.PoliID))	&&	Duplicates must not inherit PoliID#s
		*			REPLACE WebImp.OrigDate	WITH {}							&&		or OrigDates
					REPLACE WebImp.PoliID WITH PI_Lic.PoliID, WebImp.OrigDate WITH {}

				* Found previous by License and new effective is later than prior end -> "GAP"
		*		CASE !EMPTY(WebImp.cPoliLic) .AND. FOUND("PI_Lic") .AND. WebImp.Effective>PI_Lic.End .AND. (m.lcLicPolicy>m.lcSSNPolicy)
				CASE !EMPTY(m.lcLicNumber) .AND. FOUND("PI_Lic") .AND. WebImp.Effective>PI_Lic.End .AND. (m.lcLicPolicy>m.lcSSNPolicy)	;
					.AND. (IIF(WebImp.cPolicyTyp="Firm" .AND. PI_Lic.nPolicyTyp=2, .T., .F.) .OR. IIF(WebImp.cPolicyTyp="Individual" .AND. PI_Lic.nPolicyTyp=1, .T., .F.))
					m.lcPDANum	=ALLTRIM(m.lcPDANum) +"GAP"
					m.lcGapNote	="=======GAP from " +DTOC(PI_Lic.End) +"=======" +CHR(13)
		*			REPLACE WebImp.PoliID WITH PI_Lic.PoliID, WebImp.OrigDate	WITH WebImp.Effective
					REPLACE WebImp.PoliID WITH PI_Lic.PoliID, WebImp.OrigDate	WITH {}

				* no above conditions apply
				OTHERWISE
					&&	It is Ok.
		*			REPLACE WebImp.OrigDate	WITH WebImp.Effective
					REPLACE WebImp.OrigDate	WITH {}

				ENDCASE

				* Found previous by License having SS# while new does not
				IF !EMPTY(PI_Lic.cSSN4) .AND. EMPTY(WebImp.cSSN4)
					REPLACE WebImp.cSSN4	WITH PI_Lic.cSSN4
				ENDIF

			ENDIF


			IF RIGHT(m.lcPDANum, 2)=": "
				m.lcPDANum	=LEFT(m.lcPDANum, LEN(m.lcPDANum)-2)
			ENDIF

			IF LEFT(m.lcPDANum, 1)=":"
				m.lcPDANum	=SUBSTR(m.lcPDANum, 2)
			ENDIF

			REPLACE WebImp.PDANum	WITH m.lcPDANum


			*-------------------------------------------------------------------------------------------------------*
			* Business Rules:  																						|
			*	CONFORMITY states are CO, KY, NE, RI, TN, LA, NM, SD, IA, MS, ND, ID, WY                            |
			*		-Policy Holder may-not/need-not purchase a Conformity for the state in which they reside.		|
			*		-Policy Holder may not purchase Conformity for a state in which they do not reside.				|
			*-------------------------------------------------------------------------------------------------------*
			IF !EMPTY(WebImp.cConfLic)
			DO CASE
		*	* If the resident state abbreviation (WebImp.State) is also the name of one of the Conformity fieldnames,
		*	*	.AND. that field contains a "1" for a purchased Conformity,
		*	*	indicate CONF in WebImp.PDANum...  
		*	CASE TYPE("EVALUATE(WebImp.State)")='C' .and. EVALUATE(WebImp.State)="1"
		*		REPLACE WebImp.PDANum	WITH ALLTRIM(WebImp.PDANum) +IIF(!EMPTY(WebImp.PDANum), ": ", "") +"CONF"
			* If the resident state addreviation (WebImp.State) in also specified in the Conformity endorsement (WebImp.cConfLic)
			CASE ALLTRIM(WebImp.State)$WebImp.cConfLic
				REPLACE WebImp.PDANum	WITH ALLTRIM(WebImp.PDANum) +IIF(!EMPTY(WebImp.PDANum), ": ", "") +"CONF"

		*	* If the Policy Holder purchased a Conformity 
		*	*	.AND. the Policy is not for the Policy Holder's state of residence...
		*	CASE (VAL(WebImp.CO +WebImp.KY +WebImp.NE +WebImp.RI +WebImp.TN +WebImp.LA +WebImp.NM +WebImp.SD +WebImp.IA +WebImp.MS +WebImp.ND +WebImp.ID +WebImp.WY)>0)	;
		*		.AND. (SUBSTR(WebImp.PoliNumb,11,2)#WebImp.State)
		*		REPLACE WebImp.PDANum	WITH ALLTRIM(WebImp.PDANum) +IIF(!EMPTY(WebImp.PDANum), ": ", "") +"CONF"
			CASE SUBSTR(WebImp.PoliNumb, 11, 2)#ALLTRIM(WebImp.State)
				REPLACE WebImp.PDANum	WITH ALLTRIM(WebImp.PDANum) +IIF(!EMPTY(WebImp.PDANum), ": ", "") +"CONF"

			ENDCASE
			ENDIF


			* Get the FULL Premium cost (i.e. NOT prorated)
			DO CASE
			* PremiumA
			CASE WebImp.PremiumA=.T.
				SELECT ChartC
				SET FILTER TO LEFT(endorse,9)="PREMIUM A"
			* PremiumB
			CASE WebImp.PremiumB=.T.
				SELECT ChartC
				SET FILTER TO LEFT(endorse,9)="PREMIUM B"
			* regular "PREMIUM "...
			OTHERWISE
				SELECT ChartC
				SET FILTER TO LEFT(endorse,7)="PREMIUM"
			ENDCASE

			SELECT WebImp
			m.lnPremiumAmt	=IIF(SEEK(WebImp.PoliNumb, "ChartC"), IIF(WebImp.PremiumB=.T., ThisForm.GetPremiumB(), ChartC.Amount), 0.00)

			GO (RECNO())	&&	reset the relation into ChartC

			IF m.lnPremiumAmt =0
				WAIT WINDOW "WebImp.PoliNumb "+ALLTRIM(WebImp.PoliNumb) +" PREMIUM not found in ChartC!"
				LOOP
			ENDIF

			* Notes --------------------------------------------*
			PRIVATE m.llGotSSNorEIN, m.lnParaAmt, m.lcParaWords, m.lcOldNotes
			m.lcOldNotes=""
			m.llGotSSNorEIN	=!EMPTY(WebImp.cSSN4+WebImp.cEIN).and.VAL(WebImp.cSSN4+WebImp.cEIN)#0

			m.lnParaAmt	=WebImp.nTxCtAmt+WebImp.nTxCoAmt+WebImp.Surcharge

				DO CASE
				* Tax/Surcharge
				CASE (WebImp.nTxCtAmt+WebImp.nTxCoAmt) > 0 .AND. WebImp.Surcharge > 0
					m.lcParaWords	=" and $"+ALLTRIM(STR(m.lnParaAmt,7,2)) +" tax/surcharge"
				* Tax
				CASE (WebImp.nTxCtAmt+WebImp.nTxCoAmt) > 0 .AND. WebImp.Surcharge = 0
					m.lcParaWords	=" and $"+ALLTRIM(STR(m.lnParaAmt,7,2)) +" tax"
				* Surcharge
				CASE (WebImp.nTxCtAmt+WebImp.nTxCoAmt) = 0 .AND. WebImp.Surcharge > 0
					m.lcParaWords	=" and $"+ALLTRIM(STR(m.lnParaAmt,7,2)) +" surcharge"
				OTHERWISE
					m.lcParaWords	=""
				ENDCASE


			DO CASE
			* Duplicate?
		*	CASE "DUPE"$WebImp.PDANum
		*		IF EMPTY(m.lcNoteTxt)
		*			m.lcDUPEnote="******* POTENTIAL DUPE PAYMENT, ENTER NOTES IN THE ORIGINAL POLICY *******"+CHR(13)
		*			m.lcNotes	=m.lcDUPEnote+ALLTRIM(IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lmXLSNotes, m.lmNotes))	;
								+ "  $"+ALLTRIM(STR(m.lnPremiumAmt))	;
								+ " premium paid" +IIF(WebImp.OnlineChg=.T., " plus $5.00 Online Charge", "") +m.lcParaWords +". (" +ALLTRIM(m.lcUser) +", PDA#" +ALLTRIM(WebImp.PDANum) +")"
		*		ELSE
		*			m.lcNotes	=m.lcNoteTxt
		*		ENDIF

		*		m.lcOldNotes	=""

		**		REPLACE PoliID 	WITH SPACE(LEN(PoliID))		&& redundant assurance that DUPE inherits no PoliID

			* Premium 
			CASE (m.lnPremiumAmt >0) .AND. FOUND("ChartC") .AND. (WebImp.PremImport-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt = m.lnPremiumAmt) .AND. (m.lnPremiumAmt = ChartC.Amount)
				IF EMPTY(m.lcNoteTxt)
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+ALLTRIM(IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lmXLSNotes, m.lmNotes))	;
								+ "  $"+ALLTRIM(STR(m.lnPremiumAmt))	;
								+ " premium paid" +IIF(WebImp.OnlineChg=.T., " plus $5.00 Online Charge", "") +m.lcParaWords +". (" +ALLTRIM(m.lcUser) +", PDA#" +ALLTRIM(WebImp.PDANum) +")"
				ELSE
		*			m.lcNotes	=ALLTRIM(IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lmXLSNotes, m.lmNotes))+" "+ALLTRIM(WebImp.Polinumb)+":  "+m.lcNoteTxt
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lcAltNoteXLS, m.lcAltNoteWeb)+m.lcNoteTxt
				ENDIF

				DO CASE
				* If SS# Policy is later than (or equal to) the Lic# Policy
				CASE m.llGotSSNorEIN .AND. m.lcSSNPolicy>=m.lcLicPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote	&& +ALLTRIM(PI_SSN.Notes) 
					m.lcOldNotes=ALLTRIM(PI_SSN.Notes)

				* If Lic# Policy is later than the SS# Policy
				CASE !EMPTY(WebImp.cPoliLic) .AND. m.lcLicPolicy>m.lcSSNPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote 	&& +ALLTRIM(PI_Lic.Notes)
					m.lcOldNotes=ALLTRIM(PI_Lic.Notes)

				ENDCASE


			* Premium + Endorsement(s)
			CASE (m.lnPremiumAmt >0) .AND. FOUND("ChartC") .AND. (WebImp.PremImport-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt > m.lnPremiumAmt) .AND. (m.lnPremiumAmt = ChartC.Amount)
				IF EMPTY(m.lcNoteTxt)
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+ALLTRIM(IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lmXLSNotes, m.lmNotes))	;
								+" $" +ALLTRIM(STR(m.lnPremiumAmt)) +" total premium paid.  $" +ALLTRIM(STR(WebImp.PremImport-m.lnPremiumAmt-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt)) +" toward endorsements" +IIF(WebImp.OnlineChg=.T., " plus $5.00 Online Charge", "") +m.lcParaWords +"." +" ("+ALLTRIM(m.lcUser) +", PDA#" +ALLTRIM(WebImp.PDANum) +")"
				ELSE
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lcAltNoteXLS, m.lcAltNoteWeb)+m.lcNoteTxt
				ENDIF

				DO CASE
				* If SS# Policy is later than (or equal to) the Lic# Policy
				CASE m.llGotSSNorEIN .AND. m.lcSSNPolicy>=m.lcLicPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote	&& +ALLTRIM(PI_SSN.Notes)
					m.lcOldNotes=ALLTRIM(PI_SSN.Notes)

				* If Lic# Policy is later than the SS# Policy
				CASE !EMPTY(WebImp.cPoliLic) .AND. m.lcLicPolicy>m.lcSSNPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote 	&& +ALLTRIM(PI_Lic.Notes)
					m.lcOldNotes=ALLTRIM(PI_Lic.Notes)

				ENDCASE

			* Prorated Premium 
			CASE FOUND("ChartC") .AND. (WebImp.PremImport-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt = ChartC.Amount)
				IF EMPTY(m.lcNoteTxt)
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+ALLTRIM(IIF(LEFT(WebImp.EntryBy, 3)="XLS", m.lmXLSNotes, m.lmNotes))	;
								+ "  $"+ALLTRIM(STR(WebImp.PremImport-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt,11,2))	;
								+ " prorated premium paid" +IIF(WebImp.OnlineChg=.T., " plus $5.00 Online Charge", "") +m.lcParaWords +". (" +ALLTRIM(m.lcUser) +", PDA#" +ALLTRIM(WebImp.PDANum) +")"
				ELSE
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lcAltNoteXLS, m.lcAltNoteWeb)+m.lcNoteTxt
				ENDIF

				DO CASE
				* If SS# Policy is later than (or equal to) the Lic# Policy
				CASE m.llGotSSNorEIN .AND. m.lcSSNPolicy>=m.lcLicPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote	&& +ALLTRIM(PI_SSN.Notes)
					m.lcOldNotes=ALLTRIM(PI_SSN.Notes)

				* If Lic# Policy is later than the SS# Policy
				CASE !EMPTY(WebImp.cPoliLic) .AND. m.lcLicPolicy>m.lcSSNPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote 	&& +ALLTRIM(PI_Lic.Notes)
					m.lcOldNotes=ALLTRIM(PI_Lic.Notes)

				ENDCASE

			* Prorate Premium + Endorsement(s)
			CASE FOUND("ChartC") .AND. (WebImp.PremImport-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt > ChartC.Amount)
				IF EMPTY(m.lcNoteTxt)
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+ALLTRIM(IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lmXLSNotes, m.lmNotes))	;
								+" $" +ALLTRIM(STR(ChartC.Amount)) +" prorated premium paid.  $" +ALLTRIM(STR(WebImp.PremImport-ChartC.Amount-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt)) +" toward endorsements" +IIF(WebImp.OnlineChg=.T., " plus $5.00 Online Charge", "") +m.lcParaWords+"." +" ("+ALLTRIM(m.lcUser) +", PDA#" +ALLTRIM(WebImp.PDANum) +")"
				ELSE
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lcAltNoteXLS, m.lcAltNoteWeb)+m.lcNoteTxt
				ENDIF

				DO CASE
				* If SS# Policy is later than (or equal to) the Lic# Policy
				CASE m.llGotSSNorEIN .AND. m.lcSSNPolicy>=m.lcLicPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote 	&& +ALLTRIM(PI_SSN.Notes)
					m.lcOldNotes=ALLTRIM(PI_SSN.Notes)

				* If Lic# Policy is later than the SS# Policy
				CASE !EMPTY(WebImp.cPoliLic) .AND. m.lcLicPolicy>m.lcSSNPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote 	&& +ALLTRIM(PI_Lic.Notes)
					m.lcOldNotes=ALLTRIM(PI_Lic.Notes)

				ENDCASE

			* Prorated Premium + Endorsements
			CASE FOUND("ChartC") .AND. (WebImp.PremImport-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt < ChartC.Amount)
				IF EMPTY(m.lcNoteTxt)
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+ALLTRIM(IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lmXLSNotes, m.lmNotes))	;
								+"  $"+ALLTRIM(STR(WebImp.PremImport-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt,11,2))		;
								+" toward"	;
								+" total premium paid" +IIF(WebImp.OnlineChg=.T., " plus $5.00 Online Charge", "") +m.lcParaWords +"."	;
								+"  $"+ALLTRIM(STR(ChartC.Amount-WebImp.PremImport-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt)) + " short? (" +ALLTRIM(m.lcUser) +", PDA#" +ALLTRIM(WebImp.PDANum) +")"
				ELSE
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lcAltNoteXLS, m.lcAltNoteWeb)+m.lcNoteTxt
				ENDIF

				DO CASE
				* If SS# Policy is later than (or equal to) the Lic# Policy
				CASE m.llGotSSNorEIN .AND. m.lcSSNPolicy>=m.lcLicPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote 	&& +ALLTRIM(PI_SSN.Notes)
					m.lcOldNotes=ALLTRIM(PI_SSN.Notes)

				* If Lic# Policy is later than the SS# Policy
				CASE !EMPTY(WebImp.cPoliLic) .AND. m.lcLicPolicy>m.lcSSNPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote 	&& +ALLTRIM(PI_Lic.Notes)
					m.lcOldNotes=ALLTRIM(PI_Lic.Notes)

				ENDCASE

				m.lcPDANum	=ALLTRIM(m.lcPDANum) +IIF(!EMPTY(m.lcPDANum), ",", "") +"?"

			OTHERWISE
				IF EMPTY(m.lcNoteTxt)
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+ALLTRIM(IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lmXLSNotes, m.lmNotes))	;
								+"  $"+ALLTRIM(STR(WebImp.PremImport-IIF(WebImp.OnlineChg=.T., 5, 0)-m.lnParaAmt,11,2)) +" total premium paid" +IIF(WebImp.OnlineChg=.T., " plus $5.00 Online Charge", "") +m.lcParaWords+".  (" +ALLTRIM(m.lcUser) +", PDA#" +ALLTRIM(WebImp.PDANum) +")"
				ELSE
					m.lcNotes	=DTOC(DATE())+"  -  "+ALLTRIM(WebImp.Polinumb)+IIF(LEFT(WebImp.EntryBy,3)="XLS", m.lcAltNoteXLS, m.lcAltNoteWeb)+m.lcNoteTxt
				ENDIF

				DO CASE
				* If SS# Policy is later than (or equal to) the Lic# Policy
				CASE m.llGotSSNorEIN .AND. m.lcSSNPolicy>=m.lcLicPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote 	&& +ALLTRIM(PI_SSN.Notes)
					m.lcOldNotes=ALLTRIM(PI_SSN.Notes)

				* If Lic# Policy is later than the SS# Policy
				CASE !EMPTY(WebImp.cPoliLic) .AND. m.lcLicPolicy>m.lcSSNPolicy
					m.lcNotes	=m.lcNotes +CHR(13)+CHR(13) +m.lcGapNote 	&& +ALLTRIM(PI_Lic.Notes)
					m.lcOldNotes=ALLTRIM(PI_Lic.Notes)

				ENDCASE

			ENDCASE

			IF !("DUPE"$WebImp.PDANum)
				REPLACE WebImp.Notes		WITH m.lcOldNotes
			ENDIF

			IF "DUPE"$WebImp.PDAnum
				m.lcNotes	=m.lcDUPEnote+m.lcNotes
			ENDIF

			IF EMPTY(WebImp.mNewNotes)
				REPLACE WebImp.mNewNotes	WITH ALLTRIM(m.lcNotes)
			ELSE
				REPLACE WebImp.mNewNotes	WITH LEFT(m.lcNotes, LEN(m.lcNotes)-1) +IIF(EMPTY(WebImp.mNewNotes),"", ALLTRIM(WebImp.mNewNotes))
			ENDIF

			SELECT IntDupe
			GO TOP
			COUNT TO m.lnIntDupe FOR ALLTRIM(       PoliNumb)+ALLTRIM(       FirstName)+ALLTRIM(       LastName)+ALLTRIM(       Company)	;
								   = ALLTRIM(WebImp.PoliNumb)+ALLTRIM(WebImp.FirstName)+ALLTRIM(WebImp.LastName)+ALLTRIM(WebImp.Company)

			IF !EMPTY(WebImp.cSSN4 +WebImp.cEIN)
		*		COUNT TO m.lnIntSSN FOR ALLTRIM(PoliNumb) +ALLTRIM(cSSN4) = ALLTRIM(WebImp.PoliNumb) +ALLTRIM(WebImp.SSN)	;
									.AND. RECNO() # RECNO("WebImp")

		*		COUNT TO m.lnIntSSEI FOR ALLTRIM(      CSSN4       +CEIN)+ALLTRIM(LEFT(       FIRSTNAME,2)+LEFT(       LASTNAME,3))+SUBSTR(       POLINUMB,11,2)	;
									  = ALLTRIM(WebImp.CSSN4+WebImp.CEIN)+ALLTRIM(LEFT(WebImp.FIRSTNAME,2)+LEFT(WebImp.LASTNAME,3))+SUBSTR(WebImp.POLINUMB,11,2)
				COUNT TO m.lnIntSSEI FOR ALLTRIM(      CSSN4       +CEIN)+ALLTRIM(LEFT(       FIRSTNAME,2)+LEFT(       LASTNAME,3))+POLINUMB	;
									  = ALLTRIM(WebImp.CSSN4+WebImp.CEIN)+ALLTRIM(LEFT(WebImp.FIRSTNAME,2)+LEFT(WebImp.LASTNAME,3))+WebImp.POLINUMB
			ELSE
				m.lnIntSSEI	=0
			ENDIF

		*	IF !EMPTY(WebImp.cPoliLic)
		*		COUNT TO m.lnIntLic FOR ALLTRIM(       PoliNumb) +ALLTRIM(       cPoliLic)	;
		*							 == ALLTRIM(WebImp.PoliNumb) +ALLTRIM(WebImp.cPoliLic)	;
		*							.AND. RECNO() # RECNO("WebImp")
		*	ELSE
				m.lnIntLic	=0
		*	ENDIF

			COUNT TO m.lnMultiName FOR (ALLTRIM(PoliNumb)#ALLTRIM(WebImp.PoliNumb))	;
								.AND. ALLTRIM(       FirstName) +ALLTRIM(       LastName) +ALLTRIM(       Company)	;
									= ALLTRIM(WebImp.FirstName) +ALLTRIM(WebImp.LastName) +ALLTRIM(WebImp.Company)

			COUNT TO m.lnMultiSSN FOR (ALLTRIM(PoliNumb)#ALLTRIM(WebImp.PoliNumb))	;
								.AND. IIF(!EMPTY(WebImp.cSSN4), ALLTRIM(cSSN4) = ALLTRIM(WebImp.cSSN4), .F.)

		*	COUNT TO m.lnMultiLic FOR (ALLTRIM(PoliNumb)#ALLTRIM(WebImp.PoliNumb))	;
								.AND. IIF(!EMPTY(WebImp.cPoliLic), ALLTRIM(cPoliLic) = ALLTRIM(WebImp.cPoliLic), .F.)
			m.lnMultiLic=0	&& can no longer determine multiple licenses as previously done above.  Will have to invent a new way.

			SELECT WebImp

			IF (m.lnIntDupe >1	;
				.OR. m.lnIntSSEI >1	;
				.OR. m.lnIntLic >1)
				ThisForm.lFindIntPoliID	=.T.	&&	Will trigger call to IntPoliID() in procedure "PopIT"
				PRIVATE m.lcIntDUPEStr
				m.lcIntDUPEStr	="*INT DUPE:"

				IF m.lnIntDupe>1
					m.lcIntDUPEStr	=m.lcIntDUPEStr +" NAME"
				ENDIF

				IF m.lnIntSSEI>1
					m.lcIntDUPEStr	=m.lcIntDUPEStr +" SS#"
				ENDIF

				IF m.lnIntLic>1
					m.lcIntDUPEStr	=m.lcIntDUPEStr +" LIC#"
				ENDIF

				m.lcIntDUPEStr	=m.lcIntDUPEStr +"*"

				IF LEN(ALLTRIM(WebImp.PDANum) +m.lcIntDUPEStr) > LEN(WebImp.PDANum)
					REPLACE WebImp.PDANum		WITH ALLTRIM(WebImp.PDANum)+"*SEE NOTES*"
		*			REPLACE WebImp.Notes 		WITH m.lcIntDUPEStr + CHR(13) + ALLTRIM(WebImp.Notes)
					REPLACE WebImp.mNewNotes 	WITH m.lcIntDUPEStr + CHR(13) + ALLTRIM(WebImp.mNewNotes)

				ELSE
					REPLACE WebImp.PDANum		WITH ALLTRIM(WebImp.PDANum) +m.lcIntDUPEStr
				ENDIF

			ENDIF

			IF (m.lnMultiName + m.lnMultiSSN +m.lnMultiLic) >0
		*	IF (m.lnMultiName>1) .or. (m.lnMultiSSN>1) .or. (m.lnMultiLic>1) 
				ThisForm.lMultiPoliID	=.T.
				REPLACE WebImp.PDANum		WITH ALLTRIM(WebImp.PDANum)+"*SEE NOTES*"
			ENDIF

		ENDSCAN

		SELECT ChartC
		SET FILTER TO 
		SELECT WebImp

		WAIT WINDOW NOWAIT LEFT("Updating.................." +SPACE(40), 40)

		* NOTE:  2016/03/30
		*	Decision per conversation with LindaR, decided to make all new OrigDate as __/__/____ due to uncertainty
		*	about the quality of the pre-existing data that any attempted rules would act on.
		*REPLACE ALL OrigDate	WITH WebImp.Effective	FOR !WebImp.lHold .AND. EMPTY(WebImp.OrigDate) 
		REPLACE ALL OrigDate	WITH {}	FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM(Thisform.txtUser.Value)

		WAIT WINDOW NOWAIT LEFT("Updating....................." +SPACE(40), 40)
		REPLACE ALL	Prem		WITH ThisForm.GetPrem() FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM(Thisform.txtUser.Value)

		*WAIT WINDOW NOWAIT LEFT("Updating........................" +SPACE(40), 40)
		*REPLACE ALL	Address1	WITH UPPER(Address1) FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM((Thisform.txtUser.Value)
		*REPLACE ALL	Address2	WITH UPPER(Address2) FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM((Thisform.txtUser.Value)

		*WAIT WINDOW NOWAIT LEFT("Updating........................." +SPACE(40), 40)
		*REPLACE ALL	City		WITH UPPER(City) FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM((Thisform.txtUser.Value)
		*REPLACE ALL	County		WITH UPPER(County) FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM((Thisform.txtUser.Value)

		*WAIT WINDOW NOWAIT LEFT("Updating.........................." +SPACE(40), 40)
		*REPLACE ALL	State		WITH UPPER(State) FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = (Thisform.txtUser.Value)
		*REPLACE ALL	Tipe		WITH UPPER(Tipe) FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = (Thisform.txtUser.Value)

		*WAIT WINDOW NOWAIT LEFT("Updating..........................." +SPACE(40), 40)
		*REPLACE	ALL	Phone1		WITH ThisForm.FmtPhone(Phone1) FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = (Thisform.txtUser.Value)
		*REPLACE	ALL	Phone2		WITH ThisForm.FmtPhone(Phone2) FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = (Thisform.txtUser.Value)

		*WAIT WINDOW NOWAIT LEFT("Updating............................" +SPACE(40), 40)
		*REPLACE	ALL	Fax			WITH ThisForm.FmtPhone(Fax) FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = (Thisform.txtUser.Value)

		WAIT WINDOW NOWAIT LEFT("Updating............................." +SPACE(40), 40)
		*REPLACE ALL Taxes		WITH 0	FOR !WebImp.lHold .AND. (Taxes=0)  .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM(Thisform.txtUser.Value)
		REPLACE ALL Taxes		WITH (nTxCtAmt +nTxCoAmt)	FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = ALLTRIM(Thisform.txtUser.Value)
		*REPLACE ALL eMail		WITH ALLTRIM(LOWER(eMail)) FOR !WebImp.lHold .AND. ALLTRIM(Webimp.cUsername) = (Thisform.txtUser.Value)

		*SELECT PI_SSN
		*ON KEY LABEL CTRL+S	DO SelectSSN
		*BROWSE LAST	NOWAIT TITLE "SS#:  [Ctrl+S]=Browse" NOMODIFY
		*BROWSE NOWAIT TITLE "SS#:  [Ctrl+S]=Browse" NOMODIFY

		*SELECT PI_Lic
		*ON KEY LABEL CTRL+L	DO SelectLic
		*BROWSE LAST NOWAIT TITLE "License#:  [Ctrl+L]=Browse"	NOMODIFY
		*BROWSE NOWAIT TITLE "License#:  [Ctrl+L]=Browse"	NOMODIFY

		*SELECT (m.lnSelect)
		GO (m.lnRecno)
		*m.lcTitle	=m.lcMsgEdit +m.lcMsgEsc
		WAIT CLEAR
		*USE IN IntDupe

		SELECT WebImp
		SET ORDER TO 

		IF m.lcSetDeleted="OFF"
			SET DELETED OFF
		ENDIF

		GO TOP
		RETURN .T.
	ENDPROC


	PROCEDURE webimperror
		PARAMETERS m.lnERROR, m.lcMESSAGE, m.lnLINENO, m.lcPROGRAM
		*SET STEP ON 
		DO CASE
		* 108="File is in use by another user."
		* 109="Record is in use by another user."
		CASE m.lnError=108 .OR. m.lnError=109
			WAIT WINDOW TIMEOUT 5 "Attempting to add "+ALLTRIM(m.FirstName) +" "+ALLTRIM(m.LastName)
			RETRY

		OTHERWISE
			WAIT WINDOW ALLTRIM(STR(m.lnERROR)) +": " +m.lcMESSAGE +".  Line:"+ALLTRIM(STR(m.lnLINENO)) +", " +m.lcPROGRAM
		ENDCASE
	ENDPROC


	PROCEDURE locatedate
		* LocateDate(ChartB.Polinumb +ChartB.Endorse, m.Effective)->.L.
		PARAMETERS m.xcKey, m.ldEffective, m.llLocateDate
		PRIVATE m.lnSelect
		m.lnSelect	=SELECT()

		SELECT ChartC
		m.llLocateDate	=SEEK(m.xcKey)

		IF m.llLocateDate
			LOCATE REST WHILE PoliNumb +Endorse =m.xcKey	;
				FOR BETWEEN(m.ldEffective,ChartC.Effective, CTOD(LEFT(DTOC(ChartC.Effective),6) +ALLTRIM(STR(YEAR(ChartC.Effective)+1))))
			m.llLocateDate	=FOUND()
		ENDIF

		SELECT (m.lnSelect)
		RETURN m.llLocateDate
	ENDPROC


	PROCEDURE getpremiumamt
		PARAMETERS m.jcStateYearMo		&&	(ALLTRIM(GovState)+STR(YEAR(effective))+STR(MONTH(PolEff.Effective)))
		RETURN IIF(SEEK(m.jcStateYear, "PremAmt"), PremAmt.Premium, 0)
	ENDPROC


	PROCEDURE govstate
		PARAMETERS m.xcState
		m.xcState	=ALLTRIM(UPPER(m.xcState))

		DO CASE
		CASE m.xcState ='IDAHO'
			m.xcState	='ID'

		CASE m.xcState	='LOUISIANA'
			m.xcState	='LA'

		CASE m.xcState	='RHODE ISLAND' .OR. m.xcState	="RHODEISLAND"
			m.xcState	='RI'

		CASE m.xcState	='IDAHO'
			m.xcState	='ID'

		CASE m.xcState	='NORTH DAKOTA' .OR. m.xcState	='NORTH DAKOTA'
			m.xcState	='ND'

		CASE m.xcState	='SOUTH DAKOTA' .or. m.xcState	='SOUTH DAKOTA'
			m.xcState	='SD'

		CASE m.xcState ='KENTUCKY'
			m.xcState	='KY'

		CASE m.xcState ='MISSISSIPPI'
			m.xcState	='MS' 

		CASE m.xcState ='MONTANA'
			m.xcState	='MT'

		CASE m.xcState ='NEBRASKA'
			m.xcState ='NE'

		CASE m.xcState ='TENNESSEE'
			m.xcState ='TN'

		CASE m.xcState ='ALASKA'
			m.xcState ='AK'

		CASE m.xcState ='ALABAMA'
			m.xcState ='AL'

		CASE m.xcState ='COLORADO'
			m.xcState ='CO'

		CASE m.xcState ='IOWA'
			m.xcState ='IA'

		CASE m.xcState ='NORTH CAROLINA'
			m.xcState ='NC'

		CASE m.xcState ='UTAH'
			m.xcState ='UT'

		CASE m.xcState ='NEWMEXICO' .OR. m.xcState ='NEW MEXICO'
			m.xcState ='NM'

		CASE m.xcState ='WYOMING'
			m.xcState ='WY'

		OTHERWISE
			m.xcState =m.xcState

		ENDCASE

		RETURN m.xcState
	ENDPROC


	PROCEDURE getpremiumb
		PRIVATE m.lnSelect, m.lnAmount
		m.lnSelect	=SELECT()

		SELECT ChartC
		LOCATE REST WHILE ChartC.PoliNumb=WebImp.PoliNumb FOR LEFT(ChartC.Endorse, 9)="PREMIUM B"
		m.lnAmount	=ChartC.Amount

		SELECT (m.lnSelect)
		RETURN m.lnAmount
	ENDPROC


	PROCEDURE fmtssn
		PARAMETERS m.xcSSN
		PRIVATE m.lcReturn, m.lnCnt
		m.xcSSN		=ALLTRIM(m.xcSSN)
		m.lcReturn	=""

		FOR m.lnCnt =1 TO LEN(ALLTRIM(m.xcSSN))
			m.lcReturn	=(m.lcReturn +IIF(ISDIGIT(SUBSTR(m.xcSSN, m.lnCnt, 1)) .OR. SUBSTR(m.xcSSN, m.lnCnt, 1)="*", SUBSTR(m.xcSSN, m.lnCnt, 1), ""))
		NEXT

		RETURN m.lcReturn
	ENDPROC


	PROCEDURE fmtphone
		PARAMETERS m.jcPhone
		PRIVATE m.lcReturn, m.lnCnt
		m.jcPhone	=ALLTRIM(m.jcPhone)
		m.lcReturn	=""

		FOR m.lnCnt = 1 TO LEN(TRIM(m.jcPhone))
			m.lcReturn	=(m.lcReturn +IIF(ISDIGIT(SUBSTR(m.jcPhone, m.lnCnt, 1)), SUBSTR(m.jcPhone, m.lnCnt, 1), ""))
		NEXT

		IF LEN(m.lcReturn)=10
			m.lcReturn	="(" +LEFT(m.lcReturn, 3) +") " +SUBSTR(m.lcReturn, 4, 3) +"-" +SUBSTR(m.lcReturn, 7, 4)
		ENDIF

		RETURN m.lcReturn
	ENDPROC


	PROCEDURE import
		PRIVATE m.lcNameDisplay, m.lnSelect, m.lcCOrder, m.lcSetDeleted
		m.lnSelect	=SELECT()
		m.lcSetDeleted	=SET("Deleted")
		*SET PROCEDURE TO KeyGen
		SET DELETED ON

		*	Open ChartA, ChartB (w/ChartB2) & ChartC with relations
		*DO Chart IN IIF(RIGHT(CURDIR(),4)="WEB\", "..\", "") +"PROGRAMS\LIBRARY" WITH "OPEN"

		SELECT ChartC
		m.lcCOrder	=ORDER()
		SET ORDER TO PoliEnd	&& PoliNumb+Endorse+DTOS(Effective)

		SELECT ChartB
		SET ORDER TO POLIEND   && POLINUMB+ENDORSE
		*

		SELECT (m.lnSelect)

		WAIT CLEAR

		ThisForm.Errornum	=0
		ThisForm.ErrorTxt 	=""

		SELECT Web
		GO TOP
		LOCATE FOR !Web.Processed .AND. !Web.Reject .AND. ALLTRIM(Web.cUsername) = ALLTRIM(ThisForm.txtUser.Value)

		IF FOUND()
			SCATTER MEMVAR MEMO
		ELSE
			WAIT WINDOW NOWAIT "No records to import."
		ENDIF

		DO WHILE !EOF()	&&	Web.dbf

			DO WHILE !EOF()

				IF EMPTY(m.FirstName)
					m.lcNameDisplay	=ALLTRIM(m.LastName)
				ELSE
					m.lcNameDisplay	=ALLTRIM(m.FirstName) +" " +ALLTRIM(m.LastName)
				ENDIF

				IF !EMPTY(m.lcNameDisplay)
					m.lnTotalWidth	=100
					m.lnNameWidth	=LEN(m.lcNameDisplay)
					m.lnNetWidth	=(m.lnTotalWidth -m.lnNameWidth)
					m.lnPreWidth	=INT(m.lnNetWidth/2)
					m.lnPostWidth	=IIF(INT(m.lnNetWidth/2) =(m.lnNetWidth/2), m.lnPreWidth, m.lnPreWidth+1)
					WAIT WINDOW NOWAIT SPACE(m.lnPreWidth) +m.lcNameDisplay +SPACE(m.lnPostWidth)
				ENDIF

				DO CASE
				* Has this record already by Processed?
				CASE Web.Processed

				* Do we find PDA already in the Front-End (PoliInd.DBF)?
				CASE LEFT(Web.EntryBy, 3) #"XLS" .AND. ThisForm.fPDAnum(Web.PDANum)
					ThisForm.ErrorNum	=40
					ThisForm.ErrorTxt	="PDANUM already in poliind"
					ThisForm.Reject 	=.T.
					ThisForm.Processed =.T.
					GATHER MEMVAR MEMO

					IF SEEK(Web.PoliNumb +Web.PoliID,"WebImp","PoliNumb")
						RECALL IN WebImp
						REPLACE WebImp.lHold WITH .T.
					ENDIF

				* It's gotta be good.
				OTHERWISE
					ThisForm.ErrorNum	=0
					ThisForm.ErrorTxt	=""
					ThisForm.Reject	=.F.
					EXIT

				ENDCASE

				SKIP 1
				SCATTER MEMVAR MEMO
			ENDDO

			IF EOF()
				EXIT
			ENDIF

			SCATTER MEMVAR MEMO		&& Web.dbf
			m.Reject	=.F.

			* Now for endorsements
			*	First:  Get the # of endorsements per import data.
			m.EndorCnt	=IIF(m.HigherLim1,1,0)	;
						+IIF(m.HigherLim2,1,0) 	;
						+IIF(m.Appraisal,1,0)	;
						+IIF(m.PropMgr,1,0)		;
						+IIF(m.Environx,1,0)	;
						+IIF(m.FairHouse,1,0)	;
						+IIF(m.Complaints,1,0)	;
						+IIF(m.OnlineChg,1,0)	;
						+IIF(m.EarnMo,1,0)		;
						+IIF(m.LCEnviro,1,0)	;
						+IIF(m.FairHausA,1,0)	;
						+IIF(m.FairHausB,1,0)	;
						+IIF(m.PrimeRes,1,0)	;
						+IIF(m.HiLimMil1,1,0)	;
						+IIF(m.HiLimMil2,1,0)	;
						+IIF(m.ResPersInt,1,0)	;
						+IIF(m.Bndl,1,0)		;
						+IIF(m.BodyInj,1,0)		;
						+OCCURS(";",m.cApprTrain);
						+IIF(!EMPTY(m.cConfLic),1, 0);
						+IIF(m.Surcharge#0,1,0)	;
						+IIF(m.DevCon,1,0)		;
						+IIF(m.Subpoena,1,0)	;
						+IIF(m.RM10,1,0)		;
						+IIF(m.RM20,1,0)		;
						+IIF(m.LockBox,1,0)
		*				+IIF(m.CO .OR. m.KY .OR. m.NE .OR. m.RI .OR. m.TN .OR. m.LA .OR. m.NM .OR. m.SD .OR. m.IA .OR. m.MS .OR. m.ND .OR. m.ID .OR. m.WY, 1, 0)	;

			*	Second: Get the # of endorsements per RISC freebies (ChartC.DBF)
			*			and add it to the # of endorsements per import data.
			SELECT ChartC
			m.lnImportCount	=m.EndorCnt

			IF SEEK(m.PoliNumb)
				COUNT TO m.lnFreebieCount WHILE ChartC.PoliNumb =m.PoliNumb	;
					FOR	ChartC.Amount=0	;
					.AND. (EMPTY(ChartC.Tipe) .OR. ALLTRIM(Web.Tipe)$ChartC.Tipe)	;
					.AND. !("X"$ChartC.Tipe)	;
					.AND. SEEK(ChartC.PoliNumb +ChartC.Endorse,"ChartB","PoliEnd")	;
					.AND. ChartB.lBndl=.F.
				m.EndorCnt	=(m.EndorCnt +m.lnFreebieCount)
			ELSE
				m.lnFreebieCount	=0
			ENDIF

			m.lnTTLCount	=m.EndorCnt

			*	C): Prepare an array for subsequent populating with endorsements.
			SELECT Web

			DIMENSION ThisForm.laEndorse[m.lnTTLCount +1, 5]
			ThisForm.EndorTTL = 0

			*	Fouth: Populate array with endorsements per import data.
			ThisForm.GetEndor2()

			*	Filth: Populate with endorsements per RISC freebies.
			IF !ThisForm.Reject .AND. m.lnFreebieCount>0
				SELECT ChartC

				IF SEEK(m.polinumb) 
					i=m.lnImportCount+2

					SCAN WHILE ChartC.PoliNumb =m.PoliNumb	;
						FOR	ChartC.Amount=0	;
						.AND. (EMPTY(ChartC.Tipe) .OR. ALLTRIM(Web.Tipe)$ChartC.Tipe)	;
						.AND. !("X"$ChartC.Tipe)	;
						.AND. SEEK(ChartC.PoliNumb +ChartC.Endorse,"ChartB","PoliEnd")	;
						.AND. ChartB.lBndl=.F.
						ThisForm.laEndorse[i,1]	=ALLTRIM(ChartC.Endorse)
						ThisForm.laEndorse[i,2]	=ChartC.Amount
						ThisForm.laEndorse[i,3]	=''
						ThisForm.laEndorse[i,4]	=IIF(SEEK(ChartC.PoliNumb +ChartC.Endorse, "ChartB", "PoliEnd"), ChartB.nDispOrdr, i)
						ThisForm.laEndorse[i,5]	=ChartC.cForm
						i=(i+1)
					ENDSCAN

				ENDIF

			ENDIF

			*	Sith: Sort the array by it's column 4 (contains ChartB.nDispOrdr value)
			IF !ThisForm.Reject
				=ASORT(ThisForm.laEndorse,4)
			ENDIF


			SELECT Web

			IF ThisForm.Reject
				REPLACE Reject		WITH ThisForm.Reject, 	;
						Premendcal 	WITH ThisForm.EndorTTL, ;
						Processed 	WITH IIF(ThisForm.ErrorNum=40, .T., !ThisForm.Reject), ;
						EntryDate 	WITH m.EntryDate, ;
						EntryTime 	WITH m.EntryTime, ;
						ErrorNum 	WITH ThisForm.ErrorNum, ;
						ErrorTxt 	WITH ThisForm.ErrorTxt 
						WAIT WINDOW  "See table " +ALLTRIM(DBF()) +", record #"+ALLTRIM(STR(RECNO())) +": " ;
							+CHR(13)+ALLTRIM(Web.FirstName)+" "+ALLTRIM(Web.LastName)	;
							+CHR(13)+"Error: "+ Web.ErrorTxt	;
							+CHR(13)+"Record this message, then press any key to continue..."

				IF SEEK(Web.PoliNumb +Web.PoliID,"WebImp","PoliNumb")
					RECALL IN WebImp
					REPLACE WebImp.lHold WITH .T.
				ENDIF

				SKIP 1
				LOOP
			ENDIF

			IF EMPTY(m.polinumb)
				ThisForm.Reject = .T.
				ThisForm.ErrorNum = 70
				ThisForm.ErrorTxt = "missing polinumb"
			ENDIF

			IF !ThisForm.Reject

				* If this is a DUPE, assure that there is no PoliID.  Being [blank] will cause
				*	a new PoliID to be generated subsequently.
		*		IF !EMPTY(m.PoliID) .AND. (("DUPE"$m.PDANum) .OR. ("SEE NOTES"$m.PDANum))
		*			m.PoliID	=SPACE(LEN(m.PoliID))
		*		ENDIF
		* 20120702 PriscillaH:  [If it's a DUPE, I might choose to enter the matched PoliID and want it to persist.

				IF EMPTY(m.PoliID)
					* If processing an import that contains Internal Duplicates
					*	and this record isn't a DUPE, obtain an Internal PoliID
					IF (ThisForm.lFindIntPoliID=.T. .AND. !("DUPE"$m.PDANum) .AND. !("SEE NOTES"$m.PDANum))	;
						.OR. ThisForm.lMultiPoliID
						m.PoliID	=ThisForm.IntPoliID()
					ENDIF

					IF EMPTY(m.PoliID)
						m.PoliID	=ThisForm.PoliID()
					ENDIF

					REPLACE WebImp.PoliID	WITH m.PoliID
				ENDIF

				GATHER MEMVAR MEMO

				* Populate the 'front end' (PoliInd.DBF) -----------*
				*---------------------------------------------------*
				SELECT PoliInd
				APPEND BLANK		&&	PoliInd.dbf

				IF EMPTY(PoliInd.cbKey)
					m.cbKey	=GUID(36)
				ELSE
					m.cbKey	=PoliInd.cbKey
				ENDIF

				* Extract first of 1+ Policy Licenses
				m.License	=Thisform.GetLicense(Web.cPoliLic,1,"NUMBER")
				m.Tipe		=Thisform.GetLicense(Web.cPoliLic,1,"TYPE")

				* Receive Text Message?
				m.nCellNoti	=IIF(m.lTextMsg, 1, 0)

				* Set PoliInd.nPolicyTyp
				DO CASE
				CASE Web.cPolicyTyp="Individual"
					m.nPolicyTyp	=1
				CASE Web.cPolicyTyp="Firm"
					m.nPolicyTyp	=2
				OTHERWISE
					m.nPolicyTyp	=0
				ENDCASE

				* If GAP, blank the OrigDate
		*		IF "GAP"$m.pdanum
					m.OrigDate	={}	&&	Always make OrigDate blank
		*		ENDIF

				GATHER MEMVAR MEMO

				* Populate parent key into Web.dbf
				*	This will provide link for "extended data" fields
				REPLACE Web.cbPoliInd	WITH PoliInd.cbKey


				* Populate any additional STATE licenses into PoliLic
				IF ThisForm.GetLicense(Web.cPoliLic)>1
					PRIVATE m.lnLicCount, m.i
					m.lnLicCount	=ThisForm.GetLicense(Web.cPoliLic)
					SELECT PoliLic

					FOR m.i =2 TO m.lnLicCount
						APPEND BLANK
						REPLACE PoliLic.Polinumb	WITH PoliInd.Polinumb
						REPLACE PoliLic.PoliID		WITH PoliInd.PoliID
						REPLACE PoliLic.State		WITH ThisForm.GetLicense(Web.cPoliLic,m.i,"STATE")
						REPLACE PoliLic.Tipe		WITH ThisForm.GetLicense(Web.cPoliLic,m.i,"TYPE")
						REPLACE PoliLic.License		WITH ThisForm.GetLicense(Web.cPoliLic,m.i,"NUMBER")
					NEXT

					SELECT PoliInd
				ENDIF


				* Populate any additional CONFORMITY licenses into PoliLic
				IF ThisForm.GetLicense(Web.cConfLic)>0
					PRIVATE m.lnLicCount, m.i
					m.lnLicCount	=ThisForm.GetLicense(Web.cConfLic)
					SELECT PoliLic

					FOR m.i =1 TO m.lnLicCount
						APPEND BLANK
						REPLACE PoliLic.Polinumb	WITH PoliInd.Polinumb
						REPLACE PoliLic.PoliID		WITH PoliInd.PoliID
						REPLACE PoliLic.State		WITH ThisForm.GetLicense(Web.cConfLic,m.i,"STATE")
						REPLACE PoliLic.Tipe		WITH ThisForm.GetLicense(Web.cConfLic,m.i,"TYPE")
						REPLACE PoliLic.License		WITH ThisForm.GetLicense(Web.cConfLic,m.i,"NUMBER")
					NEXT

					SELECT PoliInd
				ENDIF


				* Populate the 'back end' (End_indv.DBF) ------------*
				*----------------------------------------------------*
				FOR m.lnCount = 1 TO ALEN(ThisForm.laEndorse,1)

					IF !EMPTY(ALLTRIM(m.FirstName)) 
						m.lcInsured = ALLTRIM(m.FirstName) +" " +ALLTRIM(m.LastName)
					ELSE
						m.lcInsured	= ALLTRIM(m.Company)
					ENDIF


					ThisForm.Trans(m.PoliNumb,;
						m.PoliID,;
						m.lcInsured, ;
						ThisForm.laEndorse[m.lnCount, 1],;
						ThisForm.laEndorse[m.lnCount, 2],;
						IIF(LEFT(ThisForm.laEndorse[m.lnCount, 1],7)="PREMIUM",Web.Taxes,0),;
						ALLTRIM(Web.cTxCtAuth)+IIF(!EMPTY(Web.cTxCtAuth) .AND. !EMPTY(Web.cTxCoAuth),", ","")+ALLTRIM(Web.cTxCoAuth),;
						ALLTRIM(ThisForm.laEndorse[m.lnCount, 3]), ;
						IIF(ThisForm.laEndorse[m.lnCount, 2] =0 .OR. LEFT(ThisForm.laEndorse[m.lnCount, 1],7)="PREMIUM" .OR. ALLTRIM(ThisForm.laEndorse[m.lnCount, 1])$"TN CONFORMITY", m.Effective, MAX(m.Effective, TTOD(m.tPurchase))),;
						m.end,;
						IIF(LEFT(Web.EntryBy,3)="XLS", "0", 'w')+m.batchnum,;
						m.entryby,;
						ThisForm.txtDate.Value,;
						TIME(),;
						ThisForm.laEndorse[m.lnCount, 4],;
						ThisForm.laEndorse[m.lnCount, 5],;
						m.cbKey)

					* 7. Tax_Auth:  IIF(LEFT(ThisForm.laEndorse[m.lnCount, 1],7)="PREMIUM",ALLTRIM(Web.cTxCtAuth)+IIF(!EMPTY(Web.cTxCtAuth) .AND. !EMPTY(Web.cTxCoAuth),", ","")+ALLTRIM(Web.cTxCoAuth),""),;

				NEXT

		*		* Populate any 'multi-license' (PoliLic.DBF) ---------*
		*		*-----------------------------------------------------*
		*		SELECT PoliLic
		*		SET ORDER TO IDNumb DESCENDING
		*
		*		SELECT PoliInd
		*
		*		IF SEEK(PoliInd.PoliID,"PoliLic")
		*			SELECT 0
		*
		*			SELECT Polilic.PoliNumb, PoliLic.PoliID, PoliLic.State, PoliLic.Tipe, PoliLic.cPoliLic;
		*			 FROM PoliLic;
		*			 WHERE  poliid = PoliInd.PoliID;
		*			 ORDER BY PoliNumb DESCENDING;
		*			 INTO CURSOR qPoliLic READWRITE 
		*			 
		*			IF RECCOUNT("qPoliLic")>0
		*				GO TOP
		*				PRIVATE m.lcRecentPolinumb
		*				m.lcRecentPolinumb=qPoliLic.PoliNumb
		*				DELETE FOR polinumb#m.lcRecentPolinumb
		*				REPLACE qPoliLic.PoliNumb WITH Web.PoliNumb FOR !DELETED()
		*
		*				SCAN FOR !DELETED()
		*					SCATTER MEMVAR
		*
		*					SELECT PoliLic
		*					APPEND BLANK
		*					GATHER MEMVAR
		*
		*					SELECT qPoliLic
		*				ENDSCAN
		*			ENDIF
		*
		*			USE IN qPoliLic
		*		ENDIF

			ENDIF

			SELECT Web

			* This might work later ...
			*-----------------------------------------------*
			* If any cConfLic, cApprLic or cApprTrain
			*
			* CONFORMITY licenses
		*	IF SEEK(polinumb+poliid+"CONFORMITY","End_indv","Polinumb")
		*		REPLACE End_Indv.mMulti	WITH Web.cConfLic
		*	ENDIF
		*	* APPRAISAL licenses
		*	IF SEEK(polinumb+poliid+"APPRAISAL","End_indv","Polinumb")
		*		REPLACE End_Indv.mMulti	WITH Web.cApprLic
		*	ENDIF
		*	* APPRAISER TRAINEE
		*	IF SEEK(polinumb+poliid+"APPRAISER TRAINEE","End_indv","Polinumb")
		*		REPLACE End_Indv.mMulti	WITH Web.cApprTrain
		*	ENDIF
			*-----------------------------------------------*


			REPLACE Reject	WITH ThisForm.Reject,			;
					premendcal	WITH ThisForm.EndorTTL, 	;
					Processed 	WITH !ThisForm.Reject,		;
					entrydate 	WITH m.entrydate,	;
					entrytime 	WITH m.entrytime, 	;
					errortxt 	WITH ThisForm.ErrorTxt

			SKIP 1
			SCATTER MEMVAR MEMO

			DO WHILE !EOF()

				DO CASE
				CASE Web.Processed

				CASE LEFT(Web.EntryBy, 3) #"XLS" .AND. ThisForm.fPDAnum(Web.pdanum)
					m.ErrorNum=40
					m.ErrorTxt="PDANUM already in poliind"
		*			m.Reject = .T.
					ThisForm.Reject=.T.
					m.Processed = .T.
					GATHER MEMVAR MEMO

				OTHERWISE
					m.ErrorNum = 0
					m.ErrorTxt = ""
		*			m.Reject = .T.
					ThisForm.Reject	=.T.
					EXIT
				ENDCASE

				SKIP 1
				SCATTER MEMVAR MEMO
			ENDDO

		ENDDO

		*	Close ChartA, ChartB & ChartC
		*DO Chart IN IIF(RIGHT(CURDIR(),4)="WEB\", "..\", "") +"PROGRAMS\LIBRARY" WITH "CLOSE"

		SELECT ChartC
		SET ORDER TO (m.lcCOrder)

		SELECT (m.lnSelect)
		WAIT CLEAR

		IF m.lcSetDeleted="OFF"
		*	SET DELETED OFF		&& let's try with DELETED on
		ENDIF

		RETURN
	ENDPROC


	PROCEDURE fpdanum
		PARAMETERS m.jcPDAnum
		PRIVATE m.lnSelect, m.lcOrder, m.lnRecno, m.llReturn, m.lcSetExact
		m.lnSelect	=SELECT()
		m.lcOrder	=ORDER()
		m.lnRecno	=RECNO()
		m.lcSetExact=SET("EXACT")

		IF m.lcSetExact="OFF"
			SET EXACT ON
		ENDIF

		SELECT PoliInd
		SET ORDER TO PDAnum

		IF AT(":", m.jcPDAnum)>0
			m.jcPDAnum	=ALLTRIM(SUBSTR(m.jcPDAnum, 1, AT(":", m.jcPDAnum)-1))
		ENDIF

		SEEK ALLTRIM(m.jcPDAnum)

		m.llReturn	=FOUND()

		SELECT (m.lnSelect)
		SET ORDER TO (m.lcOrder)
		GO (m.lnRecno)

		IF m.lcSetExact="OFF"
			SET EXACT OFF
		ENDIF

		RETURN m.llReturn
	ENDPROC


	PROCEDURE intpoliid
		PRIVATE m.lnSelect, m.lcIntPoliID
		m.lnSelect	=SELECT()

		SELECT IntDupe
		LOCATE FOR FirstName+LastName+Company = WebImp.FirstName+WebImp.LastName+WebImp.Company	;
			.AND. (RECNO() #RECNO("WebImp"))	;
			.AND. !EMPTY(PoliID)

		m.lcIntPoliID	=IntDupe.PoliID

		IF FOUND()
			m.OrigDate		=IntDupe.OrigDate	&& the scope of these mVars originate, 
			m.Notes			=ALLTRIM(m.Notes)+CHR(13)+ALLTRIM(IntDupe.Notes)
												&& and are private to, the calling program.  Redefining
												&& them here allows inheritance of the values of
												&& the previous policy
		ENDIF

		SELECT (m.lnSelect)
		RETURN m.lcIntPoliID
	ENDPROC


	PROCEDURE poliid
		PRIVATE m.PoliID, m.lnSelect
		m.PoliID	=""
		m.lnSelect	=SELECT()

		IF !USED("PoliID")
			USE PoliID SHARED IN 0
		ENDIF

		SELECT PoliID

		DO WHILE .T.

			IF RLOCK()
				REPLACE PoliID WITH PADL(ALLTRIM(STR(VAL(PoliID)+1)),8,"0")
				m.PoliID	=PoliID.PoliID
				UNLOCK
				EXIT
			ENDIF

			WAIT WINDOW NOWAIT "Waiting for a PoliID#..."
		ENDDO

		SELECT (m.lnSelect)
		RETURN m.PoliID
	ENDPROC


	PROCEDURE getendor2
		PRIVATE m.lnSelect, m.lcStYr, m.lcChartBOrder, m.llGotIt
		m.lnSelect	=SELECT()
		m.lcStYr	=SUBSTR(m.PoliNumb, 11, 2)+LEFT(m.PoliNumb, 2)

		SELECT ChartB
		m.lcChartBOrder	=ORDER()
		SET ORDER TO POLIIMP   && POLINUMB+CIMPORTVAR

		SELECT (m.lnSelect)
		*m.govstate	=GovState(m.govstate)
		m.DimrCnt 	=0
		m.endorTTL	=0
		m.llGotIt	=.T.	&& default to TRUE
		ThisForm.Reject = .F.

		* Higher Limit - 1 ----------------------------------------------*
		*IF m.llGotIt .AND. m.HigherLim1 
		IF m.HigherLim1 
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"HIGHERLIM1", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt, 1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt, 2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt, 3] = ''
				ThisForm.laEndorse[m.dimrCnt, 4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt, 5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 10
				ThisForm.ErrorTxt = m.lcStYr +' HIGHERLIM1 !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Higher Limit - 2 ----------------------------------------------*
		*IF m.llGotIt .AND. m.HigherLim2
		IF m.HigherLim2
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"HIGHERLIM2", "ChartB") .AND.	SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 12
				ThisForm.ErrorTxt = m.lcStYr +' HIGHERLIM2 !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Higher Limit Mil - 1 ------------------------------------------*
		*IF m.llGotIt .AND. m.HiLimMil1
		IF m.HiLimMil1
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"HILIMMIL1", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 10
				ThisForm.ErrorTxt = m.lcStYr +' HILIMMIL1 !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Higher Limit Mil - 2 ------------------------------------------*
		*IF m.llGotIt .AND. m.HiLimMil2
		IF m.HiLimMil2
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"HILIMMIL2", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 12
				ThisForm.ErrorTxt = m.lcStYr +' HILIMMIL2 !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Appraisal -----------------------------------------------------*
		*IF m.llGotIt .AND. m.Appraisal
		IF m.Appraisal
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"APPRAISAL", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 13
				ThisForm.ErrorTxt = m.lcStYr +' APPRAISAL !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Property Management -------------------------------------------*
		*IF m.llGotIt .AND. m.PropMgr
		IF m.PropMgr
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"PROPMGR", "ChartB") 	.AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 14
				ThisForm.ErrorTxt = m.lcStYr +' PROPMGR !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Environmental -------------------------------------------------*
		*IF m.llGotIt .AND. m.Environx
		IF m.Environx
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"ENVIRONX", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 15
				ThisForm.ErrorTxt = m.lcStYr +' ENVIRONX !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* LC Environmental ----------------------------------------------*
		*IF m.llGotIt .AND. m.LCEnviro
		IF m.LCEnviro
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"LCENVIRO", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 15
				ThisForm.ErrorTxt = m.lcStYr +' LCENVIRO !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Fair Housing --------------------------------------------------*
		*IF m.llGotIt .AND. m.FairHouse && .AND. !(ALLTRIM(m.GovState)$"ID" .AND. LEFT(m.polinumb,2)="09")
		IF m.FairHouse 
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"FAIRHOUSE", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIT
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 16
				ThisForm.ErrorTxt = m.lcStYr +' FAIRHOUSE !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF
		ENDIF
		*----------------------------------------------------------------*

		* Fair Housing A ------------------------------------------------*
		*IF m.llGotIt .AND. m.FairHausA
		IF m.FairHausA
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"FAIRHAUSA", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 16
				ThisForm.ErrorTxt = m.lcStYr +' FAIRHAUSA !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Fair Housing B ------------------------------------------------*
		*IF m.llGotIt .AND. m.FairHausB
		IF m.FairHausB
			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"FAIRHAUSB", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 16
				ThisForm.ErrorTxt = m.lcStYr +' FAIRHAUSB !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Complaints ----------------------------------------------------*
		*IF m.llGotIt .AND. m.Complaints
		IF m.Complaints
			m.DimrCnt 	=m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"COMPLAINTS", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' COMPLAINTS !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Online Charge -------------------------------------------------*
		*IF m.llGotIt .AND. m.OnlineChg
		IF m.OnlineChg
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

		*	IF SEEK(SPACE(LEN(m.polinumb)) +"ONLINECHG", "ChartB") 
			=SEEK(SPACE(LEN(m.polinumb)) +"ONLINECHG", "ChartB") 
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartB.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartB.nExpClm
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ""
		*	ELSE
			IF !FOUND("ChartB")
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = ' ONLINECHG !found ChartB'
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject 	=.T.
				ThisForm.llGotIt	=.F.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Earnest Money Dispute -----------------------------------------*F.
		*IF m.llGotIt .AND. m.EarnMo
		IF m.EarnMo
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"EARNMO", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' EARNMO !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Primary Residence ---------------------------------------------*
		IF m.PrimeRes
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"PRIMERES", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' PRIMERES !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* States --------------------------------------------------------*
		*IF m.CO .OR. m.KY .OR. m.NE .OR. m.RI .OR. m.TN ;
			.OR. m.LA .OR. m.NM .OR. m.SD .OR. m.IA ;
			.OR. m.MS .OR. m.ND .OR. m.ID .OR. m.WY
		IF !EMPTY(m.cConfLic)

		*	m.StStr = IIF(m.CO, 'CO,', '') ;
					+IIF(m.KY, 'KY,', '') ;
					+IIF(m.NE, 'NE,', '') ;
					+IIF(m.RI, 'RI,', '') ;
					+IIF(m.TN, 'TN,', '') ;
					+IIF(m.LA, 'LA,', '') ;
					+IIF(m.NM, 'NM,', '') ;
					+IIF(m.SD, 'SD,', '') ;
					+IIF(m.IA, 'IA,', '') ;
					+IIF(m.MS, 'MS,', '') ;
					+IIF(m.ND, 'ND,', '') ;
					+IIF(m.ID, 'ID,', '') ;
					+IIF(m.WY, 'WY,', '')

			PRIVATE m.lnSelect, m.lcStates
			m.lnSelect	=SELECT()

			SELECT StatesU
			m.lcStates	=""

			SCAN
				IF StatesU.state+"." = LEFT(m.cConflic, 3) .OR. (";"+StatesU.state$m.cconflic)
					m.lcStates	=(m.lcStates +ALLTRIM(StatesU.State)+",")
				ENDIF
			ENDSCAN

			SELECT (m.lnSelect)


			m.lcStates	=IIF(!EMPTY(m.lcStates),LEFT(m.lcStates, LEN(m.lcStates)-1),"")	&& slough off the right-most ','

			m.DimrCnt	=(m.DimrCnt +1)
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.PoliNumb +"CONFORMITY", "ChartB") .AND.	SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.DimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.DimrCnt,2] = ChartC.Amount
		*		ThisForm.laEndorse[m.DimrCnt,3] = IIF(!EMPTY(m.lcStates),LEFT(m.lcStates, LEN(m.lcStates)-1),"")
				ThisForm.laEndorse[m.DimrCnt,3] = m.lcStates
				ThisForm.laEndorse[m.DimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 18
				ThisForm.ErrorTxt = m.lcStYr +' CONFORMITY !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Premium A/B-----------------------------------------------------*
		m.DimrCnt = m.DimrCnt + 1
		m.llGotIt	=.T.

		DO CASE
		* PremiumA
		CASE Web.PremiumA=.T. .AND. SEEK(m.polinumb +"PREMIUMA", "ChartB") 
		* PremiumB
		CASE Web.PremiumB=.T. .AND. SEEK(m.polinumb +"PREMIUMB", "ChartB") 
		* regular "PREMIUM"
		CASE SEEK(m.polinumb +"PREMIUM", "ChartB") 
		* not found
		OTHERWISE
			m.llGotIt	=.F.
		ENDCASE

		IF m.llGotIt	=.T.
			ThisForm.laEndorse(m.dimrCnt,1) = ALLTRIM(ChartB.Endorse)
			ThisForm.laEndorse(m.dimrCnt,2) = Web.Prem
			ThisForm.laEndorse(m.dimrCnt,3) = ''
			ThisForm.laEndorse(m.dimrCnt,4) = ChartB.nDispOrdr
			ThisForm.laEndorse(m.dimrCnt,5) = ChartB.cForm
		ELSE
		*	ThisForm.laEndorse(m.dimrCnt,1) = 'prorate not found'
			ThisForm.laEndorse(m.dimrCnt,2) = 0.00
			ThisForm.laEndorse(m.dimrCnt,3) = ''
			ThisForm.laEndorse(m.dimrCnt,4) = 0
			ThisForm.ErrorNum = 20
			ThisForm.ErrorTxt = m.lcStYr +'PREMIUM ' +IIF(Web.PremiumB=.T., 'B ', '') +'prorate !found ChartB'
			ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
			ThisForm.Reject = .T.
		ENDIF
		*----------------------------------------------------------------*

		* Residential Personal Interest ---------------------------------*
		IF m.ResPersInt
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"RESPERSINT", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' RESPERSINT !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Bundle (BNDL)--------------------------------------------------*
		IF m.Bndl
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"BNDL"+SPACE(LEN(ChartB.cImportVar)-4)+"0", "ChartB","PoliImp") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' BNDL !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* Bodyily Injury (BODYINJ)---------------------------------------*
		IF m.BodyInj
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"BODYINJ", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' BODYINJ !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*

		* APPRAISER TRAINEE(APPTRAN)-------------------------------------*
		IF OCCURS(";",m.cApprTrain)>0
			m.llGotIt	=.T.
			PRIVATE m.lnAPPTRANcnt, m.i
			m.lnAPPTRANcnt=OCCURS(";",m.cApprTrain)

			DO CASE
			CASE SEEK(m.polinumb +"APPTRAN", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
			FOR m.i=1 TO m.lnAPPTRANcnt
				m.DimrCnt = m.DimrCnt + 1
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				IF m.i=1
					ThisForm.laEndorse[m.dimrCnt,3] = SUBSTR(m.capprtrain,1,AT(";",m.capprtrain,1)-1)
				ELSE
					ThisForm.laEndorse[m.dimrCnt,3] = SUBSTR(m.cApprTrain,AT(";",m.capprtrain,m.i-1)+1,AT(";",m.capprtrain,m.i)-AT(";",m.capprtrain,m.i-1)-1)
				ENDIF
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
			NEXT
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' APPTRAN !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*


		* SURCHARGE (SURCHARGE)------------------------------------------*
		IF m.Surcharge#0
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(SPACE(LEN(m.polinumb)) +"SURCHARGE", "ChartB") && .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
		*		.AND. Web.Effective>=ChartC.Effective
		*	CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
		*	CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartB.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = m.Surcharge
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartB.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = "Universal SURCHARGE !found ChartB"
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*


		* DEVELOPED/CONSTRUCTED (DEVCON)---------------------------------*
		IF m.DevCon
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"DEVCON", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' DEVCON !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*


		* SUBPOENA (SUBPOENA)---------------------------------*
		IF m.Subpoena
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"SUBPOENA", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' SUBPOENA !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*


		* RM10 (REVERSE MORTGAGE LOAN TRANSACTION - DMG - 10K DED)---------------------------------*
		IF m.RM10
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"RM10", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' RM10 !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*


		* RM20 (REVERSE MORTGAGE LOAN TRANSACTION - DMG - 20K DED)---------------------------------*
		IF m.RM20
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"RM20", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' RM20 !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*


		* Lockbox [LOCKBOX... ]--------------------------------------*
		IF m.Lockbox
			m.DimrCnt = m.DimrCnt + 1
			m.llGotIt	=.T.

			DO CASE
			CASE SEEK(m.polinumb +"LOCKBOX", "ChartB") .AND. SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),6), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse +LEFT(DTOS(Web.Effective),4), "ChartC")	;
				.AND. Web.Effective>=ChartC.Effective
			CASE SEEK(ChartB.PoliNumb +ChartB.Endorse, "ChartC")
			OTHERWISE
				m.llGotIt	=.F.
			ENDCASE

		*	IF m.llGotIt
				ThisForm.laEndorse[m.dimrCnt,1] = ALLTRIM(ChartC.Endorse)
				ThisForm.laEndorse[m.dimrCnt,2] = ChartC.Amount
				ThisForm.laEndorse[m.dimrCnt,3] = ''
				ThisForm.laEndorse[m.dimrCnt,4] = ChartB.nDispOrdr
				ThisForm.laEndorse[m.dimrCnt,5] = ChartC.cForm
		*	ELSE
			IF !m.llGotIt
				ThisForm.ErrorNum = 17
				ThisForm.ErrorTxt = m.lcStYr +' LOCKBOX !found ' +IIF(!FOUND("ChartB"), "ChartB", "ChartC")
				ThisForm.laEndorse[m.dimrCnt, 1] = ThisForm.ErrorTxt
				ThisForm.Reject = .T.
			ENDIF

		ENDIF
		*----------------------------------------------------------------*


		* Check if imported amount$ matches tally of endorsements -------*
		FOR m.lnCount = 1 TO (m.lnImportCount +1)
			ThisForm.EndorTTL =(ThisForm.endorTTL +ThisForm.laEndorse[m.lnCount,2])
		NEXT

		IF ThisForm.EndorTTL <>(m.PremImport-Web.Taxes)
			ThisForm.Reject = .T.
			ThisForm.ErrorNum = 30
			ThisForm.ErrorTxt = "Does not add up"
		ENDIF
		*----------------------------------------------------------------*

		SELECT ChartB
		SET ORDER TO (m.lcChartBOrder)

		SELECT (m.lnSelect)
		RETURN
	ENDPROC


	PROCEDURE trans
		PARAMETERS m.PoliNumb,;
			m.PoliID,;
			m.Insured,;
			m.Endorse,;
			m.Premium,;
			m.Tax,;
			m.Tax_Auth,;
			m.State_Eff,;
			m.Effective,;
			m.Expires,;
			m.Batch_ID,;
			m.EntryBy,;
			m.EntryDate,;
			m.EntryTime,;
			m.nDispOrdr,;
			m.cForm,;
			m.cbPoliInd

		PRIVATE m.lnSelect, m.cbKey, m.lnSelect1
		m.lnSelect	=SELECT()
		m.cbKey	=GUID(36)

		m.lnSelect1	=SELECT()
		SELECT End_Indv
		APPEND BLANK		&&	End_Indv.dbf
		GATHER MEMVAR
		*INSERT INTO End_Indv FROM MEMVAR	"INSERT" does not activate End_IndvDB triggers
		SELECT (m.lnSelect1)

		* Make a '0.00' display in the Tax field
		IF End_Indv.Tax =0
			REPLACE	End_Indv.Tax	WITH 0
		ENDIF

		* If this is a Bundle header... (lBndl is Bundle, nBndl=0 is header)
		IF SEEK(m.PoliNumb +m.Endorse,"ChartB","PoliEnd") .AND. ChartB.lBndl .AND. ChartB.nBndl=0
			m.lccImportVar=ChartB.cImportVar	&& grab the import variable.  It has been entered in all the bundled endorsements' cImportVar)
			m.lnPreBndlSelect=SELECT()

			* collect the Bundle endorsement items...
			*	- for this ImportVar
			*	- chartb entry is not a Bundle header (ChartB.nBndl>0)
			*	- endorsement is not excluded from entering database ("X"$ChartC.Tipe =.F.)
			SELECT 0
			SELECT Chartb.endorse, Chartb.cimportvar, Chartc.amount, Chartc.cform, Chartb.ckey;
			 FROM ;
			     chartb ;
			    INNER JOIN chartc ;
		 	    ON  Chartb.ckey = Chartc.ckchartb;
			 WHERE  Chartb.polinumb = End_Indv.PoliNumb;
			   AND  Chartb.cimportvar = m.lccImportVar;
			   AND  Chartb.nbndl > 0;
			   AND  "X"$ChartC.Tipe = .F.;
			 ORDER BY Chartb.nBndl;
			 INTO CURSOR qBndl  

			IF RECCOUNT()>0
				SELECT End_Indv
				PRIVATE m.cbKey, m.cbChartB, m.Premium, m.cForm
				SCATTER MEMVAR

				SELECT qBndl
				SCAN
					m.cbKey		=GUID(36)
					m.cbChartB	=qBndl.cKey
					m.Endorse	=qBndl.Endorse
					m.Premium	=qBndl.Amount
					m.cForm		=qBndl.cForm

					SELECT End_Indv
					APPEND BLANK
					GATHER MEMVAR

					SELECT qBndl
				ENDSCAN
			ENDIF

			USE IN qBndl
			SELECT (m.lnPreBndlSelect)

		ENDIF

		SELECT (m.lnSelect)
	ENDPROC


	PROCEDURE parseit
		* .ParseIt
		*	Creates *.Ed fioe, which contains copy of data 
		*		- after DatAssist()/user editing 
		*		- beofre imported to BML
		*
		*	Create new PoliID numbers for 
		*		- unmatched (iow, new)
		*		- DUPEs 
		*
		* 	Moves records from WebImp to Web if not on "hold", deleting from WebImp after moved.
		*
		*		Webimp.dbf
		*			|
		*			L--Web.dbf
		*			+->Web.cbKey = GUID(36)
		*
		*	Delete from WebImp for not on "Hold" (.lHold) by current <cUsername>
		*
		*=============================================================================================

		*IF FILE("Step.on")
		*	SET STEP ON 
		*ENDIF

		SELECT PoliInd
		m.lcPoliIndOrder	=ORDER()
		SET ORDER TO POLIID   DESCENDING && UPPER(POLIID)

		SELECT WebImp
		*COPY TO STUFF(m.lcImpFile, AT(".", m.lcImpFile), 4, ".Ed") SDF
		m.lcSetSafety	=SET("Safety")
		SET SAFETY OFF
		COPY TO STUFF(ThisForm.txtImportFile.Value, AT(".", ThisForm.txtImportFile.Value), 4, ".Ed") SDF

		IF m.lcSetSafety="ON"
			SET SAFETY ON
		ENDIF

		GO TOP

		SCAN FOR !WebImp.lHold	.and. ALLTRIM(WebImp.cUsername)=ALLTRIM(ThisForm.txtUser.Value) &&	WebImp.DBF for not on Hold and for current user
			SCATTER MEMVAR MEMO

			* If this is a DUPE, assure that there is no PoliID.  Being [blank] will cause
			*	a new PoliID to be generated subsequently.
			*	IF !EMPTY(m.PoliID) .AND. (("DUPE"$m.PDANum) .OR. ("SEE NOTES"$m.PDANum))
			*		m.PoliID	=SPACE(LEN(m.PoliID))
			*	ENDIF

			IF EMPTY(m.PoliID)

				IF (ThisForm.lFindIntPoliID .AND. !("DUPE"$m.PDANum) .AND. !("SEE NOTES"$m.PDANum))
		*			.OR. ThisForm.lMultiPoliID
					* If processing an import that contains Internal Duplicates
					*	and this record isn't a DUPE, obtain an Internal PoliID
					m.PoliID	=ThisForm.IntPoliID()
				ENDIF

				IF EMPTY(m.PoliID)
					m.PoliID	=ThisForm.PoliID()
				ENDIF

				REPLACE WebImp.PoliID	WITH m.PoliID
			ENDIF

			* Retrieve latest PoliInd.Notes 
			*	- it has been found, on rare occassions, that a Policy user may edit a PoliInd.Notes after a WebimportCSV batch process was begun.
			*	- retrieving the very latest 'version' of PoliInd.Notes at this point should improve the assurance of getting any interim edits.
			=SEEK(WebImp.PoliID,"PoliInd")	&& PoliInd.dbf is SET ORDER TO POLIID   DESCENDING && UPPER(POLIID), making the found record the most recent, "on top", data
			* Replace WebImp.Notes with (WebImp.NewNotes+PoliInd.Notes)
			IF !("DUPE"$WebImp.PDANum)
				REPLACE WebImp.Notes WITH ALLTRIM(WebImp.mNewNotes)+IIF(FOUND("PoliInd"), CHR(13)+CHR(13)+ALLTRIM(PoliInd.Notes),"")
			ELSE
				REPLACE WebImp.Notes WITH ALLTRIM(WebImp.mNewNotes)
			ENDIF

			SCATTER MEMVAR MEMO

			SELECT Web
			m.cImpfile	=JUSTFNAME(ThisForm.txtImportFile.Value)
			m.cImpfileDT=ThisForm.cImpfileDT
			APPEND BLANK
			m.cbKey	=GUID(36)
			GATHER MEMVAR MEMO

			SELECT WebImp
		ENDSCAN

		DELETE FOR !WebImp.lHold .and. ALLTRIM(Webimp.cUsername)=ALLTRIM(Thisform.txtUser.Value)

		SELECT PoliInd
		SET ORDER TO (m.lcPoliIndOrder)

		SELECT WebImp
		RETURN .T.
	ENDPROC


	PROCEDURE fixbool
		PARAMETERS m.xcBool
		RETURN IIF(m.xcBool='1', .T., .F.)
	ENDPROC


	PROCEDURE getprem
		PRIVATE m.lnSelect, m.lnPrem, m.lnLocateAmt
		m.lnLocateAmt	=0.00
		m.lnSelect	=SELECT()
		SELECT ChartC

		DO CASE
		* PremiumA
		CASE WebImp.PremiumA=.T.
			SET FILTER TO LEFT(ChartC.Endorse, 9)="PREMIUM A"
		* PremiumB
		CASE WebImp.PremiumB=.T.
			SET FILTER TO LEFT(ChartC.Endorse, 9)="PREMIUM B"
		* regular "PREMIUM "...
		OTHERWISE
			SET FILTER TO 
		ENDCASE

		IF EOF() .AND. SEEK(WebImp.PoliNumb)
			LOCATE REST FOR ChartC.Effective > WebImp.Effective	WHILE ChartC.Polinumb=WebImp.PoliNumb
			SKIP -1
			IF ChartC.polinumb=WebImp.polinumb .and. BETWEEN(WebImp.effective, ChartC.effective, WebImp.end)
				m.lnLocateAmt	=ChartC.Amount
			ENDIF
		ENDIF

		SELECT (m.lnSelect)
		GO RECNO()
		m.lnPrem	=IIF(m.lnLocateAmt#0.00, m.lnLocateAmt, ChartC.Amount)

		IF WebImp.PremiumA=.T. .OR. WebImp.PremiumB=.T.
			SELECT ChartC
			SET FILTER TO 

			SELECT (m.lnSelect)
			GO RECNO()
		ENDIF

		RETURN m.lnPrem
	ENDPROC


	PROCEDURE getimportfile
		ThisForm.txtImportFile.Value   =GETFILE('TXT',"Select import file","Pick Me",0,"Get file...")

		IF !EMPTY(ThisForm.txtImportFile.Value)
			PRIVATE laImportFile
			=ADIR(laImportFile, ThisForm.txtImportFile.Value)
			ThisForm.cmdDataAssist.Enabled=.T.
			APPEND FROM (ThisForm.txtImportFile.Value) SDF

			* Port un-edited records to BML\Quotes.dbf
			GO TOP
		*	LOCATE REST FOR ThisForm.WebRec2Quotes(laImportFile[1], DTOC(laImportFile[3]) +" @ " +laImportFile[4]) .AND. .F.
			LOCATE REST FOR ThisForm.WebRec2Quotes(ALLTRIM(Thisform.txtImportFile.Value), DTOC(laImportFile[3]) +" @ " +laImportFile[4]) .AND. .F.

			* Create purchase data as T data type.
			REPLACE WebImp.tPurchase WITH CTOT(STUFF(WebImp.cPurchase,11,1,"T")) FOR !WebImp.lHold .AND. !DELETED()

			** remove the SS# where "0000"
			REPLACE ALL WebImp.cSSN4 WITH "" FOR WebImp.cSSN4 ="0000" .AND. !WebImp.lHold .AND. !DELETED()

			** Extract SSN last-four, populate cSSN4
			*	for Individual
			** Copy the SSN content to cEIN
			*	for Firm
			REPLACE ALL WebImp.cSSN4	WITH RIGHT(ALLTRIM(WebImp.cSSN4),4) FOR !EMPTY(WebImp.Firstname) .AND. !WebImp.lHold .AND. !DELETED()
			REPLACE ALL WebImp.cEIN	WITH ALLTRIM(WebImp.cSSN4) FOR EMPTY(WebImp.Firstname) .AND. !WebImp.lHold .AND. !DELETED()

			* Assure Firm name is in Lastname
			REPLACE ALL WebImp.Lastname	WITH WebImp.Company FOR EMPTY(WebImp.Firstname) .AND. !WebImp.lHold .AND. !DELETED()

			* Set 'Policy Type'
			REPLACE ALL WebImp.nPolicyTyp	WITH IIF(!EMPTY(WebImp.Firstname),1,2)	FOR !WebImp.lHold .AND. !DELETED()

			* Place User name in WebImp rile.
			REPLACE ALL Webimp.cUsername 	WITH ALLTRIM(ThisForm.TxtUser.Value)	FOR !WebImp.lHold .AND. !DELETED()

			GO TOP
		*	ThisForm.grdWebImp.Refresh()
			ThisForm.grdWebImp.SetFocus
		ENDIF
	ENDPROC


	PROCEDURE noteedit
		MODIFY MEMO NoteTxt.NoteTxt
	ENDPROC


	PROCEDURE webrec2quotes
		*	FUNCTION WebRec2Quotes
		PARAMETERS m.xcImpFile, m.xcImpFileDT
		PRIVATE m.lnSelect, m.llAddThis
		m.lnSelect	=SELECT()
		m.llAddThis	=!SEEK(LEFT(WebImp.PDAnum,LEN(Quotes.PDA)),"Quotes","PDA")		&& if not found by PDA, add this


		IF m.llAddThis
			* Quotes / WebImp
			*-----------------------------
			m.PDA		=WebImp.PDAnum
			m.STATE		=WebImp.State
			m.INDORFIRM	=IIF(!EMPTY(WebImp.FirstName),"individual","firm")
			m.FNAME		=WebImp.FirstName
			m.MNAME		=WebImp.MiddleName
			m.LNAME		=WebImp.LastName
			m.FIRM		=WebImp.Company
			m.RESSTATE	=WebImp.State
			m.EFFECTIVE	=CTOD(LEFT(DTOC(WebImp.effective),6)+RIGHT(DTOC(WebImp.effective),2))
			m.MANSTATE	=WebImp.State
			m.APPRAISAL	=IIF(WebImp.Appraisal=.t.,'1','0')
			m.PROPMGMT	=IIF(WebImp.PropMgr=.T.,'1','0')
			m.ENVIRO	=IIF(WebImp.Environx=.T.,'1','0')
			m.FAIRHOUSE	=IIF(WebImp.FairHouse=.T.,'1','0')
			m.COMPLAINTS=IIF(WebImp.Complaints=.t.,'1','0')
			*m.HILIMCHK	=WebImp.HigherLim1
			m.HILIM		=IIF(WebImp.HigherLim1=.T., '1', IIF(WebImp.HigherLim2=.T., '2', '0'))
			m.PREM_A	=IIF(WebImp.PremiumA=.T.,'1','0')
			m.PREP_B	=IIF(WebImp.PremiumB=.T.,'1','0')
			m.EARNEST	=IIF(WebImp.EarnMo=.T.,'1','0')
			m.LC_ENVIRO	=IIF(WebImp.LcEnviro=.T.,'1','0')
			m.FAIRHOUSEA=IIF(WebImp.FairHausA=.T.,'1','0')
			m.FAIRHOUSEB=IIF(WebImp.FairHausB=.T.,'1','0')
			m.PRIMERES	=IIF(WebImp.PrimeRes=.T.,'1','0')
			*m.HILIMCHKB	=WebImp.HiLimMil1
			m.HIMLIMB	=IIF(WebImp.HiLimMil1=.T., '1', IIF(WebImp.HiLimMil2=.T., '2', '0'))
			m.HADCLAIM	=IIF(WebImp.ClaimAlert=.T.,'1','0')
			m.POLITOTAL	=ALLTRIM(STR(WebImp.Prem))
			m.EMAIL		=WebImp.Email
			m.POLISTREET=ALLTRIM(WebImp.Address1)+IIF(!EMPTY(WebImp.Address2),", "+ALLTRIM(WebImp.Address2),"")
			m.POLICITY	=WebImp.City
			m.POLISTATE	=WebImp.State
			m.POLIZIP	=ALLTRIM(WebImp.ZIP)+IIF(!EMPTY(WebImp.ZIP4),"-"+ALLTRIM(WebImp.ZIP4),"")
			m.SSNTAXID	=ALLTRIM(WebImp.cSSN4)+ALLTRIM(WebImp.cEIN)
		*	m.LLICTYPE	=WebImp.Tipe
			m.cPoliLic	=WebImp.cPoliLic
			m.LICPENDING=""
			m.FIRMNAME	=WebImp.Company
			m.PHONE		=WebImp.Phone1
			m.FAX		=WebImp.FAX
			m.TRANSACTID=WebImp.cTransact
			m.RESPERS	=IIF(WebImp.Respersint=.T.,'1','0')
			m.CREATED	=TTOC(WebImp.tPurchase)	&& CTOD(SUBSTR(WebImp.cpurchase,6,5)+"-"+LEFT(WebImp.cpurchase,4))
			m.BUNDLE	=IIF(WebImp.Bndl=.T.,'1','0')
			m.BODYPROPDM=IIF(WebImp.BodyInj=.T.,'1','0')
			m.dPoliFrom	=Webimp.Effective
			m.dPoliThru	=Webimp.End
		*	m.PoliNumb	=ThisForm.GetPolinumb(WebImp.State, CTOD(LEFT(DTOC(WebImp.effective),6)+RIGHT(DTOC(WebImp.effective),2)), @m.dPoliFrom, @m.dPoliThru)
			m.PoliNumb	=ThisForm.GetPolinumb(WebImp.State, Webimp.Effective)
			m.mEndorse	=""
			m.cImpFile	=m.xcImpFile
			m.cImpFileDT=m.xcImpFileDT
			m.cPmtID	=WebImp.cPmtID

			SELECT Quotes
			APPEND BLANK
			GATHER MEMVAR MEMO
			REPLACE Quotes.mEndorse WITH ThisForm.GetEndorse(Quotes.polinumb)
		ENDIF

		SELECT (m.lnSelect)
		* m.llAddThis =.T., because it was added
		* m.llAddThis =.F., because it was found by PDA already in table
		RETURN m.llAddThis
	ENDPROC


	PROCEDURE getendorse
		*	FUNCTION GetEndorse
		PARAMETERS m.xcPolinumb
		PRIVATE m.lmEndorse, m.llChartBUsed, m.lnChartBrecno
		m.lmEndorse	=""	&&	default to [blank]
		m.llChartBUsed	=USED("ChartB")

		IF !m.llChartBUsed
			USE ChartB IN 0
		ELSE
			m.lnChartBrecno	=RECNO("ChartB")
		ENDIF

		IF Quotes.APPRAISAL	='1' .AND. SEEK(m.xcPoliNumb+"APPRAISAL ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.PROPMGMT	='1' .AND. SEEK(m.xcPoliNumb+"PROPMGR   ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.ENVIRO	='1' .AND. SEEK(m.xcPolinumb+"ENVIRONX  ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.FAIRHOUSE	='1' .AND. SEEK(m.xcPolinumb+"FAIRHOUSE ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.COMPLAINTS='1' .AND. SEEK(m.xcPolinumb+"COMPLAINTS", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		*IF Quotes.HILIMCHK	='1' .AND. SEEK(m.xcPolinumb+"HIGHERLIM1", "ChartB", "PoliImp")
		*	m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		*ENDIF

		IF Quotes.HILIM		='1' .AND. SEEK(m.xcPolinumb+"HIGHERLIM1", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.HILIM		='2' .AND. SEEK(m.xcPolinumb+"HIGHERLIM2", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.PREM_A	='1' .AND. SEEK(m.xcPolinumb+"PREMIUMA  ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.PREP_B	='1' .AND. SEEK(m.xcPolinumb+"PREMIUMB  ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.EARNEST	='1' .AND. SEEK(m.xcPolinumb+"EARNMO    ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.LC_ENVIRO	='1' .AND. SEEK(m.xcPolinumb+"LCENVIRO  ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.FAIRHOUSEA='1' .AND. SEEK(m.xcPolinumb+"FAIRHAUSA ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.FAIRHOUSEB='1' .AND. SEEK(m.xcPolinumb+"FAIRHAUSB ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.PRIMERES	='1' .AND. SEEK(m.xcPolinumb+"PRIMERES  ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		*IF Quotes.HILIMCHKB ='1' .AND. SEEK(m.xcPolinumb+"HILIMMIL1 ", "ChartB", "PoliImp")
		*	m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		*ENDIF

		IF Quotes.HIMLIMB	='1' .AND. SEEK(m.xcPolinumb+"HILIMMIL1 ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.HIMLIMB	='2' .AND. SEEK(m.xcPolinumb+"HILIMMIL2 ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.RESPERS	='1' .AND. SEEK(m.xcPolinumb+"RESPERSINT", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.BUNDLE	='1' .AND. SEEK(m.xcPolinumb+"BNDL      ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF Quotes.BODYPROPDM='1' .AND. SEEK(m.xcPolinumb+"BODYINJ   ", "ChartB", "PoliImp")
			m.lmEndorse	=m.lmEndorse +ALLTRIM(ChartB.Endorse) +CHR(13)
		ENDIF

		IF !m.llChartBUsed
			USE IN ChartB
		ELSE
			GO (m.lnChartBrecno) IN ChartB
		ENDIF

		RETURN m.lmEndorse
	ENDPROC


	PROCEDURE getpolinumb
		*	FUNCTION GetPolinumb
		PARAMETERS m.xcState, m.xdEffective	&&	, m.xdPoliFrom, m.xdPoliThru
		PRIVATE m.lnSelect, m.lcPolinumb, m.llChartAused, m.lcChartAorder, m.lnChartArecno
		m.lnSelect	=SELECT()

		m.llChartAused	=USED("ChartA")

		IF m.llChartAused
			SELECT ChartA
			m.lnChartArecno	=RECNO()
			m.lcChartAorder	=ORDER()
		ELSE
			SELECT 0
			USE ChartA
		ENDIF

		SET ORDER TO STATEYR   && STATE+LEFT(POLINUMB,2)
		SEEK m.xcState
		LOCATE REST WHILE SUBSTR(ChartA.Polinumb,11,2)=m.xcState	;
			FOR BETWEEN(m.xdEffective, ChartA.Effective, ChartA.End-1)	;
			.AND. SUBSTR(ChartA.Polinumb, 4, 2)=IIF(WebImp.cCovType="RE","EO","AP")
		m.lcPolinumb	=IIF(EMPTY(m.xdEffective), "", ChartA.Polinumb)
		*m.xdPoliFrom	=ChartA.Effective
		*m.xdPoliThru	=ChartA.End

		IF m.llChartAused
			SET ORDER TO (m.lcChartAorder)
			GO (m.lnChartArecno)
		ELSE
			USE IN ChartA
		ENDIF

		SELECT (m.lnSelect)
		RETURN m.lcPolinumb
	ENDPROC


	PROCEDURE setuserresource
		PARAMETERS m.xcUser
		* If no resource file is yet established for the app/user,
		*	create it from the empty 'newuser' table.
		IF !FILE("RESOURCEFILES\"+m.xcUser+".DBF")
			PRIVATE m.lnSelect
			m.lnSelect	=SELECT()
			SELECT 0
			USE ("RESOURCEFILES\newuser")
			COPY TO ("RESOURCEFILES\"+m.xcuser)
			USE
			SELECT (m.lnSelect)
		ENDIF

		SET RESOURCE TO ("RESOURCEFILES\"+m.xcuser)
	ENDPROC


	PROCEDURE browseedit
		PRIVATE m.lnSelect, m.lcBrowseColumns
		m.lnSelect =SELECT()

		SELECT WebImp
		m.lcBrowseColumns=Thisform.Browsecolumns()
		BROWSE LAST FIELDS &lcBrowseColumns

		SELECT (m.lnSelect)
		RETURN .T.
	ENDPROC


	PROCEDURE browsecolumns
		PRIVATE m.lnSelect, m.lcEditColumns

		m.lnSelect=SELECT()
		SELECT WebImpxEdit
		m.lcEditColumns=""

		SCAN FOR WebImpxEdit.lDisp
			* Fieldname
			m.lcEditColumns	=m.lcEditColumns+ALLTRIM(WebImpxEdit.Field_name)
			* Header
				* If NOT Imported, wrap with "()"
			m.lcEditColumns	=m.lcEditColumns+[:H="]+IIF(!WebImpxEdit.lImported,"(","")+ALLTRIM(WebImpxEdit.cHdr)+IIF(!WebImpxEdit.lImported,")","")+["]
			* Edit NOT allowed
			IF !WebImpxEdit.lEdit
				m.lcEditColumns	=m.lcEditColumns+":R"
			ENDIF

			m.lcEditColumns	=m.lcEditColumns+", "
		ENDSCAN

		m.lcEditColumns	=LEFT(m.lcEditColumns, LEN(ALLTRIM(m.lcEditColumns))-1)

		SELECT (m.lnSelect)
		RETURN m.lcEditColumns
	ENDPROC


	PROCEDURE getcsvfile
		ThisForm.txtImportFile.Value   =GETFILE('CSV',"Select import file","Pick Me",0,"Get file...")


		IF FILE("Step.on")
			SET STEP ON 
		ENDIF


		IF !EMPTY(ThisForm.txtImportFile.Value)
			PRIVATE laImportFile
			=ADIR(laImportFile, ThisForm.txtImportFile.Value)
			ThisForm.cImpfileDT=ALLTRIM(DTOC(laImportFile[3]) +" @ " +laImportFile[4])
			ThisForm.cmdDataAssist.Enabled=.T.

			* Create a fresh <user>+"CSV" table and append selected CSV
			SELECT CSVimp

			PRIVATE m.lcSafety
			m.lcSafety	=SET("Safety")
			IF m.lcSafety="ON"
				SET SAFETY OFF
			ENDIF

			COPY STRUCTURE TO (ALLTRIM(ThisForm.txtUser.Value)+"CSV")

			IF m.lcSafety="ON"
				SET SAFETY ON
			ENDIF

			USE (ALLTRIM(ThisForm.txtUser.Value)+"CSV")	ALIAS CSVimp
			APPEND FROM (ThisForm.txtImportFile.Value) CSV

			* Port <user>+"CSV" to Webimp via CSVxWebImp concordence
			m.lnFieldCount	=FCOUNT("CSVimp")

			PRIVATE m.i

			* Scan each CSVimp record
			SCAN	&& CSVimp
				* New WebImp record added to receive current CSVimport record content
				APPEND BLANK IN WebImp
		*		WAIT WINDOW NOWAIT "Getting CSV records:  "+ALLTRIM(STR(RECNO())) +" of " +ALLTRIM(STR(RECCOUNT())) 
				WAIT WINDOW NOWAIT "Getting "+JUSTFNAME(ThisForm.txtImportFile.Value)+" records:  "+ALLTRIM(STR(RECNO())) +" of " +ALLTRIM(STR(RECCOUNT())) 
				REPLACE webimp.cUserName WITH ALLTRIM(ThisForm.txtUser.Value)	&&	oNet.UserName

				* Analyze each field name
				FOR i= 1 TO m.lnFieldCount

					* 1 Find the field in CSVxWebimport
					IF SEEK(FIELD(m.i,"CSVimp"),"CSVxWebImp","cColumn")

						* Scan for each occurence of the field in CSVxWebimp
						SELECT CSVxWebimp
						SCAN WHILE ALLTRIM(cColumn) == ALLTRIM(FIELD(m.i,"CSVimp"))

							SELECT CSVimp
							DO CASE
							* Replace
							CASE CSVxWebimp.cCSVaction="R"
								REPLACE ("Webimp."+ALLTRIM(CSVxWebImp.WebimpFiel)) WITH EVALUATE(CSVxWebimp.cCSVpopula)

							* Error
							OTHERWISE
							ENDCASE

							SELECT CSVxWebimp
						ENDSCAN

						SELECT CSVimp
					ELSE
						* Error
					ENDIF

				NEXT

			ENDSCAN

			WAIT WINDOW NOWAIT "Populating Quotes data..."
			* Port un-edited WebImp records to BML\Quotes.dbf
			SELECT WebImp
			GO TOP
		*	LOCATE REST FOR ThisForm.WebRec2Quotes(laImportFile[1], DTOC(laImportFile[3]) +" @ " +laImportFile[4]) .AND. .F.
			LOCATE REST FOR ThisForm.WebRec2Quotes(laImportFile[1], ThisForm.cImpfileDT) .AND. .F.

		*	* Create purchase data as T data type.
		*	REPLACE WebImp.tPurchase WITH CTOT(STUFF(WebImp.cPurchase,11,1,"T")) FOR !WebImp.lHold .AND. !DELETED()

		*	** remove the SS# where "000000000"
		*	REPLACE ALL WebImp.SSN WITH "" FOR WebImp.SSN ="000000000" .AND. !WebImp.lHold .AND. !DELETED()

			** Extract SSN last-four, populate cSSN4
			*	for Individual
			** Copy the SSN content to cEIN
			*	for Firm
		*	REPLACE ALL WebImp.cSSN4	WITH RIGHT(ALLTRIM(WebImp.SSN),4) FOR !EMPTY(WebImp.Firstname) .AND. !WebImp.lHold .AND. !DELETED()
		*	REPLACE ALL WebImp.cEIN	WITH ALLTRIM(WebImp.SSN) FOR EMPTY(WebImp.Firstname) .AND. !WebImp.lHold .AND. !DELETED()

			* Assure Firm name is in Lastname
			REPLACE ALL WebImp.Lastname	WITH WebImp.Company FOR EMPTY(WebImp.Firstname) .AND. !WebImp.lHold .AND. !DELETED() .AND. ALLTRIM(WebImp.cUsername)=ALLTRIM(ThisForm.txtUser.Value)

			* Set 'Policy Type'
		*	REPLACE ALL WebImp.nPolicyTyp	WITH IIF(!EMPTY(WebImp.Firstname),1,2)	FOR !WebImp.lHold .AND. !DELETED()

			* Place User name in WebImp rile.
		*	REPLACE ALL Webimp.cUsername 	WITH ALLTRIM(ThisForm.TxtUser.Value)	FOR !WebImp.lHold .AND. !DELETED()

			* Assure only cSSN4 .OR. cEIN (not both)
			SCAN FOR !EMPTY(WebImp.cSSN4) .AND. !EMPTY(WebImp.cEIN)
				* If Individual
				IF WebImp.cPolicytyp="Individual"
					REPLACE WebImp.mNewNotes 	WITH [- Removed "]+ALLTRIM(WebImp.cEIN)+[" from EI# (cEIN) field.  (This is an Individual.)]
					REPLACE WebImp.cEIN			WITH SPACE(LEN(WebImp.cEIN))
				ELSE
				* If Firm
					REPLACE WebImp.mNewNotes	WITH [- Removed "]+ALLTRIM(WebImp.cSSN4)+[" from SS4# (cSSN4) field.  (This is a Firm.)]
					REPLACE WebImp.cSSN4		WITH SPACE(LEN(WebImp.cSSN4))
				ENDIF
			ENDSCAN

			* Format license for states:  AK, IA, NM, RI
			PRIVATE m.lnLicCount, m.lcCorrectedLic, m.lcState, m.lcCorrectedLic, m.lcNUMBER
			SCAN 
				m.lnLicCount=ThisForm.GetLicense(WebImp.cPoliLic)
				m.lcCorrectedLic=""

				FOR m.i =1 TO m.lnLicCount
					m.lcState = ThisForm.GetLicense(WebImp.cPoliLic,m.i,"STATE")

		*			IF m.lcState$"AK,NM,RI,IA"
					IF m.lcState$"IA"

						DO CASE
						* Remove any leading zeroes from NUMBER
		*				CASE m.lcState$"AK,NM,RI"
		*					m.lcCorrectedLic	=m.lcCorrectedLic	;
												+ThisForm.GetLicense(WebImp.cPoliLic,m.i,"STATE")+"."	;
												+ThisForm.GetLicense(WebImp.cPoliLic,m.i,"TYPE")+"."	;
												+ALLTRIM(STR(VAL(ThisForm.GetLicense(WebImp.cPoliLic,m.i,"NUMBER"))))	;
												+";"
						* Remove "000" suffix if present
						*	- assure lic# is 5 length by padding front with zeroes.
						CASE m.lcState$"IA"
							m.lcCorrectedLic	=m.lcCorrectedLic	;
												+m.lcState+"."	;
												+ThisForm.GetLicense(WebImp.cPoliLic,m.i,"TYPE")+"."

							m.lcNUMBER			=ThisForm.GetLicense(WebImp.cPoliLic,m.i,"NUMBER")

							IF m.lcNUMBER#"PENDING"
								m.lcNUMBER	=ALLTRIM(STR(VAL(m.lcNUMBER)))
							ENDIF

							IF RIGHT(m.lcNUMBER,3)	="000"
								m.lcNUMBER	=LEFT(m.lcNUMBER, LEN(m.lcNUMBER)-3)
							ENDIF

							m.lcNUMBER			=IIF(!EMPTY(m.lcNUMBER) .AND. ALLTRIM(UPPER(m.lcNUMBER))#"PENDING", PADL(m.lcNUMBER,5,"0"), m.lcNUMBER)
							m.lcCorrectedLic	=m.lcCorrectedLic	;
													+m.lcNUMBER		;
													+";"
						ENDCASE
					ELSE
						m.lcCorrectedLic	=m.lcCorrectedLic	;
												+m.lcState+"."	;
												+ThisForm.GetLicense(WebImp.cPoliLic,m.i,"TYPE")+"."	;
												+ThisForm.GetLicense(WebImp.cPoliLic,m.i,"NUMBER")		;
												+";"

					ENDIF

				NEXT

				IF !EMPTY(m.lcCorrectedLic)
					m.lcCorrectedLic	=LEFT(m.lcCorrectedLic, LEN(m.lcCorrectedLic)-1)
					REPLACE WebImp.cPoliLic	WITH m.lcCorrectedLic
				ENDIF

			ENDSCAN


			GO TOP
			LOCATE FOR !DELETED()
			ThisForm.grdWebImp.Refresh()
		*	ThisForm.grdWebImp.SetFocus
			WAIT CLEAR
		ENDIF
	ENDPROC


	PROCEDURE mtooltiptext
		LPARAMETERS tnLeft, tnTop, tcText

		IF !EMPTY(tcText)
		    WITH THIS
		            IF VARTYPE(.lblToolTipText)    =    [U]
		                .AddObject([lblToolTipText],[label])

		                WITH .lblToolTipText
		                    .AutoSize =    .T.
		                    .BackColor = RGB(255,255,215)
		                    .BorderStyle = 1
		                    .FontName = [MS Sans Serif]
		                ENDWITH
		            ENDIF
		    
		            WITH .lblToolTipText
		                .Caption = " "+tcText+" "
		                .Left = tnLeft
		                .Top = tnTop - 35
		                .Visible = .T.
		        ENDWITH
		    ENDWITH
		ENDIF
	ENDPROC


	PROCEDURE getlicense
		PARAMETERS m.xcLic, m.xnLic, m.xcPart
		PRIVATE m.lxReturn, m.lcLicense, m.lnStart, m.lnLength

		DO CASE
		* Part of license # is returned
		CASE !EMPTY(m.xcPart) .AND. TYPE("m.xcPart")='C'
			m.lcLicense=ThisForm.GetLicense(m.xcLic, m.xnLic)

			DO CASE
			CASE m.xcPart="STATE"
				m.lxReturn=LEFT(m.lcLicense,2)
			CASE m.xcPart="TYPE"
				m.lnStart=AT(".",m.lcLicense,1)+1
				m.lnLength=AT(".",m.lcLicense,2)-m.lnStart
				m.lxReturn=SUBSTR(m.lcLicense, m.lnStart, m.lnLength)
			CASE m.xcPart="NUMBER"
				m.lnStart=AT(".",m.lcLicense,2)+1
				m.lxReturn=SUBSTR(m.lcLicense, m.lnStart)
			OTHERWISE
				m.lxReturn=SPACE(LEN(m.xcLic))
			ENDCASE

		* Full license # is returned
		CASE !EMPTY(m.xnLic) .AND. TYPE("m.xnLic")='N'
			m.lcLicense=ALLTRIM(m.xcLic)

			IF m.xnLic=1
				m.lnStart=1
				m.lnLength=AT(";",m.lcLicense,1)-1

				IF m.lnLength<1
					m.lxReturn=m.lcLicense
				ELSE
					m.lxReturn=SUBSTR(m.lclicense, m.lnstart, m.lnlength)
				ENDIF
			ELSE
				m.lnStart=AT(";",m.lclicense,m.xnlic-1)+1

				IF m.lnStart>1
					m.lnLength= AT(";",m.lcLicense, m.xnLic)-m.lnStart

					IF m.lnLength<1
						m.lxReturn=ALLTRIM(SUBSTR(m.lcLicense,m.lnStart))
					ELSE
						m.lxReturn=SUBSTR(m.lcLicense,m.lnStart,m.lnLength)
					ENDIF
				ELSE
					m.lxReturn=SPACE(LEN(m.xcLic))
				ENDIF

			ENDIF

		* Count of licenses is returned
		OTHERWISE
			m.lcLicense=ALLTRIM(m.xcLic)
			m.lxReturn=IIF(!EMPTY(m.lcLicense),1,0)+OCCURS(";",m.lcLicense)

		ENDCASE

		RETURN m.lxReturn
	ENDPROC


	PROCEDURE Init
		PRIVATE m.lnSelect, m.lnRecno, m.llResume, aVersion
		ThisForm.Width=(SYSMETRIC(1)-50)
		ThisForm.Height=(SYSMETRIC(2)-50)
		=ADIR(aVersion,"webimpcsv.exe")
		ThisForm.Caption=ALLTRIM(ThisForm.Caption) +":  version " +DTOC(aVersion[1,3]) +"@" +aVersion[1,4]
		RELEASE aVersion
		*SET STEP ON 
		m.lnSelect	=SELECT()
		SELECT WebImp
		m.lnRecno	=RECNO()
		LOCATE FOR WebImp.lHold=.F. .AND. !DELETED() .AND. !EOF()
		m.llResume	=FOUND()

		*IF MIN(m.lnRecno, RECCOUNT())>0
		*	GO MIN(m.lnRecno, RECCOUNT())
		*ELSE
		*	GO TOP
		*ENDIF

		IF m.llResume
			ThisForm.shape4.BackColor= RGB(255,0,0)
			ThisForm.label4.BackColor= RGB(255,0,0)
			ThisForm.cmdHoldsOnly.Enabled=.F.
			ThisForm.txtuser.Enabled=.F.
			ThisForm.grdWebImp.clmHold.ReadOnly=.F.
			Thisform.cmdImport.Enabled=.T.
			Thisform.grdWebImp.SetFocus()
			ThisForm.cmdClear.Enabled=.T.
		ELSE
			GO TOP
		ENDIF
	ENDPROC


	PROCEDURE cmddataassist.Click
		This.Enabled= .F.

		IF ThisForm.DataAssist()
			ThisForm.cmdImport.Enabled=.T.
			ThisForm.grdWebImp.clmHold.ReadOnly= .F.
			ThisForm.btnNoteEdit.Enabled=.F.
		ELSE
			This.Enabled=.T.
			ThisForm.grdWebImp.clmHold.ReadOnly= .F.
		ENDIF

		ThisForm.shape3.BackColor= RGB(128,128,128)
		ThisForm.label3.BackColor= RGB(128,128,128)
		ThisForm.shape4.BackColor= RGB(255,0,0)
		ThisForm.label4.backcolor= RGB(255,0,0)
		ThisForm.Refresh()
		ThisForm.grdWebImp.SetFocus()
	ENDPROC


	PROCEDURE cmdgetimportfile.Valid
		ThisForm.shape2.BackColor= RGB(128,128,128)
		ThisForm.label2.BackColor= RGB(128,128,128)
		ThisForm.shape3.BackColor= RGB(255,0,0)
		ThisForm.label3.BackColor= RGB(255,0,0)
		ThisForm.cmdRejects.Init()
		ThisForm.cmdRejects.Refresh()
		ThisForm.cmdClear.Init()
		This.Enabled= .F.
	ENDPROC


	PROCEDURE cmdgetimportfile.Click
		*ThisForm.Getimportfile()
		ThisForm.GetCSVfile()
		ThisForm.btnNoteEdit.Enabled= .T.
		REPLACE NoteTxt.NoteTxt	WITH ""
	ENDPROC


	PROCEDURE cmdimport.Click
		IF FILE("Step.On")
			SET STEP ON 
		ENDIF

		This.Enabled=.F.
		ThisForm.ParseIt()
		ThisForm.Import()
		*ThisForm.Refresh()
		ThisForm.edtNote.Refresh()
		ThisForm.grdWebImp.SetFocus()
		ThisForm.shape4.BackColor= RGB(128,128,128)
		ThisForm.label4.BackColor= RGB(128,128,128)
		ThisForm.cmdClear.Refresh()
		ThisForm.grdWebImp.SetFocus()
	ENDPROC


	PROCEDURE txtuser.When
		ThisForm.shape1.BackColor=RGB(255,0,0)
		ThisForm.label1.BackColor=RGB(255,0,0)
	ENDPROC


	PROCEDURE txtuser.Init
		oNet = CREATEOBJECT('WScript.Network')
		*This.Value	=oNet.UserName
		This.Value=LOWER(LEFT(oNet.Username,1)+SUBSTR(oNet.Username,AT(".",oNet.Username)+1))
		RELEASE oNet
		This.Refresh
	ENDPROC


	PROCEDURE txtuser.Valid
		*ThisForm.SetUserResource(ALLTRIM(This.Value))
		ThisForm.txtDate.Enabled=.T.
		ThisForm.cmdGetImportFile.Enabled=.F.
		ThisForm.cmdImport.Enabled=.F.
		ThisForm.cmdholdsOnly.Enabled=.F.
	ENDPROC


	PROCEDURE txtdate.Init
		This.Value=DATE()
		This.Refresh
	ENDPROC


	PROCEDURE txtdate.Valid
		ThisForm.cmdGetImportFile.Enabled=.T.
		ThisForm.cmdDataAssist.Enabled=.F.
		ThisForm.cmdImport.Enabled=.F.
		ThisForm.shape1.BackColor= RGB(128,128,128)
		ThisForm.label1.BackColor= RGB(128,128,128)
		ThisForm.shape2.BackColor= RGB(255,0,0)
		ThisForm.label2.BackColor= RGB(255,0,0)
	ENDPROC


	PROCEDURE cmdexit.Click
		IF USED("IntDupe")
			USE IN IntDupe
		ENDIF

		SELECT WebImp
		CLEAR EVENTS
		QUIT
	ENDPROC


	PROCEDURE grdwebimp.AfterRowColChange
		LPARAMETERS nColIndex
		ThisForm.edtNote.Refresh
		ThisForm.edtNewNotes.Refresh
		ThisForm.grdPI_SSN.Refresh
		ThisForm.grdPI_Lic.Refresh
	ENDPROC


	PROCEDURE grdwebimp.clmsubpoena.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO POLINUMB   && POLINUMB+POLIID
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO PDANUM   && PDANUM
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO PDANUM   && PDANUM
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO FLNAME   && FIRSTNAME+LASTNAME
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO LFNAME   && LASTNAME+FIRSTNAME
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO COMPANY   && COMPANY
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO ADDRESS1   && ADDRESS1
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO CITY   && CITY
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO STATE   && STATE
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO ZIP   && ZIP
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO PHONE1   && PHONE1
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO EMAIL   && EMAIL
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE hold.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE hold.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO SSEIKEY   && ALLTRIM(CSSN4+CEIN)+ALLTRIM(LEFT(FIRSTNAME,2)+LEFT(LASTNAME,3))+SUBSTR(POLINUMB,11,2)+STR(10000-IIF(LEFT(POLINUMB,2)>"50",VAL("19"+LEFT(POLINUMB,2)),VAL("20"+LEFT(POLINUMB,2))))
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.Click
		SET ORDER TO 
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.DblClick
		SET ORDER TO SSEIKEY   && ALLTRIM(CSSN4+CEIN)+ALLTRIM(LEFT(FIRSTNAME,2)+LEFT(LASTNAME,3))+SUBSTR(POLINUMB,11,2)+STR(10000-IIF(LEFT(POLINUMB,2)>"50",VAL("19"+LEFT(POLINUMB,2)),VAL("20"+LEFT(POLINUMB,2))))
		ThisForm.grdWebImp.Refresh()
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE header1.MouseEnter
		LPARAMETERS nButton, nShift, nXCoord, nYCoord
		THISFORM.mToolTipText(nXCoord,nYCoord,THIS.ToolTipText)
	ENDPROC


	PROCEDURE header1.MouseLeave
		LPARAMETERS nButton, nShift, nXCoord, nYCoord

		WITH THISFORM
		    IF VARTYPE(.lblToolTipText) = [O]
		        .lblToolTipText.Visible = .F.
		    ENDIF     
		ENDW
	ENDPROC


	PROCEDURE cmdholdsonly.Click
		ThisForm.grdWebImp.clmHold.ReadOnly=.F.
		Thisform.txtUser.Enabled=.F.
		Thisform.cmdImport.Enabled=.T.
		Thisform.grdWebImp.SetFocus()
		This.Enabled=.F.

		ThisForm.Refresh()
	ENDPROC


	PROCEDURE cmdholdsonly.Init
		PRIVATE m.lnSelect, m.lnRecno
		m.lnSelect	=SELECT()

		SELECT WebImp
		m.lnRecno	=RECNO()
		LOCATE FOR WebImp.lHold=.T.
		This.Enabled=FOUND()
		This.Visible=This.Enabled
		This.Refresh()

		IF MIN(m.lnRecno, RECCOUNT())>0
			GO MIN(m.lnRecno, RECCOUNT())
		ELSE
			GO TOP
		ENDIF

		SELECT (m.lnSelect)
	ENDPROC


	PROCEDURE cmdrejects.Click
		PRIVATE m.lnSelect, m.lnRecno
		m.lnSelect	=SELECT()

		SELECT Web
		m.lnRecno	=RECNO()
		GO BOTT
		BROWSE TITLE "Rejects:  [Esc]=Exit" LAST FOR Web.Reject=.T. .AND. !DELETED() NOAPPEND 

		IF MIN(m.lnRecno,RECCOUNT())>0
			GO MIN(m.lnRecno,RECCOUNT())
		ENDIF

		SELECT (m.lnSelect)
	ENDPROC


	PROCEDURE cmdrejects.Init
		PRIVATE m.lnSelect, m.lnWebRecno, m.lnCount, m.lcSetDeleted
		m.lnSelect		=SELECT()
		m.lnCount		=0		&&	default to Zero
		m.lcSetDeleted	=SET("Deleted")

		IF m.lcSetDeleted="OFF"
			SET DELETED ON
		ENDIF

		m.lnWebRecno	=RECNO("Web")

		IF SEEK(.T.,"Web","Reject")
			SELECT Web
			SET ORDER TO Reject
			=SEEK(.T.)
			COUNT TO m.lnCount WHILE Web.Reject=.T. 
			SET ORDER TO 

			IF m.lnWebRecno>0
				GO MIN(m.lnWebRecno, RECCOUNT())
			ENDIF

			SELECT (m.lnSelect)
		ENDIF

		IF m.lcSetDeleted="OFF"
			SET DELETED OFF
		ENDIF

		This.Caption="Rejects(" +ALLTRIM(STR(m.lnCount)) +")"
		This.Enabled=(m.lnCount>0)
		This.Refresh()
		PRIVATE m.lnSelect, m.lnWebRecno, m.lnCount, m.lcSetDeleted
		m.lnSelect		=SELECT()
		m.lnWebRecno	=RECNO("Web")
		m.lnCount		=0		&&	default to Zero
		m.lcSetDeleted	=SET("Deleted")

		IF m.lcSetDeleted="OFF"
			SET DELETED ON
		ENDIF

		IF SEEK(.T.,"Web","Reject")
			SELECT Web
			SET ORDER TO Reject
			=SEEK(.T.)
			COUNT TO m.lnCount WHILE Web.Reject=.T. 
			SET ORDER TO 
			GO (m.lnWebRecno)
			SELECT (m.lnSelect)
		ENDIF

		IF m.lcSetDeleted="OFF"
			SET DELETED OFF
		ENDIF

		This.Caption="Rejects(" +ALLTRIM(STR(m.lnCount)) +")"
		This.Enabled=(m.lnCount>0)
		This.Refresh()
	ENDPROC


	PROCEDURE cmdclear.Click
		PRIVATE m.lnSelect
		m.lnSelect	=SELECT()

		SELECT WebImp
		DELETE ALL FOR WebImp.lHold=.F.
		GO TOP

		SELECT (m.lnSelect)
		ThisForm.Refresh()
	ENDPROC


	PROCEDURE cmdclear.Refresh
		PRIVATE m.lnSelect, m.lnRecno
		m.lnSelect	=SELECT()

		SELECT WebImp
		m.lnRecno	=RECNO()
		LOCATE FOR WebImp.lHold=.F. .AND. !DELETED()
		This.Enabled=FOUND()

		IF m.lnRecno <=RECCOUNT()
			GO m.lnRecno
		ELSE
			GO TOP
		ENDIF

		SELECT (m.lnSelect)
	ENDPROC


	PROCEDURE btnnoteedit.Refresh
		This.ToolTipText=ALLTRIM(NoteTxt.NoteTxt)
	ENDPROC


	PROCEDURE btnnoteedit.Click
		ThisForm.NoteEdit()
		This.Refresh()
	ENDPROC


	PROCEDURE check1.Click
		IF NameRef.lLink
			REPLACE WebImp.PoliID	WITH NameRef.PoliID

			* Assure Notes are current:  search PoliInd and if found but the PoliInd Notes differ, update NameRef Notes.
		*	IF SEEK(NameRef.cbKey,"PoliInd","cbkey") .and. ALLTRIM(NameRef.Notes)#ALLTRIM(PoliInd.Notes)
		*		REPLACE NameRef.PoliID	WITH PoliInd.Notes
		*	ENDIF

			* Append the selected's Notes into the import, insert first "GAP" if required.
			PRIVATE m.lcGapNote
			IF WebImp.Effective>NameRef.End
				IF !EMPTY(WebImp.PDANum)
					REPLACE WebImp.PDANum	WITH ALLTRIM(WebImp.PDANum) +":GAP"
				ELSE
					REPLACE WebImp.PDANum	WITH ALLTRIM(WebImp.PDANum) +"GAP"
				ENDIF

				m.lcGapNote	="=======GAP from " +DTOC(NameRef.End) +"=======" +CHR(13)
			ELSE
				m.lcGapNote	=""
			ENDIF

			IF WebImp.PoliNumb = NameRef.PoliNumb
				REPLACE WebImp.PDAnum WITH IIF(!EMPTY(WebImp.PDAnum), ALLTRIM(WebImp.PDAnum)+":", "")+"DUPE"
			ENDIF

		*	REPLACE WebImp.Notes	WITH ALLTRIM(WebImp.Notes) +CHR(13) +m.lcGapNote +ALLTRIM(NameRef.Notes) 
			REPLACE WebImp.mNewNotes	WITH ALLTRIM(WebImp.mNewNotes) +CHR(13) +m.lcGapNote 	&& +ALLTRIM(NameRef.Notes) 
			REPLACE WebImp.Notes		WITH ALLTRIM(NameRef.Notes)

			IF EMPTY(WebImp.cSSN4) .AND. !EMPTY(NameRef.cSSN4)
				REPLACE WebImp.cSSN4	WITH NameRef.cSSN4
			ENDIF

		*	IF !EMPTY(NameRef.OrigDate)
		*		REPLACE WebImp.OrigDate	WITH NameRef.OrigDate
		*	ENDIF

			ThisForm.GrdWebImp.Refresh()
			ThisForm.EdtNote.Refresh()
			ThisForm.EdtNewNotes.Refresh()
		ENDIF


	ENDPROC


	PROCEDURE cmdbrowseedit.Click
		ThisForm.BrowseEdit()
		This.Refresh()
	ENDPROC


	PROCEDURE cmdbrowseedit.Refresh
		This.ToolTipText=ALLTRIM(NoteTxt.NoteTxt)
	ENDPROC


	PROCEDURE lblorder.Refresh
		PRIVATE m.lnSelect
		m.lnSelect	=SELECT()
		SELECT WebImp
		This.Caption="*Order = "+IIF(!EMPTY(ORDER()), ORDER(),"<natural>")+"*"
		This.ToolTipText="Key = "+KEY()
		SELECT (m.lnSelect)
	ENDPROC


ENDDEFINE
*
*-- EndDefine: dataimport
**************************************************
