'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************

'Program Name: Export PDF
'09-10-2018
'เป็นโปรแกรมกรอก ChangeRevisionPageOfDrawing ให้โดยอัตโนมัติ
'Programmed by Somkid Bothaisong
'Concept 

'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************
Public Sub Main()
	'On Error Resume Next
	Call ExportPDF()
	MessageBox.Show("Complete...", "Complete...")
End Sub

Public Sub ExportPDF()
	
	'On Error Resume Next

	Dim oDoc As DrawingDocument
        oDoc = ThisApplication.ActiveDocument
    
    Dim oTitleBlock As Inventor.TitleBlock
    Dim oTextBox As Inventor.TextBox
    Dim oSheet As Sheet
    
    Dim lPos As Long
    Dim sSheetName As String
    
    Dim SheetIndex As Integer
    SheetIndex = 1
    
    For Each oSheet In oDoc.Sheets
        lPos = InStr(oSheet.Name, ":")
        sSheetName = Left(oSheet.Name, lPos - 1)		
		PublishPDF(sSheetName, SheetIndex)			
        SheetIndex = SheetIndex + 1
    Next
End Sub

Public Sub PublishPDF(ByVal SheetName As String, ByVal Index As Integer)

	'On Error Resume Next
    ' Get the PDF translator Add-In.
    Dim PDFAddIn As TranslatorAddIn
        PDFAddIn = ThisApplication.ApplicationAddIns.ItemById("{0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4}") '{2B4DB491-D7A7-46E8-89EA-601FEB825999}  0AC6FD96-2F4D-42CE-8BE0-8AEA580399E4 

    'Set a reference to the active document (the document to be published).
    Dim oDocument As Inventor.Document
		oDocument = ThisApplication.ActiveDocument 
        ThisApplication.ActiveDocument.Sheets.Item(Index).Activate
    Dim oContext As TranslationContext
        oContext = ThisApplication.TransientObjects.CreateTranslationContext
    oContext.Type = kFileBrowseIOMechanism
    ' Create a NameValueMap object
    Dim oOptions As NameValueMap
        oOptions = ThisApplication.TransientObjects.CreateNameValueMap
    ' Create a DataMedium object
    Dim oDataMedium As DataMedium
        oDataMedium = ThisApplication.TransientObjects.CreateDataMedium
    ' Check whether the translator has 'SaveCopyAs' options
	
    If PDFAddIn.HasSaveCopyAsOptions(oDocument, oContext, oOptions) Then
        ' Options for drawings...
        oOptions.Value("All_Color_AS_Black") = 0
        'oOptions.Value("Remove_Line_Weights") = 0
        oOptions.Value("Vector_Resolution") = 400
        oOptions.Value("Sheet_Range") = kPrintSheetRange
        oOptions.Value("Custom_Begin_Sheet") = Index
        oOptions.Value("Custom_End_Sheet") = Index
    End If
	
	oCustomPropertySet = ThisDoc.Document.PropertySets.Item("Inventor User Defined Properties")
	RegisterNO = iProperties.Value("Custom", "10. REGISTER NO.")
	Unit = iProperties.Value("Custom", "16.Unit")
	TOOL_NAME = iProperties.Value("Custom", "1.TOOL NAME")
	Model = iProperties.Value("Custom", "2.ORDER")
	
	Dim sPath As String
	sPath = ThisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath
	sPath = sPath  & "\" & TOOL_NAME & "\" & RegisterNO & "_" & Model & "_" & TOOL_NAME & Unit	
	If Len(FileSystem.Dir(sPath, vbDirectory)) = 0 Then
		FileSystem.MkDir(sPath)		
	End If
	
    'Set the destination file name
	Dim FileNameText As String
	If SheetName = "DWG First Page" Then
		If Unit = "" Then
			FileNameText = RegisterNO & "-MEDWG" & ".PDF"
		Else
			FileNameText = RegisterNO & "-MEDWG" & Unit & ".PDF"
		End If
	Else
		FileNameText = RegisterNO & "-MEDWG-" & SheetName & ".PDF"
	End If	
    oDataMedium.FileName = sPath & "\" & FileNameText
    Call PDFAddIn.SaveCopyAs(oDocument, oContext, oOptions, oDataMedium)
End Sub