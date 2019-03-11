Imports System.IO
'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************

'Program Name: ExportSheetName
'26-09-2018
'เป็นโปรแกรมกรอก ChangeRevisionPageOfDrawing ให้โดยอัตโนมัติ
'Programmed by Somkid Bothaisong
'Concept 

'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************



Public Class GobalVariable
	Dim Index As Integer = 10000
	Public RevisionStr(Index) As String
End Class



Public Sub Main()

	'On Error Resume Next	

	Dim oDoc As DrawingDocument 
	oDoc = ThisApplication.ActiveDocument
	Dim oSheet As Sheet

	Dim oPromptEntry
	Dim oCurrentSheet
	oCurrentSheet = oDoc.ActiveSheet.Name

	Dim RegisterNO As String
	Dim DRAWING_BY As String
	Dim TOOL_NAME As String
	Dim Model As String
	Dim PartNo As String
	
	Dim sPath As String

	RegisterNO = iProperties.Value("Custom", "10. REGISTER NO.")	
	DRAWING_BY = iProperties.Value("Custom", "6.DRAWING BY")
	TOOL_NAME = iProperties.Value("Custom", "1.TOOL NAME")
	Model = iProperties.Value("Custom", "2.ORDER")
	Unit = iProperties.Value("Custom", "16.Unit")
	DateRegister = iProperties.Value("Custom", "9. DATE")
	
		
	'processed update when rule is run so save doesn't have to occur to see change
	iLogicVb.UpdateWhenDone = True
	
	sPath = ThisApplication.DesignProjectManager.ActiveDesignProject.WorkspacePath	
	sPath = sPath & "\" & TOOL_NAME & "\" & RegisterNO & "_" & Model & "_" & TOOL_NAME & Unit
	If Len(FileSystem.Dir(sPath, vbDirectory)) = 0 Then
		FileSystem.MkDir(sPath)	
	End If

	Dim fileName As String  = sPath & "\" & RegisterNO & "-MEDWG" & Unit & ".TXT"
	Dim wr1 As StreamWriter = New StreamWriter(fileName)
	
	i = 0
	
	For Each oSheet In oDoc.Sheets
		i = i+1
		ThisApplication.ActiveDocument.Sheets.Item(i).Activate   		
		SheetName=oDoc.ActiveSheet.Name
		LResult=Len(SheetName)		
		For J = 0 To LResult-1
			If SheetName(J) = ":" Then				
				Index_1 = J
			End If	
		Next J
		SheetName = Left(SheetName, Index_1)
		PartNo = SheetName	
		
		If SheetName = "DWG First Page" Then
			SheetName = RegisterNO & "-MEDWG" & Unit
			wr1.WriteLine(SheetName)
			wr1.WriteLine(Model)
			wr1.WriteLine(TOOL_NAME)
			wr1.WriteLine(SheetName)
			wr1.WriteLine(DRAWING_BY)			
			wr1.WriteLine("0")
			wr1.WriteLine(DateRegister)
			
		
		
		Else
			wr1.WriteLine(RegisterNO & "-MEDWG-" & SheetName)
			wr1.WriteLine(Model)
			wr1.WriteLine(TOOL_NAME)
			wr1.WriteLine(PartNo)
			wr1.WriteLine(DRAWING_BY)			
			wr1.WriteLine("0")
			wr1.WriteLine(DateRegister)
		End If

	Next

	wr1.Close()
	MessageBox.Show(fileName, "Title")
End Sub


Public GVariable As New GobalVariable


Function RevisionFirstPage()	
	Dim RevisionFirstPageInt As String = "0"	
	Dim oDoc As DrawingDocument													
	oDoc = ThisApplication.ActiveDocument	
	ThisApplication.ActiveDocument.Sheets.Item(1).Activate
	Dim oSheet As Sheet															
	oSheet = oDoc.ActiveSheet													
	oBorder = oSheet.Border														
	oTextBoxes = oBorder.Definition.Sketch.TextBoxes
	Dim RevisionRaw As String	
	Dim Index As Integer = 0
	Dim TextName As String
	For Index = 1 To 12
		TextName = "<R" & CStr(Index) & "Revision>"		
		For Each oTextBox In oBorder.Definition.Sketch.TextBoxes
			SText = oTextBox.Text			
			If SText = TextName Then
				RevisionRaw = oBorder.GetResultText(oTextBox)
				If RevisionRaw = " " Then					
				Else
					RevisionFirstPageInt = RevisionRaw
					GVariable.RevisionStr(0) = RevisionFirstPageInt
				End If
					
			End If
		Next
	Next Index	

End Function

Function RevisionEachPage()
	Dim RevisionEachPageStr As String = "0"
	Dim oDoc As DrawingDocument
	oDoc = ThisApplication.ActiveDocument
	Dim oSheet As Sheet

	Dim oPromptEntry
	Dim oCurrentSheet
	oCurrentSheet = oDoc.ActiveSheet.Name


	i = 0
	For Each oSheet In oDoc.Sheets
		i = i+1
		ThisApplication.ActiveDocument.Sheets.Item(i).Activate
		If i > 1
    		oTitleBlock=oSheet.TitleBlock
    		oTextBoxes=oTitleBlock.Definition.Sketch.TextBoxes
			For Each oTextBox In oTitleBlock.Definition.Sketch.TextBoxes
				name=oTextBox.Text
				Select oTextBox.Text
        		Case "<Revision>"
            		RevisionEachPageStr = oTitleBlock.GetResultText(oTextBox)
					GVariable.RevisionStr(i)= RevisionEachPageStr
    			End Select
			Next
		End If

	Next

	
End Function


