'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************

'Program Name: Copy Sheets Name
'03-10-2017 : new
'16-11-2018 : Revise 
'เป็นโปรแกรมกรอก Part No. and Register No. ให้โดยอัตโนมัติ
'Programmed by Somkid Bothaisong
'Concept 

'***********************************************************************************************************
'***********************************************************************************************************
'***********************************************************************************************************


Public Class GobalVariable
	Public TotalPage As Integer
End Class

Sub Main()
	CopySheetsNmae()
	CopyPartNo()
	ChangeTotalPage()
End Sub

Public TotalPage As New GobalVariable()

Sub CopySheetsNmae()

	'On Error Resume Next

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
		Try
    		oTitleBlock=oSheet.TitleBlock
    		oTextBoxes=oTitleBlock.Definition.Sketch.TextBoxes
			SheetName=oDoc.ActiveSheet.Name
			LResult=Len(SheetName)
			For J = 0 To LResult-1
				If SheetName(J) = ":" Then
					Index_1 = J
				End If	
			Next J
			SheetName = Left(SheetName,Index_1)
			For Each oTextBox In oTitleBlock.Definition.Sketch.TextBoxes
				If oTextBox.Text = "<PART NO.>"  Then      		 
            		oPromptEntry  =  oTitleBlock.GetResultText(oTextBox)      		
					oTitleBlock.SetPromptResultText(oTextBox, SheetName)				
				ElseIf oTextBox.Text = "<Revision>" Then
					oPromptEntry  =  oTitleBlock.GetResultText(oTextBox)         		
					oTitleBlock.SetPromptResultText(oTextBox, "1")
				ElseIf oTextBox.Text = "<Revise page1>" Then
					oPromptEntry  =  oTitleBlock.GetResultText(oTextBox) 		
					oTitleBlock.SetPromptResultText(oTextBox, "-")
				ElseIf oTextBox.Text = "<Detail>" Then
					oPromptEntry  =  oTitleBlock.GetResultText(oTextBox)			
					oTitleBlock.SetPromptResultText(oTextBox, "New Release")
				End If
    			
			Next
		Catch
			MessageBox.Show("Page : " & i & " No TitleBlock", "Title")	
		End Try
	Next
	TotalPage.TotalPage = i
	

End Sub

Sub CopyPartNo()

	'On Error Resume Next

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
		Try
    		oBorder=oSheet.Border
    		oTextBoxes=oBorder.Definition.Sketch.TextBoxes
			SheetName=oDoc.ActiveSheet.Name
			LResult=Len(SheetName)
			For J = 0 To LResult-1
				If SheetName(J) = ":" Then
					Index_1 = J
				End If	
			Next J
			SheetName = Left(SheetName,Index_1)
			For Each oTextBox In oBorder.Definition.Sketch.TextBoxes
				If oTextBox.Text = "<PART NO.>" Then
            		oPromptEntry  =  oBorder.GetResultText(oTextBox)		
					oBorder.SetPromptResultText(oTextBox, SheetName)
    			ElseIf oTextBox.Text = "<PAGE>" Then
					oPromptEntry  =  oBorder.GetResultText(oTextBox)		
					oBorder.SetPromptResultText(oTextBox, i)
				End If
			Next
		Catch
			MessageBox.Show("Page : " & i & " No Border", "Title")	
		End Try
	Next
	
End Sub

Sub ChangeTotalPage()

	'On Error Resume Next

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
		If i = 1 Then
			Try
    			oBorder=oSheet.Border
    			oTextBoxes=oBorder.Definition.Sketch.TextBoxes
			
				For Each oTextBox In oBorder.Definition.Sketch.TextBoxes
					If oTextBox.Text = "<TOTAL PAGE>" Then
            			oPromptEntry  =  oBorder.GetResultText(oTextBox)		
						oBorder.SetPromptResultText(oTextBox, TotalPage.TotalPage)    			
					End If
				Next
			Catch
				
			End Try
		End If
	Next
	

	
End Sub