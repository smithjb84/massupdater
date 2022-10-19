' TODO: This module exists as a convenient location for the code that does the real
'       work when a command is executed.  If you're converting VBA macros into add-in 
'       commands you can copy the macros here, make changes to make them VB.NET compatible, 
'       and change any references to "ThisApplication" to "g_inventorApplication".

Public Module CommandFunctions
	' Function that's called when the button is clicked.
	Public Sub SampleCommandFunction()
		Dim oDoc As Inventor.Document
		oDoc = g_inventorApplication.ActiveDocument

		If oDoc.DocumentType = Inventor.DocumentTypeEnum.kDrawingDocumentObject Then

			Dim oSheet As Inventor.Sheet
			oSheet = oDoc.ActiveSheet

			Dim oDrawingView As Inventor.DrawingView
			Dim oModel As Inventor.Document
			Dim mass As Double

			If oSheet.DrawingViews.Count <> 0 Then
				oDrawingView = oSheet.DrawingViews.Item(1)
				oModel = oDrawingView.ReferencedDocumentDescriptor.ReferencedDocument
				mass = oModel.ComponentDefinition.MassProperties.Mass
				oDoc.Update()
			End If

		End If
	End Sub
End Module
