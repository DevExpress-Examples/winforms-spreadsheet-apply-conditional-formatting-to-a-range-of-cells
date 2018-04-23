Imports Microsoft.VisualBasic
Imports System
Imports System.Windows.Forms
Imports DevExpress.Spreadsheet
Imports System.Diagnostics

Namespace ConditionalFormatting_Examples
	Partial Public Class Form1
		Inherits Form

		Private workbook As IWorkbook

		Public Sub New()
			InitializeComponent()

			'  Access a workbook.
			workbook = spreadsheetControl1.Document
			InitTreeListControl()

		End Sub

		Private Sub InitTreeListControl()
			Dim examples As New GroupsOfSpreadsheetExamples()
			InitData(examples)
			DataBinding(examples)
		End Sub

		Private Sub InitData(ByVal examples As GroupsOfSpreadsheetExamples)
'			#Region "GroupNodes"
			examples.Add(New SpreadsheetNode("Conditional Formatting Examples"))
'			#End Region

'			#Region "ExampleNodes"
			' Add nodes to the "Conditional Formatting" group of examples.
			examples(0).Groups.Add(New SpreadsheetExample("Format values that are above or below average", ConditionalFormatting.AddAverageConditionalFormattingAction))
			examples(0).Groups.Add(New SpreadsheetExample("Format cells that are not between two specified values", ConditionalFormatting.AddRangeConditionalFormattingAction))
			examples(0).Groups.Add(New SpreadsheetExample("Format top ranked values", ConditionalFormatting.AddRankConditionalFormattingAction))
			examples(0).Groups.Add(New SpreadsheetExample("Format cells that contain the given text", ConditionalFormatting.AddTextConditionalFormattingAction))
			examples(0).Groups.Add(New SpreadsheetExample("Format only unique values", ConditionalFormatting.AddSpecialConditionalFormattingAction))
			examples(0).Groups.Add(New SpreadsheetExample("Format today's date", ConditionalFormatting.AddTimePeriodConditionalFormattingAction))
			examples(0).Groups.Add(New SpreadsheetExample("Format values that are greater than a specified value", ConditionalFormatting.AddExpressionConditionalFormattingAction))
			examples(0).Groups.Add(New SpreadsheetExample("Use a formula to determine which cells to format", ConditionalFormatting.AddFormulaExpressionConditionalFormattingAction))
			examples(0).Groups.Add(New SpreadsheetExample("Format cells by using a two-color scale", ConditionalFormatting.AddColorScale2ConditionalFormattingAction))
			examples(0).Groups.Add(New SpreadsheetExample("Format cells by using a three-color scale", ConditionalFormatting.AddColorScale3ConditionalFormattingAction))
			examples(0).Groups.Add(New SpreadsheetExample("Format cells by using data bars", ConditionalFormatting.AddDataBarConditionalFormattingAction))
            examples(0).Groups.Add(New SpreadsheetExample("Format cells by using a custom icon set", ConditionalFormatting.AddIconSetConditionalFormattingAction))
            examples(0).Groups.Add(New SpreadsheetExample("Apply conditional formatting to a complex range", ConditionalFormatting.AddComplexRangeConditionalFormattingAction))
'			#End Region
		End Sub

		Private Sub DataBinding(ByVal examples As GroupsOfSpreadsheetExamples)
			treeList1.DataSource = examples
			treeList1.ExpandAll()
			treeList1.BestFitColumns()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			LoadDocumentFromFile()
			Dim example As SpreadsheetExample = TryCast(treeList1.GetDataRecordByNode(treeList1.FocusedNode), SpreadsheetExample)
			If example Is Nothing Then
				Return
			End If
			Dim action As Action(Of IWorkbook) = example.Action
			action(workbook)
			SaveDocumentToFile()
		End Sub

		' ------------------- Load and Save a Document -------------------
		Private Sub LoadDocumentFromFile()
'			#Region "#LoadDocumentFromFile"
			' Load a workbook from a file.
			workbook.LoadDocument("Documents\Document.xlsx", DocumentFormat.OpenXml)
'			#End Region ' #LoadDocumentFromFile
		End Sub

		Private Sub SaveDocumentToFile()
'			#Region "#SaveDocumentToFile"
			' Save the modified document to a file.
			workbook.SaveDocument("Documents\SavedDocument.xlsx", DocumentFormat.OpenXml)
'			#End Region ' #SaveDocumentToFile
		End Sub
	End Class
End Namespace
