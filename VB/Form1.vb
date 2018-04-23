Imports Microsoft.VisualBasic
Imports System
Imports System.Collections.Generic
Imports System.ComponentModel
Imports System.Data
Imports System.Drawing
Imports System.Text
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDOCVARIABLEBasics
	Partial Public Class Form1
		Inherits Form
		Public Sub New()
			InitializeComponent()

			Dim details As New List(Of DetailInfo)()

			details.Add(New DetailInfo(1, "Detail1"))
			details.Add(New DetailInfo(2, "Detail2"))

			richEditControl1.Options.MailMerge.DataSource = details

			AddHandler richEditControl1.Document.CalculateDocumentVariable, AddressOf Document_CalculateDocumentVariable

			richEditControl1.LoadDocument("Template.rtf")
			ShowFieldCodes()
		End Sub

		Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button1.Click
			richEditControl1.LoadDocument("Template.rtf")
			ShowFieldCodes()
		End Sub

		Private Sub button2_Click(ByVal sender As Object, ByVal e As EventArgs) Handles button2.Click
			Dim myMergeOptions As MailMergeOptions = richEditControl1.Document.CreateMailMergeOptions()
			myMergeOptions.MergeMode = MergeMode.NewParagraph

			Dim server As New RichEditDocumentServer()

			AddHandler server.CalculateDocumentVariable, AddressOf Document_CalculateDocumentVariable
			richEditControl1.Document.MailMerge(myMergeOptions, server.Document)

			richEditControl1.CreateNewDocument()
			richEditControl1.Document.AppendDocumentContent(server.Document.Range)
		End Sub

		Private Sub Document_CalculateDocumentVariable(ByVal sender As Object, ByVal e As DevExpress.XtraRichEdit.CalculateDocumentVariableEventArgs)
			Dim detailId As Integer = -1

			If Int32.TryParse(e.Arguments(0).Value, detailId) Then
				Dim server As New RichEditDocumentServer()

				Dim path As String = System.IO.Directory.GetCurrentDirectory() & "\Detail" & detailId.ToString() & ".rtf"

				server.LoadDocument(path)

				e.Value = server
				e.Handled = True
			End If
		End Sub

		Private Sub ShowFieldCodes()
			Dim doc As Document = richEditControl1.Document
			doc.BeginUpdate()
			For Each f As Field In doc.Fields
				f.ShowCodes = True
			Next f
			doc.EndUpdate()
		End Sub

	End Class

	Public Class DetailInfo
		Private id As Integer

		Public Property DetailId() As Integer
			Get
				Return id
			End Get
			Set(ByVal value As Integer)
				id = value
			End Set
		End Property
		Private description_Renamed As String

		Public Property Description() As String
			Get
				Return description_Renamed
			End Get
			Set(ByVal value As String)
				description_Renamed = value
			End Set
		End Property

		Public Sub New(ByVal id As Integer, ByVal description As String)
			Me.DetailId = id
			Me.Description = description
		End Sub
	End Class

End Namespace