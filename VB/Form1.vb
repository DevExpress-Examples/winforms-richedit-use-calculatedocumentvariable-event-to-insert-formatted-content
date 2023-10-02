Imports System
Imports System.Collections.Generic
Imports System.Windows.Forms
Imports DevExpress.XtraRichEdit.API.Native
Imports DevExpress.XtraRichEdit

Namespace RichEditDOCVARIABLEBasics

    Public Partial Class Form1
        Inherits Form

        Public Sub New()
            InitializeComponent()
            Dim details As List(Of DetailInfo) = New List(Of DetailInfo)()
            details.Add(New DetailInfo(1, "Detail1"))
            details.Add(New DetailInfo(2, "Detail2"))
            richEditControl1.Options.MailMerge.DataSource = details
            AddHandler richEditControl1.Document.CalculateDocumentVariable, New CalculateDocumentVariableEventHandler(AddressOf Document_CalculateDocumentVariable)
            richEditControl1.LoadDocument("Template.rtf")
            ShowFieldCodes()
        End Sub

        Private Sub button1_Click(ByVal sender As Object, ByVal e As EventArgs)
            richEditControl1.LoadDocument("Template.rtf")
            ShowFieldCodes()
        End Sub

        Private Sub button2_Click(ByVal sender As Object, ByVal e As EventArgs)
            Dim myMergeOptions As MailMergeOptions = richEditControl1.Document.CreateMailMergeOptions()
            myMergeOptions.MergeMode = MergeMode.NewParagraph
            Dim server As RichEditDocumentServer = New RichEditDocumentServer()
            AddHandler server.CalculateDocumentVariable, New CalculateDocumentVariableEventHandler(AddressOf Document_CalculateDocumentVariable)
            richEditControl1.Document.MailMerge(myMergeOptions, server.Document)
            richEditControl1.CreateNewDocument()
            richEditControl1.Document.AppendDocumentContent(server.Document.Range)
        End Sub

        Private Sub Document_CalculateDocumentVariable(ByVal sender As Object, ByVal e As CalculateDocumentVariableEventArgs)
            Dim detailId As Integer = -1
            If Integer.TryParse(e.Arguments(0).Value, detailId) Then
                Dim server As RichEditDocumentServer = New RichEditDocumentServer()
                Dim path As String = String.Format("{0}\Detail{1}.rtf", IO.Directory.GetCurrentDirectory(), detailId.ToString())
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
            Next

            doc.EndUpdate()
        End Sub
    End Class

    Public Class DetailInfo

        Private id As Integer

        Public Property DetailId As Integer
            Get
                Return id
            End Get

            Set(ByVal value As Integer)
                id = value
            End Set
        End Property

        Private descriptionField As String

        Public Property Description As String
            Get
                Return descriptionField
            End Get

            Set(ByVal value As String)
                descriptionField = value
            End Set
        End Property

        Public Sub New(ByVal id As Integer, ByVal description As String)
            DetailId = id
            Me.Description = description
        End Sub
    End Class
End Namespace
