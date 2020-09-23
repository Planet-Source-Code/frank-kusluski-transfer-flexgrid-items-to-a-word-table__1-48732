VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Transfer FlexGrid to Word Table"
   ClientHeight    =   5175
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4950
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5175
   ScaleWidth      =   4950
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton Command2 
      Caption         =   "&Download cool ActiveX control Word OCX!"
      Height          =   495
      Left            =   1440
      TabIndex        =   3
      Top             =   4560
      Width           =   3375
   End
   Begin VB.CommandButton Command1 
      Caption         =   "&Transfer"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   4560
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   4095
      Left            =   120
      TabIndex        =   0
      Top             =   360
      Width           =   4695
      _ExtentX        =   8281
      _ExtentY        =   7223
      _Version        =   393216
   End
   Begin VB.Label Label1 
      Caption         =   "FlexGrid"
      ForeColor       =   &H8000000D&
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   735
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'for loading my web page
Private Declare Function ShellExecute Lib "shell32.dll" Alias _
        "ShellExecuteA" (ByVal hwnd As Long, ByVal lpOperation _
        As String, ByVal lpFile As String, ByVal lpParameters _
        As String, ByVal lpDirectory As String, ByVal nShowCmd _
        As Long) As Long


Private Sub Command1_Click()
'Early object binding
Dim oWord As Word.Application
Dim oDoc As Word.Document
Dim oRange As Word.Range
'Uncomment below for late object binding
'Dim oWord As Object
'Dim oDoc As Object
'Dim oRange As Object
Dim row As Integer
Dim col As Integer
Dim i As Integer
Dim n As Integer
Dim sTemp As String
Dim arr() As String
  
ReDim arr(MSFlexGrid1.Rows - 1, MSFlexGrid1.Cols - 1)
  
'Create an instance of Word
Set oWord = CreateObject("Word.Application")

'Show Word to the user
oWord.Visible = True

'Add a new, blank document
Set oDoc = oWord.Documents.Add

'Get the current document's range object

'Store FlexGrid items to a two dimensional array
For row = 0 To MSFlexGrid1.Rows - 1
    n = 0
    For col = 0 To MSFlexGrid1.Cols - 1
        arr(i, n) = MSFlexGrid1.TextMatrix(row, col)
        n = n + 1
    Next
    i = i + 1
Next

'Store array items to a string
For i = LBound(arr, 1) To UBound(arr, 1)
    For n = LBound(arr, 2) To UBound(arr, 2)
        sTemp = sTemp & arr(i, n)
        If n = UBound(arr, 2) Then
           sTemp = sTemp & vbCrLf
        Else
           sTemp = sTemp & vbTab
        End If
    Next
Next

'get the current document's range object and move to end of document
Set oRange = oDoc.Bookmarks("\EndOfDoc").Range

oRange.Text = sTemp

'convert the text to a table and format the table
oRange.ConvertToTable vbTab, Format:=wdTableFormatColorful2

Set oRange = Nothing

End Sub

Private Sub Command2_Click()
ShellExecute 0, vbNullString, "http://ic.net/~kusluski", vbNullString, vbNullString, 1
End Sub

Private Sub Form_Load()
Dim i As Integer
Dim n As Integer
'populate FlexGrid
With MSFlexGrid1
     .Rows = 1
     .Cols = 4
     .ColWidth(1) = 1100
     .ColWidth(2) = 1100
     .ColWidth(3) = 1100
     'Add field headers
     For i = 1 To 3
         .col = i
         .Text = "Col " & i
     Next
     'Add data
     For i = 1 To 15
         .Rows = .Rows + 1
         .row = .Rows - 1
         .col = 0
         .Text = "Row " & i
         For n = 1 To 3
             .col = n
             .Text = "Row " & i & ",Col " & n
         Next
     Next
End With

End Sub
