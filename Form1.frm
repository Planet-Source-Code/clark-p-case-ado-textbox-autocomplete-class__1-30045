VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   915
   ClientLeft      =   60
   ClientTop       =   345
   ClientWidth     =   2880
   LinkTopic       =   "Form1"
   ScaleHeight     =   915
   ScaleWidth      =   2880
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text2 
      Height          =   285
      Left            =   960
      TabIndex        =   1
      Top             =   480
      Width           =   1815
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   960
      TabIndex        =   0
      Top             =   120
      Width           =   1815
   End
   Begin VB.Label Label2 
      Caption         =   "Months"
      Height          =   255
      Left            =   240
      TabIndex        =   3
      Top             =   480
      Width           =   615
   End
   Begin VB.Label Label1 
      Caption         =   "Days"
      Height          =   255
      Left            =   240
      TabIndex        =   2
      Top             =   120
      Width           =   495
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim AutoComplete1 As New AutoComplete
Dim AutoComplete2 As New AutoComplete

Private Sub Form_Load()
Dim AppPath As String

AppPath = App.Path
If Right$(AppPath, 1) <> "\" Then AppPath = AppPath & "\"

With AutoComplete1
    .DatabasePathName = AppPath & "AutoComplete.mdb"
    .TableName = "Days"
    .FieldName = "DayNames"
    Set .theTextbox = Text1
    .Init
End With


With AutoComplete2
    .DatabasePathName = AppPath & "AutoComplete.mdb"
    .TableName = "Months"
    .FieldName = "MonthNames"
    Set .theTextbox = Text2
    .Init
End With

End Sub

Private Sub Text1_Change()
    AutoComplete1.AutoComplete
End Sub

Private Sub Text1_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoComplete1.CheckKey KeyCode
End Sub

Private Sub Text2_Change()
    AutoComplete2.AutoComplete
End Sub

Private Sub Text2_KeyDown(KeyCode As Integer, Shift As Integer)
    AutoComplete2.CheckKey KeyCode
End Sub

