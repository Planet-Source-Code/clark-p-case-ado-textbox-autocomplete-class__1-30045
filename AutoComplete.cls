VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "AutoComplete"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Attribute VB_Description = "Makes an SSCombo's text portion act like an IE style Autocomplete box."
Attribute VB_Ext_KEY = "SavedWithClassBuilder" ,"Yes"
Attribute VB_Ext_KEY = "Member0" ,"AutoComplete"
Attribute VB_Ext_KEY = "Top_Level" ,"Yes"
'********************************************************************************************
'This Class adds autocomplete functionality to textboxes, you start typing
'and the textbox is completed with the rest of the word, just like IE 4.0
'This Class does the search in a table, not from a listbox or other things.
'you will need to do 5 things to use this class:
'1. Choose the textbox to add autocomplete to
'2. Declare an instance of the class in a form's Declaration section
'3. In the Form Load event, set the DatabasePathName, Tablename, FieldName and theTextBox properties
'4. In the OnKeyDown event of the textbox, call the CheckKey method
'5. In the Change Event of the textbox, call the AutoComplete method
'If you have more than one autocomplete textbox on the same form, declare
'multiple instances of the class
'Written by Robert Rowe, Coal Software & Systems, robert_rowe@yahoo.com
'Based on original code from: MauTheMan, mautheman@yahoo.com
'********************************************************************************************
Option Explicit
'local variable(s) to hold property value(s)
Private mvarDatabasePathName As String
Private mvarTableName As String
Private mvarFieldName As String
Private mvartheTextbox As Object
Dim my_db As New ADODB.Connection
Dim my_tb As New ADODB.Recordset
Dim IsDelOrBack As Boolean

Public Function AutoComplete() As Boolean
Dim OldLen As Integer
'On Error GoTo eh

AutoComplete = False



If theTextbox <> "" And IsDelOrBack = False Then
    OldLen = Len(theTextbox)
    my_tb.Find FieldName & " like " & Quote(theTextbox & "*"), , adSearchForward
    
    If Not IsNull(my_tb) Then
            If my_tb.BOF = True Or my_tb.EOF = True Then
                my_tb.MoveFirst
                Exit Function
            Else
                theTextbox = my_tb(FieldName) & ""
            End If
            
        If theTextbox.SelText = "" Then
            theTextbox.SelStart = OldLen
        Else
            theTextbox.SelStart = InStr(theTextbox, theTextbox.SelText)
        End If

        theTextbox.SelLength = Len(theTextbox)
        AutoComplete = True
    End If
End If

Exit Function
eh:
    MsgBox "Unexpected Error #" & Err.Number & " " & Err.Description, vbExclamation, "Error!"
    Resume Next
End Function

Public Function CheckKey(TheKey As Integer) As Boolean
On Error GoTo eh
'TheKey is the KeyCode - all you gotta do is write CheckIsDelOrBack(KeyCode) on the KeyDown event...
    If TheKey = vbKeyBack Or TheKey = vbKeyDelete Then
        IsDelOrBack = True
        CheckKey = True
    Else
        IsDelOrBack = False
        CheckKey = False
    End If
Exit Function


eh:
    MsgBox "Unexpected Error #" & Err.Number & " " & Err.Description, vbExclamation, "Error!"
    Resume Next
End Function

Public Property Set theTextbox(ByVal vData As Object)
'used when assigning an Object to the property, on the left side of a Set statement.
'Syntax: Set x.theTextbox = Form1
    Set mvartheTextbox = vData
End Property


Public Property Get theTextbox() As Object
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.theTextbox
    Set theTextbox = mvartheTextbox
End Property

Public Sub Init()
On Error GoTo eh
Dim SQL As String

my_db.Open "Provider = Microsoft.Jet.OLEDB.4.0;" & _
            "Data Source = " & DatabasePathName
            
SQL = "SELECT DISTINCT " & FieldName & " from " & TableName

my_tb.Open SQL, my_db, adOpenStatic

'Set my_db = DBEngine.Workspaces(0).OpenDatabase(DatabasePathName)
'SQL = "Select Distinct " & FieldName & " from " & TableName
'Set my_tb = my_db.OpenRecordset(SQL, dbOpenSnapshot)
    
Exit Sub
eh:
    MsgBox "Unexpected Error #" & Err.Number & " " & Err.Description, vbExclamation, "Error!"
    Resume Next
End Sub

Private Function ReplaceText(Text As String, charfind As String, charchange As String) As String
Dim replace As Integer
Dim thechar, thechars

'Replaces charfind with charchange in the string text
'Example: text1 = ReplaceText(text1, "a", "A")
'This will replace any "a" in Text1 with "A"
    
    If InStr(Text, charfind) = 0 Then
    ReplaceText = Text
    Exit Function
End If

For replace = 1 To Len(Text)
    thechar = Mid(Text, replace, 1)
    thechars = thechars & thechar
    If thechar = charfind Then
        thechars = Mid(thechars, 1, Len(thechars) - 1) + charchange
    End If
Next replace

ReplaceText = thechars

End Function
Private Function Quote(TheText As String) As String
     Quote = ReplaceText(TheText, "'", "''")
     Quote = "'" & Quote & "'"
End Function

Public Property Let FieldName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.FieldName = 5
    mvarFieldName = vData
End Property

Public Property Get FieldName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.FieldName
    FieldName = mvarFieldName
End Property

Public Property Let TableName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.TableName = 5
    mvarTableName = vData
End Property

Public Property Get TableName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.TableName
    TableName = mvarTableName
End Property

Public Property Let DatabasePathName(ByVal vData As String)
'used when assigning a value to the property, on the left side of an assignment.
'Syntax: X.DatabasePathName = 5
    mvarDatabasePathName = vData
End Property

Public Property Get DatabasePathName() As String
'used when retrieving value of a property, on the right side of an assignment.
'Syntax: Debug.Print X.DatabasePathName
    DatabasePathName = mvarDatabasePathName
End Property

