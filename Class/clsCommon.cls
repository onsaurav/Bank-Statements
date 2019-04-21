VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "clsCommon"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Title          :   clsCommon
'Author         :   Saurav Biswas Kartik
'URL/Mail       :   onsaurav@yahoo.com/gmail.com/hotmail.com
'Description    :   Class Module used to perform the most common options.
'Created        :   Saurav Biswas / Oct-25-2010
'Modified       :   Saurav Biswas /

Public Function DBConnection(cn As ADODB.Connection) As Boolean
       'Summary    :   This function will use for connecting the database
       'Created    :   Saurav Biswas / Oct-25-2010
       'Modified   :   Saurav Biswas /
       'Parameters :   An ADODB connection used as parameter what will be & _
                       connected by this function.
                       
       On Error GoTo Ext
       With cn
            'if the connection is already on then off the connection
            If .State = 1 Then .Close
            'Connection string with the database name supplied. Database name get from the Function DBName
            .ConnectionString = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=  '" & DBName & "';Mode=ReadWrite|Share Deny None;Persist Security Info=False;"
            .Open
       End With
       'if connection i open then return true else false
       If cn.State = 1 Then DBConnection = True Else DBConnection = False
       Exit Function
Ext:
       'if an error occured then show the error message and return false
       MsgBox "Error No : " & Err.Number & Chr(13) & "Error : " & Err.Description, vbInformation, "Database Connection ...": Err.Clear
       DBConnection = False
End Function

Public Function SQL(strSQL As String, rs As Recordset) As Recordset
       'Summary    :   This function will use to open a recordset
       'Created    :   Saurav Biswas / Oct-25-2010
       'Modified   :   Saurav Biswas /
       'Parameters :   Two parameters, first one strSQL is a string used to supply the & _
                       SQL Querry and other one is ADODB recordset what will be opened by this function.
       
       If rs.State = 1 Then rs.Close
       rs.ActiveConnection = cOnn
       rs.CursorLocation = adUseClient
       rs.CursorType = adOpenDynamic
       rs.LockType = adLockOptimistic
       rs.Source = strSQL
       Debug.Print strSQL
       rs.Open
       Set SQL = rs
End Function

Public Function DBName() As String
       'Summary    :   This function will use to get the database path
       'Created    :   Saurav Biswas / Oct-25-2010
       'Modified   :   Saurav Biswas /
       'Parameters :
       
       On Error GoTo Ext
       'String for holding the database name from the file
       Dim DBFileName As String
       'Reade the database full path and name from the file
       Open App.Path & "\DBName.ons" For Input As #1
       Input #1, DBFileName
       'Set the database name as the function return
       DBName = App.Path & "\" & DBFileName
       Close #1
       Exit Function
Ext:
       'if an error occured then show the error message and return ""(blank space)
       MsgBox "Error No : " & Err.Number & Chr(13) & "Error : " & Err.Description, vbInformation, "Database file ...": Err.Clear
       DBName = ""
End Function

Public Function IsAlphaNumeric(CharCode As Integer) As Boolean
       'Summary    :   Function used to check (CharCode) is alfa/numeric
       'Created    :   Saurav Biswas / Oct-25-2010
       'Modified   :   Saurav Biswas /
       'Parameters :
       
       On Error Resume Next
       If (CharCode >= 65 And CharCode <= 90) Then
          IsAlphaNumeric = True
       ElseIf (CharCode >= 97 And CharCode <= 122) Then
          IsAlphaNumeric = True
       ElseIf (CharCode >= 48 And CharCode <= 57) Then
          IsAlphaNumeric = True
       Else
          IsAlphaNumeric = False
       End If
End Function

Public Function RandString(strLen As Integer) As String
       'Summary    :   Generate the Random String
       'Created    :   Saurav Biswas / Oct-25-2010
       'Modified   :   Saurav Biswas /
       'Parameters :
       
       On Error Resume Next
       Dim xChar As Integer
       Randomize
       While strLen <> 0
             xChar = Rnd * 255
             If IsAlphaNumeric(xChar) = True Then
                RandString = RandString & Chr(xChar)
                strLen = strLen - 1
             End If
       Wend
 End Function
 
 Public Function LoadCombo(strSQL As String, cbo As ComboBox) As String
       'Summary    :   This function will use to load a combo box with the specific querry
       'Created    :   Saurav Biswas / Oct-25-2010
       'Modified   :   Saurav Biswas /
       'Parameters :   one parameter is sql querry string, another one is the combo box
       
       Dim rsAutoId As New ADODB.Recordset
       
       cbo.Clear
       Set rsAutoId = ONSCommon.SQL(strSQL, rsAutoId)
       While Not rsAutoId.EOF And Not rsAutoId.BOF
             cbo.AddItem (rsAutoId.Fields(0).Value)
             rsAutoId.MoveNext
       Wend
End Function
