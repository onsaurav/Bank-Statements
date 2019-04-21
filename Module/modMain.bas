Attribute VB_Name = "modMain"
'Title          :   modMain
'Author         :   Saurav Biswas Kartik
'URL/Mail       :   onsaurav@yahoo.com/gmail.com/hotmail.com
'Description    :   Main module for the global function and methods
'Created        :   Saurav Biswas / Oct-25-2010
'Modified       :   Saurav Biswas /

'Declare the instance of the common class
Public ONSCommon As New clsCommon

'Global ADODB connection declared
Public cOnn As New ADODB.Connection

Public Sub Main()
       'Perform the database connection if true then open the project (Main form)
       If ONSCommon.DBConnection(cOnn) = False Then End
       frmMenu.Show
End Sub

Public Function GenerateAutoId(Initial As String, TableName As String, FieldName As String) As String
       'Summary    :   This function will use to generate teh Auto Id
       'Created    :   Saurav Biswas / Oct-25-2010
       'Modified   :   Saurav Biswas /
       'Parameters :   Three string type parameter. Initial to submit the id initial and the tablename and field name.
       
       Dim rsAutoId As New ADODB.Recordset
       Set rsAutoId = ONSCommon.SQL("SELECT DISTINCT " & FieldName & " FROM " & TableName & " ORDER BY " & FieldName & " DESC", rsAutoId)
       If rsAutoId.RecordCount > 0 Then
          GenerateAutoId = Initial & "-" & Year(Date) & "-" & Month(Date) & "-" & Format(Right(rsAutoId(FieldName), 5) + 1, "00000")
       Else
          GenerateAutoId = Initial & "-" & Year(Date) & "-" & Month(Date) & "-" & "00001"
       End If
End Function
