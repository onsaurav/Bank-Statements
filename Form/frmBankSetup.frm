VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmBankSetup 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bank Setup"
   ClientHeight    =   6255
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12135
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6255
   ScaleWidth      =   12135
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4455
      Left            =   0
      TabIndex        =   13
      Top             =   840
      Width           =   12135
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   325
         Left            =   4680
         TabIndex        =   10
         Top             =   4000
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   325
         Left            =   3600
         TabIndex        =   9
         Top             =   4000
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   325
         Left            =   2520
         TabIndex        =   8
         Top             =   4000
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   325
         Left            =   1440
         TabIndex        =   7
         Top             =   4000
         Width           =   1095
      End
      Begin VB.TextBox txtFax 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   4
         Top             =   2400
         Width           =   4335
      End
      Begin VB.TextBox txtURL 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   6
         Top             =   3120
         Width           =   4335
      End
      Begin VB.TextBox txtEmail 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   5
         Top             =   2760
         Width           =   4335
      End
      Begin VB.TextBox txtPhone1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   3
         Top             =   2040
         Width           =   4335
      End
      Begin VB.TextBox txtAddress 
         Appearance      =   0  'Flat
         Height          =   885
         Left            =   1440
         TabIndex        =   2
         Top             =   1080
         Width           =   4335
      End
      Begin VB.TextBox txtBankName 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1440
         TabIndex        =   1
         Top             =   720
         Width           =   4335
      End
      Begin VB.TextBox txtBankId 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1440
         TabIndex        =   0
         Top             =   360
         Visible         =   0   'False
         Width           =   1695
      End
      Begin MSComctlLib.ListView lvwBank 
         Height          =   4095
         Left            =   6120
         TabIndex        =   14
         Top             =   240
         Width           =   5775
         _ExtentX        =   10186
         _ExtentY        =   7223
         View            =   3
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   1
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Bank Name"
            Object.Width           =   9569
         EndProperty
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "URL"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   20
         Top             =   3120
         Width           =   1095
      End
      Begin VB.Label Label6 
         BackStyle       =   0  'Transparent
         Caption         =   "Email"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   19
         Top             =   2760
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "FAX"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   18
         Top             =   2400
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Phone"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   17
         Top             =   2040
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Address"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   16
         Top             =   1080
         Width           =   1095
      End
      Begin VB.Label Label2 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Name"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   15
         Top             =   720
         Width           =   1095
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E6C495&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   12
      Top             =   5400
      Width           =   12135
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6C495&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   11
      Top             =   0
      Width           =   12135
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Setup"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   18
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   120
         TabIndex        =   21
         Top             =   240
         Width           =   4335
      End
   End
End
Attribute VB_Name = "frmBankSetup"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title          :   frmBankSetup
'Author         :   Saurav Biswas Kartik
'URL/Mail       :   onsaurav@yahoo.com/gmail.com/hotmail.com
'Description    :   The form is used to set the banks
'Created        :   Saurav Biswas / Oct-25-2010
'Modified       :   Saurav Biswas /

Private Sub cmdAdd_Click()
        On Error GoTo Ext
        Dim rsBank As New ADODB.Recordset
        If CheckForm() = False Then Exit Sub
        
        Set rsBank = ONSCommon.SQL("SELECT * FROM Bank WHERE BnkBankId = '" & txtBankId.Text & "' OR BnkBankName = '" & txtBankName.Text & "'", rsBank)
        If rsBank.RecordCount > 0 Then
           MsgBox "Sorry! The bank id or name already exist.", vbInformation, "Sorry ..."
           txtBankName.SetFocus: Exit Sub
        End If
        
        cOnn.Execute "INSERT INTO Bank (BnkBankID, BnkBankName, BnkAddress, BnkPhone, BnkFax, BnkEmail, BnkURL) VALUES ('" & txtBankId.Text & "', '" & txtBankName.Text & "', '" & txtAddress.Text & "', '" & txtPhone1.Text & "', '" & txtFax.Text & "', '" & txtEmail.Text & "', '" & txtURL.Text & "')"
        MsgBox "Add the bank successfully.", vbInformation, "Add bank ..."
        Call LoadBanks
        Call Clear: txtBankName.SetFocus
        Exit Sub
Ext:
        'if an error occured then show the error message and return ""(blank space)
        MsgBox "Error No : " & Err.Number & Chr(13) & "Error : " & Err.Description, vbInformation, "Add bank ...": Err.Clear
End Sub

Private Sub cmdClear_Click()
        Call Clear
End Sub

Private Sub cmdDelete_Click()
        On Error GoTo Ext
        Dim rsBank As New ADODB.Recordset
        
        If CheckForm() = False Then Exit Sub
        If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub
        
        Set rsBank = ONSCommon.SQL("SELECT * FROM Bank WHERE BnkBankId = '" & txtBankId.Text & "'", rsBank)
        If rsBank.RecordCount = 0 Then
           MsgBox "Sorry! The bank id not found for edit.", vbInformation, "Sorry ...": Exit Sub
        End If
        
        cOnn.Execute "DELETE FROM Bank WHERE BnkBankID = '" & txtBankId.Text & "'"
        MsgBox "Delete the bank successfully.", vbInformation, "Delete bank ..."
        Call LoadBanks
        Call Clear: txtBankName.SetFocus
        Exit Sub
Ext:
        'if an error occured then show the error message and return ""(blank space)
        MsgBox "Error No : " & Err.Number & Chr(13) & "Error : " & Err.Description, vbInformation, "Add bank ...": Err.Clear
End Sub

Private Sub cmdEdit_Click()
        On Error GoTo Ext
        Dim rsBank As New ADODB.Recordset
        
        If CheckForm() = False Then Exit Sub
        If MsgBox("Are you sure you want to edit this record?", vbQuestion + vbYesNo, "Edit") = vbNo Then Exit Sub
        
        Set rsBank = ONSCommon.SQL("SELECT * FROM Bank WHERE BnkBankId = '" & txtBankId.Text & "'", rsBank)
        If rsBank.RecordCount = 0 Then
           MsgBox "Sorry! The bank id not found for edit.", vbInformation, "Sorry ...": Exit Sub
        End If
        
        Set rsBank = ONSCommon.SQL("SELECT * FROM Bank WHERE BnkBankId <> '" & txtBankId.Text & "' AND BnkBankName = '" & txtBankName.Text & "'", rsBank)
        If rsBank.RecordCount > 0 Then
           MsgBox "Sorry! The bank name already exist", vbInformation, "Sorry ..."
           txtBankName.SetFocus: Exit Sub
        End If
        
        cOnn.Execute "UPDATE Bank SET BnkBankName = '" & txtBankName.Text & "', BnkAddress = '" & txtAddress.Text & "', BnkPhone = '" & txtPhone1.Text & "', BnkFax = '" & txtFax.Text & "', BnkEmail = '" & txtEmail.Text & "', BnkURL = '" & txtURL.Text & "' WHERE BnkBankID = '" & txtBankId.Text & "'"
        MsgBox "Edit the bank successfully.", vbInformation, "Edit bank ..."
        Call LoadBanks
        Call Clear: txtBankName.SetFocus
        Exit Sub
Ext:
        'if an error occured then show the error message and return ""(blank space)
        MsgBox "Error No : " & Err.Number & Chr(13) & "Error : " & Err.Description, vbInformation, "Add bank ...": Err.Clear
End Sub

Private Sub Form_Activate()
        On Error Resume Next
        Me.SetFocus
End Sub

Private Sub Form_Load()
        Call Clear
        Call LoadBanks
End Sub

Private Sub lvwBank_DblClick()
        'Summary    :   Method used to load all information of the selected banl
        'Created    :   Saurav Biswas / Oct-25-2010
        'Modified   :   Saurav Biswas /
        'Parameters :
                
        Dim TEMP
        Dim rsBank As New ADODB.Recordset
        
        Call Clear
        If lvwBank.ListItems.Count = 0 Then Exit Sub
        
        'Split the List view item's key
        TEMP = Split(lvwBank.SelectedItem.Key, "<ONS>")
        Set rsBank = ONSCommon.SQL("SELECT * FROM Bank WHERE BnkBankId = '" & TEMP(0) & "'", rsBank)
        If rsBank.RecordCount > 0 Then
           txtBankId.Text = IIf(IsNull(rsBank!BnkBankId) = True, "", rsBank!BnkBankId)
           txtBankName.Text = IIf(IsNull(rsBank!BnkBankName) = True, "", rsBank!BnkBankName)
           txtAddress.Text = IIf(IsNull(rsBank!BnkAddress) = True, "", rsBank!BnkAddress)
           txtPhone1.Text = IIf(IsNull(rsBank!BnkPhone) = True, "", rsBank!BnkPhone)
           txtFax.Text = IIf(IsNull(rsBank!BnkFax) = True, "", rsBank!BnkFax)
           txtEmail.Text = IIf(IsNull(rsBank!BnkEmail) = True, "", rsBank!BnkEmail)
           txtURL.Text = IIf(IsNull(rsBank!BnkURL) = True, "", rsBank!BnkURL)
        End If
        txtBankName.SetFocus
End Sub

Private Sub txtAddress_KeyPress(KeyAscii As Integer)
        'Checking for invalid input string and send to next tab index if press Enter
        Dim Invalid As String
        Invalid = "',[];"""
        If InStr(Invalid, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtBankName_KeyPress(KeyAscii As Integer)
        'Checking for invalid input string and send to next tab index if press Enter
        Dim Invalid As String
        Invalid = "',[];"""
        If InStr(Invalid, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtEmail_KeyPress(KeyAscii As Integer)
        'Checking for invalid input string and send to next tab index if press Enter
        Dim Invalid As String
        Invalid = "',[];"""
        If InStr(Invalid, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtFax_KeyPress(KeyAscii As Integer)
        'Checking for invalid input string and send to next tab index if press Enter
        Dim Invalid As String
        Invalid = "',[];"""
        If InStr(Invalid, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtPhone1_KeyPress(KeyAscii As Integer)
        'Checking for invalid input string and send to next tab index if press Enter
        Dim Invalid As String
        Invalid = "',[];"""
        If InStr(Invalid, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtURL_KeyPress(KeyAscii As Integer)
        'Checking for invalid input string and send to next tab index if press Enter
        Dim Invalid As String
        Invalid = "',[];"""
        If InStr(Invalid, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Clear()
        'Summary    :   Function used to clear the form elements
        'Created    :   Saurav Biswas / Oct-25-2010
        'Modified   :   Saurav Biswas /
        'Parameters :
        
        txtBankId.Text = ONSCommon.RandString(3)
        txtBankName.Text = ""
        txtAddress.Text = ""
        txtPhone1.Text = ""
        txtFax.Text = ""
        txtEmail.Text = ""
        txtURL.Text = ""
End Sub

Private Function CheckForm() As Boolean
        'Summary    :   Function used to check the form inputs
        'Created    :   Saurav Biswas / Oct-25-2010
        'Modified   :   Saurav Biswas /
        'Parameters :
        
        If Trim(txtBankId.Text) = "" Then
           MsgBox "The Bank Id is Blank.", vbInformation, "Sorry ...": CheckForm = False: Exit Function
        End If
        If Trim(txtBankName.Text) = "" Then
           MsgBox "The Bank Name is Blank.", vbInformation, "Sorry ...": CheckForm = False: txtBankName.SetFocus: Exit Function
        End If
        CheckForm = True
End Function

Private Sub LoadBanks()
        'Summary    :   Function used to load all banks from database in the listview
        'Created    :   Saurav Biswas / Oct-25-2010
        'Modified   :   Saurav Biswas /
        'Parameters :
        
        Dim LTM As ListItem
        Dim rsBank As New ADODB.Recordset
        
        lvwBank.ListItems.Clear
        Set rsBank = ONSCommon.SQL("SELECT * FROM Bank ORDER BY BnkBankName", rsBank)
        While Not rsBank.EOF And Not rsBank.BOF
              Set LTM = lvwBank.ListItems.Add(, rsBank!BnkBankId & "<ONS>" & rsBank.AbsolutePosition, rsBank!BnkBankName)
              rsBank.MoveNext
        Wend
End Sub
