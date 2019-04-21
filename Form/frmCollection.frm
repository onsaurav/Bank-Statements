VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Object = "{86CF1D34-0C5F-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCT2.OCX"
Begin VB.Form frmCollection 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Collection"
   ClientHeight    =   6615
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   11160
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   6615
   ScaleWidth      =   11160
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6C495&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   19
      Top             =   0
      Width           =   11175
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Collection"
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
         TabIndex        =   20
         Top             =   240
         Width           =   4335
      End
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E6C495&
      BorderStyle     =   0  'None
      Caption         =   "v"
      Height          =   855
      Left            =   0
      TabIndex        =   18
      Top             =   5760
      Width           =   11175
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   0  'None
      Height          =   4815
      Left            =   0
      TabIndex        =   8
      Top             =   840
      Width           =   11175
      Begin VB.TextBox txtSearch 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   6600
         TabIndex        =   25
         Top             =   4440
         Width           =   3255
      End
      Begin VB.CommandButton cmdClear 
         Caption         =   "Clear"
         Height          =   325
         Left            =   5160
         TabIndex        =   12
         Top             =   4380
         Width           =   1095
      End
      Begin VB.CommandButton cmdDelete 
         Caption         =   "Delete"
         Height          =   325
         Left            =   4080
         TabIndex        =   11
         Top             =   4380
         Width           =   1095
      End
      Begin VB.CommandButton cmdEdit 
         Caption         =   "Edit"
         Height          =   325
         Left            =   3000
         TabIndex        =   10
         Top             =   4380
         Width           =   1095
      End
      Begin VB.CommandButton cmdAdd 
         Caption         =   "Add"
         Height          =   325
         Left            =   1920
         TabIndex        =   9
         Top             =   4380
         Width           =   1095
      End
      Begin MSComCtl2.DTPicker dpkChequeDate 
         Height          =   300
         Left            =   1920
         TabIndex        =   5
         Top             =   2280
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85721089
         CurrentDate     =   40476
      End
      Begin VB.ComboBox cboBank 
         Appearance      =   0  'Flat
         Height          =   315
         Left            =   1920
         Style           =   2  'Dropdown List
         TabIndex        =   3
         Top             =   1560
         Width           =   4335
      End
      Begin VB.TextBox txtCollectionId 
         Appearance      =   0  'Flat
         Enabled         =   0   'False
         Height          =   285
         Left            =   1920
         TabIndex        =   0
         Top             =   840
         Width           =   4335
      End
      Begin VB.TextBox txtCheque 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   4
         Top             =   1920
         Width           =   4335
      End
      Begin VB.TextBox txtAmount 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1920
         TabIndex        =   6
         Top             =   2640
         Width           =   1455
      End
      Begin VB.TextBox txtRemarks 
         Appearance      =   0  'Flat
         Height          =   765
         Left            =   1920
         TabIndex        =   7
         Top             =   3000
         Width           =   4335
      End
      Begin MSComctlLib.ListView lvwCollection 
         Height          =   4095
         Left            =   6600
         TabIndex        =   13
         Top             =   240
         Width           =   4215
         _ExtentX        =   7435
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
         NumItems        =   2
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Collection No"
            Object.Width           =   3704
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cheque No"
            Object.Width           =   3175
         EndProperty
      End
      Begin MSComCtl2.DTPicker dpkDate 
         Height          =   300
         Left            =   1920
         TabIndex        =   1
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85721089
         CurrentDate     =   40476
      End
      Begin MSComCtl2.DTPicker dpkTime 
         Height          =   300
         Left            =   4800
         TabIndex        =   2
         Top             =   1200
         Width           =   1455
         _ExtentX        =   2566
         _ExtentY        =   529
         _Version        =   393216
         Format          =   85721090
         CurrentDate     =   40476
      End
      Begin VB.Label Label6 
         Alignment       =   1  'Right Justify
         Appearance      =   0  'Flat
         BackColor       =   &H80000005&
         BackStyle       =   0  'Transparent
         BorderStyle     =   1  'Fixed Single
         Caption         =   "SEARCH"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   8.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H80000008&
         Height          =   285
         Left            =   9840
         TabIndex        =   26
         Top             =   4440
         Width           =   975
      End
      Begin VB.Label Label7 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque No"
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
         Left            =   360
         TabIndex        =   24
         Top             =   2280
         Width           =   1095
      End
      Begin VB.Label Label11 
         BackStyle       =   0  'Transparent
         Caption         =   "Time"
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
         Left            =   4080
         TabIndex        =   23
         Top             =   1200
         Width           =   615
      End
      Begin VB.Label Label10 
         BackStyle       =   0  'Transparent
         Caption         =   "Date"
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
         Left            =   360
         TabIndex        =   22
         Top             =   1200
         Width           =   1095
      End
      Begin VB.Label Label8 
         BackStyle       =   0  'Transparent
         Caption         =   "Collection No"
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
         Left            =   360
         TabIndex        =   21
         Top             =   840
         Width           =   1215
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
         Left            =   360
         TabIndex        =   17
         Top             =   1560
         Width           =   1095
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque No"
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
         Left            =   360
         TabIndex        =   16
         Top             =   1920
         Width           =   1095
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Amount"
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
         Left            =   360
         TabIndex        =   15
         Top             =   2640
         Width           =   1095
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Remarks"
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
         Left            =   360
         TabIndex        =   14
         Top             =   3000
         Width           =   1095
      End
   End
End
Attribute VB_Name = "frmCollection"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title          :   frmC ollection
'Author         :   Saurav Biswas Kartik
'URL/Mail       :   onsaurav@yahoo.com/gmail.com/hotmail.com
'Description    :   The form is used to insert the collection
'Created        :   Saurav Biswas / Oct-25-2010
'Modified       :   Saurav Biswas /

Private Sub cboBank_KeyPress(KeyAscii As Integer)
        'Checking for invalid input string and send to next tab index if press Enter
        Dim Invalid As String
        Invalid = "',[];"""
        If InStr(Invalid, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub cmdAdd_Click()
        On Error GoTo Ext
        Dim rsCollection As New ADODB.Recordset
        If CheckForm() = False Then Exit Sub
        
        Set rsCollection = ONSCommon.SQL("SELECT * FROM ChequeDetail WHERE CqdCollectionNo = '" & txtCollectionId.Text & "' OR (CqdChequeNo = '" & txtCheque.Text & "' And CqdBankId = '" & cboBank.Tag & "')", rsCollection)
        If rsCollection.RecordCount > 0 Then
           MsgBox "Sorry! The Collection id or check no already exist.", vbInformation, "Sorry ..."
           txtCheque.SetFocus: Exit Sub
        End If
        
        cOnn.Execute "INSERT INTO ChequeDetail (CqdCollectionNo, CqdChequeNo, CqdBankId, CqdChequeDate, CqdAmount, CqdCollected, CqdTreateDate, CqdDate, CqdTime, CqdRemarks) VALUES ('" & txtCollectionId.Text & "', '" & txtCheque.Text & "', '" & cboBank.Tag & "', '" & FormatDateTime(dpkChequeDate, vbShortDate) & "', " & Val(txtAmount.Text) & ", 'N', '" & FormatDateTime(Date, vbShortDate) & "', '" & FormatDateTime(Date, vbShortDate) & "', '" & FormatDateTime(Time, vbLongTime) & "', '" & txtRemarks.Text & "')"
        MsgBox "Add the Collection successfully.", vbInformation, "Add Collection ..."
        Call LoadCollections("")
        Call Clear: dpkDate.SetFocus
        Exit Sub
Ext:
        'if an error occured then show the error message and return ""(blank space)
        MsgBox "Error No : " & Err.Number & Chr(13) & "Error : " & Err.Description, vbInformation, "Add Collection ...": Err.Clear
End Sub

Private Sub cmdClear_Click()
        Call Clear
End Sub

Private Sub cmdDelete_Click()
        On Error GoTo Ext
        Dim rsCollection As New ADODB.Recordset
        If CheckForm() = False Then Exit Sub
        If MsgBox("Are you sure you want to delete this record?", vbQuestion + vbYesNo, "Delete") = vbNo Then Exit Sub
        
        Set rsCollection = ONSCommon.SQL("SELECT * FROM ChequeDetail WHERE CqdCollectionNo = '" & txtCollectionId.Text & "'", rsCollection)
        If rsCollection.RecordCount = 0 Then
           MsgBox "Sorry! Collection id not found for edit", vbInformation, "Sorry ..."
           txtCheque.SetFocus: Exit Sub
        End If
        
        cOnn.Execute "DELETE FROM ChequeDetail WHERE CqdCollectionNo = '" & txtCollectionId.Text & "'"
        cOnn.Execute "DELETE FROM ChequeHistory WHERE ChsCollectionNo = '" & txtCollectionId.Text & "'"
        
        MsgBox "Delete the Collection successfully.", vbInformation, "Edit Collection ..."
        Call LoadCollections("")
        Call Clear: dpkDate.SetFocus
        Exit Sub
Ext:
        'if an error occured then show the error message and return ""(blank space)
        MsgBox "Error No : " & Err.Number & Chr(13) & "Error : " & Err.Description, vbInformation, "Delete Collection ...": Err.Clear
End Sub

Private Sub cmdEdit_Click()
        On Error GoTo Ext
        Dim rsCollection As New ADODB.Recordset
        If CheckForm() = False Then Exit Sub
        If MsgBox("Are you sure you want to edit this record?", vbQuestion + vbYesNo, "Edit") = vbNo Then Exit Sub
        
        Set rsCollection = ONSCommon.SQL("SELECT * FROM ChequeDetail WHERE CqdCollectionNo = '" & txtCollectionId.Text & "'", rsCollection)
        If rsCollection.RecordCount = 0 Then
           MsgBox "Sorry! Collection id not found for edit", vbInformation, "Sorry ..."
           txtCheque.SetFocus: Exit Sub
        End If
        
        Set rsCollection = ONSCommon.SQL("SELECT * FROM ChequeDetail WHERE CqdCollectionNo <> '" & txtCollectionId.Text & "' AND (CqdChequeNo = '" & txtCheque.Text & "' And CqdBankId = '" & cboBank.Tag & "')", rsCollection)
        If rsCollection.RecordCount > 0 Then
           MsgBox "Sorry! Same cheque already exist", vbInformation, "Sorry ..."
           txtCheque.SetFocus: Exit Sub
        End If
        
        cOnn.Execute "UPDATE ChequeDetail SET CqdChequeNo = '" & txtCheque.Text & "', CqdBankId = '" & cboBank.Tag & "', CqdChequeDate = '" & FormatDateTime(dpkChequeDate, vbShortDate) & "', CqdAmount =  " & Val(txtAmount.Text) & ", CqdDate = '" & FormatDateTime(Date, vbShortDate) & "', CqdTime ='" & FormatDateTime(Time, vbShortDate) & "', CqdRemarks = '" & txtRemarks.Text & "' WHERE CqdCollectionNo = '" & txtCollectionId.Text & "'"
        MsgBox "Edit the Collection successfully.", vbInformation, "Edit Collection ..."
        Call LoadCollections("")
        Call Clear: dpkDate.SetFocus
        Exit Sub
Ext:
        'if an error occured then show the error message and return ""(blank space)
        MsgBox "Error No : " & Err.Number & Chr(13) & "Error : " & Err.Description, vbInformation, "Edit Collection ...": Err.Clear
End Sub

Private Sub dpkChequeDate_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub dpkDate_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub dpkTime_KeyDown(KeyCode As Integer, Shift As Integer)
        If KeyCode = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub Form_Activate()
        On Error Resume Next
        Me.SetFocus
End Sub

Private Sub Form_Load()
        Call ONSCommon.LoadCombo("SELECT BnkBankName FROM Bank ORDER BY BnkBankName", cboBank)
        Call Clear
        Call LoadCollections("")
End Sub

Private Sub lvwCollection_DblClick()
        'Summary    :   Method used to load all information of the selected banl
        'Created    :   Saurav Biswas / Oct-25-2010
        'Modified   :   Saurav Biswas /
        'Parameters :
                
        Dim TEMP
        Dim rsCollection As New ADODB.Recordset
        
        Call Clear
        If lvwCollection.ListItems.Count = 0 Then Exit Sub
        
        'Split the List view item's key
        TEMP = Split(lvwCollection.SelectedItem.Key, "<ONS>")
        Set rsCollection = ONSCommon.SQL("SELECT * FROM ChequeDetail, Bank WHERE CqdBankId=BnkBankId AND CqdCollectionNo = '" & TEMP(0) & "'", rsCollection)
        If rsCollection.RecordCount > 0 Then
           txtCollectionId.Text = IIf(IsNull(rsCollection!CqdCollectionNo) = True, "", rsCollection!CqdCollectionNo)
           dpkDate.Value = IIf(IsNull(rsCollection!CqdDate) = True, Date, rsCollection!CqdDate)
           dpkTime.Value = IIf(IsNull(rsCollection!CqdTime) = True, Time, rsCollection!CqdTime)
           cboBank.Text = IIf(IsNull(rsCollection!BnkBankName) = True, "", rsCollection!BnkBankName)
           txtCheque.Text = IIf(IsNull(rsCollection!CqdChequeNo) = True, "", rsCollection!CqdChequeNo)
           dpkChequeDate.Value = IIf(IsNull(rsCollection!CqdChequeDate) = True, Date, rsCollection!CqdChequeDate)
           txtAmount.Text = IIf(IsNull(rsCollection!CqdAmount) = True, "", rsCollection!CqdAmount)
           txtRemarks.Text = IIf(IsNull(rsCollection!CqdRemarks) = True, "", rsCollection!CqdRemarks)
        End If
        dpkDate.SetFocus
End Sub

Private Sub txtAmount_Change()
        If IsNumeric(txtAmount.Text) = False Then txtAmount.Text = ""
End Sub

Private Sub txtAmount_KeyPress(KeyAscii As Integer)
        'Checking for invalid input string and send to next tab index if press Enter
        If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtCheque_KeyPress(KeyAscii As Integer)
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
        
        txtCollectionId.Text = GenerateAutoId("CL", "ChequeDetail", "CqdCollectionNo")
        dpkDate.Value = Date
        dpkTime.Value = Time
        If cboBank.ListCount > 0 Then cboBank.ListIndex = 0
        txtCheque.Text = ""
        dpkChequeDate.Value = Date
        txtAmount.Text = ""
        txtRemarks.Text = ""
End Sub

Private Function CheckForm() As Boolean
        'Summary    :   Function used to check the form inputs
        'Created    :   Saurav Biswas / Oct-25-2010
        'Modified   :   Saurav Biswas /
        'Parameters :
        
        If Trim(txtCollectionId.Text) = "" Then
           MsgBox "The Collection Id is Blank.", vbInformation, "Sorry ...": CheckForm = False: Exit Function
        End If
        
        If Trim(txtCheque.Text) = "" Then
           MsgBox "Sorry ... Invalid cheque amount.", vbInformation, "Sorry ...": CheckForm = False: txtCheque.SetFocus: Exit Function
        End If
        
        If Trim(txtAmount.Text) = "" Then
           MsgBox "Sorry ... Invalid collection amount.", vbInformation, "Sorry ...": CheckForm = False: txtAmount.SetFocus: Exit Function
        End If
        
        Dim rsCheck As New ADODB.Recordset
        Set rsCheck = ONSCommon.SQL("SELECT * FROM Bank WHERE BnkBankName = '" & cboBank.Text & "'", rsCheck)
        If rsCheck.RecordCount > 0 Then
           cboBank.Tag = IIf(IsNull(rsCheck!BnkBankId) = True, "", rsCheck!BnkBankId)
        Else
           MsgBox "Sorry ... Invalid bank SELECTed.", vbInformation, "Sorry ..."
           cboBank.Tag = "": CheckForm = False: cboBank.SetFocus: Exit Function
        End If
        CheckForm = True
End Function

Private Sub LoadCollections(StrId As String)
        'Summary    :   Function used to load all Collections from database in the listview
        'Created    :   Saurav Biswas / Oct-25-2010
        'Modified   :   Saurav Biswas /
        'Parameters :
        
        Dim LTM As ListItem
        Dim rsCollection As New ADODB.Recordset
        
        lvwCollection.ListItems.Clear
        Set rsCollection = ONSCommon.SQL("SELECT * FROM ChequeDetail WHERE CqdCollectionNo Like '%" & StrId & "%' ORDER BY CqdCollectionNo", rsCollection)
        While Not rsCollection.EOF And Not rsCollection.BOF
              Set LTM = lvwCollection.ListItems.Add(, rsCollection!CqdCollectionNo & "<ONS>" & rsCollection.AbsolutePosition, rsCollection!CqdCollectionNo)
              LTM.SubItems(1) = IIf(IsNull(rsCollection!CqdChequeNo) = True, "", rsCollection!CqdChequeNo)
              rsCollection.MoveNext
        Wend
End Sub

Private Sub txtRemarks_KeyPress(KeyAscii As Integer)
        'Checking for invalid input string and send to next tab index if press Enter
        Dim Invalid As String
        Invalid = "',[];"""
        If InStr(Invalid, Chr(KeyAscii)) <> 0 Then KeyAscii = 0
        If KeyAscii = 13 Then SendKeys ("{TAB}")
End Sub

Private Sub txtSearch_Change()
        LoadCollections (Trim(txtSearch.Text))
End Sub
