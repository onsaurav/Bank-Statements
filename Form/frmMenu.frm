VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmMenu 
   BackColor       =   &H00FFFFFF&
   Caption         =   "Bank Status Monitoring"
   ClientHeight    =   9135
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   14535
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   9135
   ScaleWidth      =   14535
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame fmHistory 
      Appearance      =   0  'Flat
      BackColor       =   &H00E6C495&
      BorderStyle     =   0  'None
      Caption         =   "Frame3"
      ForeColor       =   &H80000008&
      Height          =   2775
      Left            =   5280
      TabIndex        =   8
      Top             =   3120
      Visible         =   0   'False
      Width           =   8055
      Begin MSComctlLib.ListView lvwHistory 
         Height          =   2055
         Left            =   240
         TabIndex        =   9
         Top             =   480
         Width           =   7575
         _ExtentX        =   13361
         _ExtentY        =   3625
         View            =   3
         Arrange         =   1
         LabelWrap       =   -1  'True
         HideSelection   =   -1  'True
         FullRowSelect   =   -1  'True
         GridLines       =   -1  'True
         _Version        =   393217
         ForeColor       =   -2147483640
         BackColor       =   -2147483643
         BorderStyle     =   1
         Appearance      =   0
         NumItems        =   5
         BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            Text            =   "Collection No"
            Object.Width           =   3528
         EndProperty
         BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   1
            Text            =   "Cheque No"
            Object.Width           =   3175
         EndProperty
         BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   2
            Text            =   "Amount"
            Object.Width           =   2028
         EndProperty
         BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   3
            Text            =   "Status"
            Object.Width           =   1852
         EndProperty
         BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
            SubItemIndex    =   4
            Text            =   "Treate Date"
            Object.Width           =   2205
         EndProperty
      End
      Begin VB.Label lblClose 
         BackStyle       =   0  'Transparent
         Caption         =   "X"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   13.5
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   7560
         TabIndex        =   11
         Top             =   120
         Width           =   255
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Cheque History"
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   14.25
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   375
         Left            =   240
         TabIndex        =   10
         Top             =   60
         Width           =   4335
      End
   End
   Begin MSComctlLib.ProgressBar P 
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   7920
      Visible         =   0   'False
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   450
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00E6C495&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   4
      Top             =   8280
      Width           =   14535
      Begin VB.CommandButton cmdClose 
         Caption         =   "Close"
         Height          =   300
         Left            =   13440
         TabIndex        =   5
         Top             =   360
         Width           =   975
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         BackColor       =   &H00FFFFFF&
         BackStyle       =   0  'Transparent
         Caption         =   $"frmMenu.frx":0000
         BeginProperty Font 
            Name            =   "Times New Roman"
            Size            =   8.25
            Charset         =   0
            Weight          =   400
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   630
         Left            =   120
         TabIndex        =   6
         Top             =   120
         Width           =   12135
         WordWrap        =   -1  'True
      End
   End
   Begin MSComctlLib.ListView lvwCheque 
      Height          =   7215
      Left            =   4200
      TabIndex        =   1
      Top             =   960
      Width           =   10215
      _ExtentX        =   18018
      _ExtentY        =   12726
      View            =   3
      Arrange         =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   0
      NumItems        =   6
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Collection No"
         Object.Width           =   3440
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Bank Name"
         Object.Width           =   4762
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Cheque No"
         Object.Width           =   3087
      EndProperty
      BeginProperty ColumnHeader(4) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   3
         Text            =   "Amount"
         Object.Width           =   2028
      EndProperty
      BeginProperty ColumnHeader(5) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   4
         Text            =   "Status"
         Object.Width           =   1852
      EndProperty
      BeginProperty ColumnHeader(6) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   5
         Text            =   "Treate Date"
         Object.Width           =   2205
      EndProperty
   End
   Begin MSComctlLib.TreeView tvwOptions 
      Height          =   7215
      Left            =   120
      TabIndex        =   0
      Top             =   960
      Width           =   3855
      _ExtentX        =   6800
      _ExtentY        =   12726
      _Version        =   393217
      Style           =   7
      BorderStyle     =   1
      Appearance      =   0
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00E6C495&
      BorderStyle     =   0  'None
      Height          =   855
      Left            =   0
      TabIndex        =   2
      Top             =   0
      Width           =   14535
      Begin VB.Image imgBankSetup 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   12360
         Picture         =   "frmMenu.frx":017E
         Stretch         =   -1  'True
         ToolTipText     =   "Bank Setup"
         Top             =   120
         Width           =   975
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "Bank Status Monitoring"
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
         TabIndex        =   3
         Top             =   240
         Width           =   4335
      End
      Begin VB.Image imgCollection 
         Appearance      =   0  'Flat
         BorderStyle     =   1  'Fixed Single
         Height          =   615
         Left            =   13440
         Picture         =   "frmMenu.frx":4A0D
         Stretch         =   -1  'True
         ToolTipText     =   "Collection"
         Top             =   120
         Width           =   975
      End
   End
   Begin VB.Menu mnuMenu 
      Caption         =   "Menu"
      Visible         =   0   'False
      Begin VB.Menu mnuHonored 
         Caption         =   "Honored"
      End
      Begin VB.Menu mnuDisHonored 
         Caption         =   "DisHonored"
      End
      Begin VB.Menu mnuSended 
         Caption         =   "Sended"
      End
      Begin VB.Menu mnuHistory 
         Caption         =   "History"
      End
      Begin VB.Menu mnuRefresh 
         Caption         =   "Refresh"
      End
   End
End
Attribute VB_Name = "frmMenu"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
'Title          :   frmMenu
'Author         :   Saurav Biswas Kartik
'URL/Mail       :   onsaurav@yahoo.com/gmail.com/hotmail.com
'Description    :   The main form of this project, Holding the check monitoring status and other options
'Created        :   Saurav Biswas / Oct-25-2010
'Modified       :   Saurav Biswas /

Private Sub cmdClose_Click()
        'confirmation of user to close the project
        If MsgBox("Are you sure you want to close it?", vbQuestion + vbYesNo, "Close ...") = vbNo Then Exit Sub
        End
End Sub

Private Sub Form_Load()
        Call LoadTree
End Sub

Private Sub imgBankSetup_Click()
        frmBankSetup.Show
End Sub

Private Sub imgCollection_Click()
        frmCollection.Show
End Sub

Private Sub imgTreatement_Click()
        frmTreatement.Show
End Sub

Public Function LoadTree()
       'Summary    :   This function will use to load the tree view
       'Created    :   Saurav Biswas / Oct-25-2010
       'Modified   :   Saurav Biswas /
       'Parameters :
                       
       Dim NRoot As Node
       Dim rsBank As New ADODB.Recordset
       Dim rsCheque As New ADODB.Recordset
            
       'Clear the treeview
       tvwOptions.Nodes.Clear
       
       'Load the root node in the treeview
       Set NRoot = tvwOptions.Nodes.Add(, , "Root", "Collection")
       NRoot.Expanded = True: NRoot.Bold = True
       
       Set rsBank = ONSCommon.SQL("SELECT Distinct BnkBankId, BnkBankName FROM Bank WHERE BnkBankId IN (SELECT DISTINCT CqdBankId FROM ChequeDetail) ORDER BY BnkBankName", rsBank)
       P.Visible = True: P.Min = 0: P.Value = 0: If rsBank.RecordCount > 0 Then P.Max = rsBank.RecordCount
       While Not rsBank.EOF And Not rsBank.BOF
             DoEvents: P.Value = rsBank.RecordCount
             Set NRoot = tvwOptions.Nodes.Add("Root", tvwChild, rsBank!BnkBankId & "<ONS>" & rsBank.AbsolutePosition, rsBank!BnkBankName)
             NRoot.Bold = True
             Set rsCheque = ONSCommon.SQL("SELECT DISTINCT CqdCollectionNo FROM ChequeDetail WHERE CqdBankId = '" & rsBank!BnkBankId & "' ORDER BY CqdCollectionNo", rsCheque)
             While Not rsCheque.EOF And Not rsCheque.BOF
                   Set NRoot = tvwOptions.Nodes.Add(rsBank!BnkBankId & "<ONS>" & rsBank.AbsolutePosition, tvwChild, rsBank!BnkBankId & "<ONS>" & rsCheque!CqdCollectionNo & "<ONS>" & rsCheque.AbsolutePosition, rsCheque!CqdCollectionNo)
                   rsCheque.MoveNext
             Wend
             rsBank.MoveNext
       Wend
       
       'Expend the root node
       tvwOptions.Nodes.Item("Root").Expanded = True
       P.Visible = False
End Function

Private Sub lblClose_Click()
        fmHistory.Visible = False
End Sub

Private Sub lvwCheque_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
        If Button = vbRightButton Then PopupMenu mnuMenu
End Sub

Private Sub mnuDisHonored_Click()
        Dim TEMP
        Dim rsCheck As New ADODB.Recordset
        If lvwCheque.ListItems.Count = 0 Then MsgBox "Sorry! No item exist.", vbExclamation, "Sorry ...": Exit Sub
        
        TEMP = Split(lvwCheque.SelectedItem.Key, "<ONS>")
        If lvwCheque.SelectedItem.ListSubItems.Item(4).Text = "DisHonored" Then
           MsgBox "Cheque already honored.", vbExclamation, "Sorry ...": Exit Sub
        Else
           Set rsCheck = ONSCommon.SQL("SELECT * FROM ChequeDetail WHERE CqdCollectionNo = '" & TEMP(0) & "'", rsCheck)
           If rsCheck.RecordCount > 0 Then
              cOnn.Execute "UPDATE ChequeDetail SET CqdTreateDate = '" & FormatDateTime(Date, vbShortDate) & "', CqdCollected = 'D' WHERE CqdCollectionNo = '" & TEMP(0) & "'"
              cOnn.Execute "INSERT INTO ChequeHistory (ChsCollectionNo, ChsChequeNo, ChsChequeDate, ChsBankId, ChsTreatment, ChsDate, ChsTime, ChsUser, ChsTreatmentDate) VALUES ('" & rsCheck!CqdCollectionNo & "', '" & rsCheck!CqdChequeNo & "', '" & rsCheck!CqdChequeDate & "', '" & rsCheck!CqdBankId & "', 'D', '" & FormatDateTime(Date, vbShortDate) & "', '" & FormatDateTime(Time, vbLongTime) & "', 'AUTO', '" & FormatDateTime(Date, vbShortDate) & "')"
           End If
        End If
        Call tvwOptions_DblClick
End Sub

Private Sub mnuHistory_Click()
        Dim LTM As ListItem
        Dim rsCheque As New ADODB.Recordset
        
        If tvwOptions.Nodes.Count = 0 Then Exit Sub
        TEMP = Split(lvwCheque.SelectedItem.Key, "<ONS>")
        
        lvwHistory.ListItems.Clear
        Set rsCheque = ONSCommon.SQL("SELECT * FROM ChequeHistory WHERE ChsCollectionNo = '" & TEMP(0) & "' ORDER BY ChsTreatmentDate", rsCheque)
        While Not rsCheque.EOF And Not rsCheque.BOF
              Set LTM = lvwHistory.ListItems.Add(, rsCheque!ChsCollectionNo & "<ONS>" & rsCheque.AbsolutePosition, rsCheque!ChsCollectionNo)
              LTM.SubItems(1) = IIf(IsNull(rsCheque!ChsChequeNo) = True, "", rsCheque!ChsChequeNo)
              LTM.SubItems(2) = lvwCheque.SelectedItem.ListSubItems.Item(3).Text
              
              If rsCheque!ChsTreatment = "H" Then
                 LTM.SubItems(3) = "Honored"
              ElseIf rsCheque!ChsTreatment = "D" Then
                 LTM.SubItems(3) = "DisHonored"
              ElseIf rsCheque!ChsTreatment = "S" Then
                 LTM.SubItems(3) = "Sended"
              Else
                 LTM.SubItems(3) = "Un Treate"
              End If
              
              LTM.SubItems(4) = IIf(IsNull(rsCheque!ChsTreatmentDate) = True, "", rsCheque!ChsTreatmentDate)
              rsCheque.MoveNext
        Wend
        fmHistory.Visible = True
End Sub

Private Sub mnuHonored_Click()
        Dim TEMP
        Dim rsCheck As New ADODB.Recordset
        If lvwCheque.ListItems.Count = 0 Then MsgBox "Sorry! No item exist.", vbExclamation, "Sorry ...": Exit Sub
        
        TEMP = Split(lvwCheque.SelectedItem.Key, "<ONS>")
        If lvwCheque.SelectedItem.ListSubItems.Item(4).Text = "Honored" Then
           MsgBox "Cheque already honored.", vbExclamation, "Sorry ...": Exit Sub
        Else
           Set rsCheck = ONSCommon.SQL("SELECT * FROM ChequeDetail WHERE CqdCollectionNo = '" & TEMP(0) & "'", rsCheck)
           If rsCheck.RecordCount > 0 Then
              cOnn.Execute "UPDATE ChequeDetail SET CqdTreateDate = '" & FormatDateTime(Date, vbShortDate) & "', CqdCollected = 'H' WHERE CqdCollectionNo = '" & TEMP(0) & "'"
              cOnn.Execute "INSERT INTO ChequeHistory (ChsCollectionNo, ChsChequeNo, ChsChequeDate, ChsBankId, ChsTreatment, ChsDate, ChsTime, ChsUser, ChsTreatmentDate) VALUES ('" & rsCheck!CqdCollectionNo & "', '" & rsCheck!CqdChequeNo & "', '" & rsCheck!CqdChequeDate & "', '" & rsCheck!CqdBankId & "', 'H', '" & FormatDateTime(Date, vbShortDate) & "', '" & FormatDateTime(Time, vbLongTime) & "', 'AUTO', '" & FormatDateTime(Date, vbShortDate) & "')"
           End If
        End If
        Call tvwOptions_DblClick
End Sub

Private Sub mnuRefresh_Click()
        lvwCheque.ListItems.Clear
        Call LoadTree
End Sub

Private Sub mnuSended_Click()
        Dim TEMP
        Dim rsCheck As New ADODB.Recordset
        If lvwCheque.ListItems.Count = 0 Then MsgBox "Sorry! No item exist.", vbExclamation, "Sorry ...": Exit Sub
        
        TEMP = Split(lvwCheque.SelectedItem.Key, "<ONS>")
        If lvwCheque.SelectedItem.ListSubItems.Item(4).Text = "Sended" Then
           MsgBox "Cheque already honored.", vbExclamation, "Sorry ...": Exit Sub
        Else
           Set rsCheck = ONSCommon.SQL("SELECT * FROM ChequeDetail WHERE CqdCollectionNo = '" & TEMP(0) & "'", rsCheck)
           If rsCheck.RecordCount > 0 Then
              cOnn.Execute "UPDATE ChequeDetail SET CqdTreateDate = '" & FormatDateTime(Date, vbShortDate) & "', CqdCollected = 'S' WHERE CqdCollectionNo = '" & TEMP(0) & "'"
              cOnn.Execute "INSERT INTO ChequeHistory (ChsCollectionNo, ChsChequeNo, ChsChequeDate, ChsBankId, ChsTreatment, ChsDate, ChsTime, ChsUser, ChsTreatmentDate) VALUES ('" & rsCheck!CqdCollectionNo & "', '" & rsCheck!CqdChequeNo & "', '" & rsCheck!CqdChequeDate & "', '" & rsCheck!CqdBankId & "', 'S', '" & FormatDateTime(Date, vbShortDate) & "', '" & FormatDateTime(Time, vbLongTime) & "', 'AUTO', '" & FormatDateTime(Date, vbShortDate) & "')"
           End If
        End If
        Call tvwOptions_DblClick
End Sub

Private Sub tvwOptions_DblClick()
        'Summary    :   This method will used to load cdetail collection
        'Created    :   Saurav Biswas / Oct-25-2010
        'Modified   :   Saurav Biswas /
        'Parameters :
       
        Dim TEMP
        Dim LTM As ListItem
        Dim rsCheque As New ADODB.Recordset
        
        If tvwOptions.Nodes.Count = 0 Then Exit Sub
        TEMP = Split(tvwOptions.SelectedItem.Key, "<ONS>")
        
        lvwCheque.ListItems.Clear
        If UBound(TEMP) = 1 Then
           Set rsCheque = ONSCommon.SQL("SELECT * FROM ChequeDetail, Bank WHERE BnkBankId=CqdBankId AND CqdBankId = '" & TEMP(0) & "' ORDER BY CqdCollectionNo", rsCheque)
        ElseIf UBound(TEMP) = 2 Then
           Set rsCheque = ONSCommon.SQL("SELECT * FROM ChequeDetail, Bank WHERE BnkBankId=CqdBankId AND CqdBankId = '" & TEMP(0) & "' AND CqdCollectionNo = '" & TEMP(1) & "' ORDER BY CqdCollectionNo", rsCheque)
        Else
           Set rsCheque = ONSCommon.SQL("SELECT * FROM ChequeDetail, Bank WHERE BnkBankId=CqdBankId ORDER BY CqdCollectionNo", rsCheque)
        End If
        
        While Not rsCheque.EOF And Not rsCheque.BOF
              Set LTM = lvwCheque.ListItems.Add(, rsCheque!CqdCollectionNo & "<ONS>" & rsCheque.AbsolutePosition, rsCheque!CqdCollectionNo)
              LTM.SubItems(1) = IIf(IsNull(rsCheque!BnkBankName) = True, "", rsCheque!BnkBankName)
              LTM.SubItems(2) = IIf(IsNull(rsCheque!CqdChequeNo) = True, "", rsCheque!CqdChequeNo)
              LTM.SubItems(3) = IIf(IsNull(rsCheque!CqdAmount) = True, "", rsCheque!CqdAmount)
              
              If rsCheque!CqdCollected = "H" Then
                 LTM.SubItems(4) = "Honored"
              ElseIf rsCheque!CqdCollected = "D" Then
                 LTM.SubItems(4) = "DisHonored"
              ElseIf rsCheque!CqdCollected = "S" Then
                 LTM.SubItems(4) = "Sended"
              Else
                 LTM.SubItems(4) = "Un Treate"
              End If
              
              LTM.SubItems(5) = IIf(IsNull(rsCheque!CqdTreateDate) = True, "", rsCheque!CqdTreateDate)
              rsCheque.MoveNext
        Wend
        Call ColorGrid
End Sub

Private Sub ColorGrid()
        For i = 1 To lvwCheque.ListItems.Count
            If lvwCheque.ListItems(i).ListSubItems.Item(4).Text = "Honored" Then
               lvwCheque.ListItems(i).ForeColor = vbGreen
               For j = 1 To lvwCheque.ListItems(i).ListSubItems.Count
                   lvwCheque.ListItems(i).ListSubItems.Item(j).ForeColor = vbGreen
               Next j
            ElseIf lvwCheque.ListItems(i).ListSubItems.Item(4).Text = "DisHonored" Then
               lvwCheque.ListItems(i).ForeColor = vbRed
               For j = 1 To lvwCheque.ListItems(i).ListSubItems.Count
                   lvwCheque.ListItems(i).ListSubItems.Item(j).ForeColor = vbRed
               Next j
            ElseIf lvwCheque.ListItems(i).ListSubItems.Item(4).Text = "Sended" Then
               lvwCheque.ListItems(i).ForeColor = vbBlue
               For j = 1 To lvwCheque.ListItems(i).ListSubItems.Count
                   lvwCheque.ListItems(i).ListSubItems.Item(j).ForeColor = vbBlue
               Next j
            Else
               lvwCheque.ListItems(i).ForeColor = vbBlack
               For j = 1 To lvwCheque.ListItems(i).ListSubItems.Count
                   lvwCheque.ListItems(i).ListSubItems.Item(j).ForeColor = vbBlack
               Next j
            End If
        Next i
End Sub
