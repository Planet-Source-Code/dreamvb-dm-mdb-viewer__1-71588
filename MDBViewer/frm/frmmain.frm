VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Object = "{F9043C88-F6F2-101A-A3C9-08002B2F49FB}#1.2#0"; "COMDLG32.OCX"
Begin VB.Form frmmain 
   Caption         =   "DM MDB Viewer v1.0"
   ClientHeight    =   4260
   ClientLeft      =   60
   ClientTop       =   630
   ClientWidth     =   8625
   LinkTopic       =   "Form1"
   ScaleHeight     =   4260
   ScaleWidth      =   8625
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cboTables 
      Height          =   315
      Left            =   585
      Style           =   2  'Dropdown List
      TabIndex        =   4
      Top             =   405
      Width           =   1215
   End
   Begin MSComDlg.CommonDialog CD1 
      Left            =   7395
      Top             =   3120
      _ExtentX        =   847
      _ExtentY        =   847
      _Version        =   393216
   End
   Begin MSComctlLib.ListView LstV 
      Height          =   750
      Left            =   0
      TabIndex        =   2
      Top             =   840
      Width           =   1500
      _ExtentX        =   2646
      _ExtentY        =   1323
      View            =   3
      LabelEdit       =   1
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      BorderStyle     =   1
      Appearance      =   1
      NumItems        =   0
   End
   Begin MSComctlLib.StatusBar sBar1 
      Align           =   2  'Align Bottom
      Height          =   300
      Left            =   0
      TabIndex        =   1
      Top             =   3960
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   529
      _Version        =   393216
      BeginProperty Panels {8E3867A5-8586-11D1-B16A-00C0F0283628} 
         NumPanels       =   2
         BeginProperty Panel1 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   1
            Object.Width           =   12144
         EndProperty
         BeginProperty Panel2 {8E3867AB-8586-11D1-B16A-00C0F0283628} 
            AutoSize        =   2
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   6660
      Top             =   3135
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   16711935
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   1
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frmmain.frx":0000
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.Toolbar tBar1 
      Align           =   1  'Align Top
      Height          =   360
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   8625
      _ExtentX        =   15214
      _ExtentY        =   635
      ButtonWidth     =   609
      ButtonHeight    =   582
      Appearance      =   1
      Style           =   1
      ImageList       =   "ImageList1"
      _Version        =   393216
      BeginProperty Buttons {66833FE8-8583-11D1-B16A-00C0F0283628} 
         NumButtons      =   1
         BeginProperty Button1 {66833FEA-8583-11D1-B16A-00C0F0283628} 
            Key             =   "OPEN"
            ImageIndex      =   1
         EndProperty
      EndProperty
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00FFFFFF&
      Index           =   3
      X1              =   0
      X2              =   720
      Y1              =   765
      Y2              =   765
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00808080&
      Index           =   2
      X1              =   0
      X2              =   720
      Y1              =   750
      Y2              =   750
   End
   Begin VB.Label lblTable 
      AutoSize        =   -1  'True
      Caption         =   "Tables:"
      Height          =   195
      Left            =   30
      TabIndex        =   3
      Top             =   465
      Width           =   525
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00808080&
      Index           =   1
      X1              =   0
      X2              =   720
      Y1              =   360
      Y2              =   360
   End
   Begin VB.Line lnTop 
      BorderColor     =   &H00FFFFFF&
      Index           =   0
      X1              =   0
      X2              =   720
      Y1              =   375
      Y2              =   375
   End
   Begin VB.Menu mnuFile 
      Caption         =   "&File"
      Begin VB.Menu mnuExit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu mnuAbout 
      Caption         =   "&About"
   End
End
Attribute VB_Name = "frmmain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private db As Database
Private rc As Recordset
Private td As TableDef
Private tField As Field
Private DBOpen As Boolean

Private Sub GetFieldData(ByVal TableName As String, TListView As ListView)
Dim fCount As Integer
On Error Resume Next

    'This sub adds the data to the listview
    
    'Clear the listview
    TListView.ListItems.Clear
    'First lets open the record set
    Set rc = db.OpenRecordset(TableName)
    
    With rc
        While Not rc.EOF
            'Add the first item
            TListView.ListItems.Add , , rc.Fields(0)
            'Add the subitems
            For fCount = 1 To (rc.Fields.Count - 1)
                TListView.ListItems(TListView.ListItems.Count).SubItems(fCount) = rc.Fields(fCount)
            Next fCount
            'Get the next record
            .MoveNext
        Wend
    End With
    
    Set rc = Nothing
    
End Sub

Private Sub GetFieldNames(ByVal TableName As String, TListView As ListView)
    'This sub adds all the field names to the listview control
    
    'Clear the headers
    TListView.ColumnHeaders.Clear
    'First we need to open the record set
    Set rc = db.OpenRecordset(TableName)
    'Now we can add the field names
    For Each tField In rc.Fields
        TListView.ColumnHeaders.Add , , tField.Name
    Next tField
    
    Set rc = Nothing
    Set tField = Nothing
End Sub

Private Sub GetTables(CboBox As ComboBox)
    CboBox.Clear
    'Get all Table names
    For Each td In db.TableDefs
        If (td.Attributes And dbSystemObject) Then
        Else
            CboBox.AddItem td.Name
        End If
    Next td
    
    Set td = Nothing
End Sub

Private Function OpenDB(ByVal Filename As String) As Boolean
On Error GoTo OpenErr:
    'Open the Database
    Set db = OpenDatabase(Filename, False)
    'Tells us the database is open
    DBOpen = True
    OpenDB = True
    Exit Function
OpenErr:
    OpenDB = False
End Function

Private Function GetDLGName() As String
On Error GoTo OpenErr:
    
    With CD1
        .CancelError = True
        .DialogTitle = "Open"
        .Filter = "Database Files(*.mdb)|*.mdb|"
        .ShowOpen
        'Return filename
        GetDLGName = .Filename
    End With
    
    Exit Function
OpenErr:
    If (Err.Number = cdlCancel) Then
        Err.Clear
    End If
End Function

Private Sub cboTables_Click()
    'Add field names
    Call GetFieldNames(cboTables.Text, LstV)
    'Add the data
    Call GetFieldData(cboTables.Text, LstV)
    'Update statusbar
    sBar1.Panels(2).Text = "Records: " & LstV.ListItems.Count
End Sub

Private Sub Form_Resize()
On Error Resume Next
    'Resize the controls
    lnTop(0).X2 = frmmain.ScaleWidth
    lnTop(1).X2 = lnTop(0).X2
    lnTop(2).X2 = lnTop(0).X2
    lnTop(3).X2 = lnTop(0).X2
    LstV.Width = lnTop(0).X2
    LstV.Height = (frmmain.ScaleHeight - sBar1.Height - LstV.Top)
End Sub

Private Sub mnuAbout_Click()
    MsgBox frmmain.Caption & vbCrLf & vbTab & "By DreamVB", vbInformation, "About"
End Sub

Private Sub mnuExit_Click()
    'Check if the database is open
    If (DBOpen) Then
        db.Close
    End If
    
    Unload frmmain
End Sub

Private Sub tBar1_ButtonClick(ByVal Button As MSComctlLib.Button)
Dim lFile As String
    Select Case Button.Key
        Case "OPEN"
            lFile = GetDLGName()
            If Len(lFile) Then
                LstV.ListItems.Clear
                sBar1.Panels(2).Text = ""
                'Load the database
                If Not OpenDB(lFile) Then
                    MsgBox "Cannot open database.", vbInformation, frmmain.Caption
                Else
                    'List the tables
                    Call GetTables(cboTables)
                    'Update statusbar
                    sBar1.Panels(1).Text = lFile
                End If
            End If
    End Select
End Sub
