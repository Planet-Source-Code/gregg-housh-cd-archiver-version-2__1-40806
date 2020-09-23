VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "mscomctl.ocx"
Begin VB.Form frmMain 
   Caption         =   "CDArchive"
   ClientHeight    =   7485
   ClientLeft      =   165
   ClientTop       =   855
   ClientWidth     =   11280
   Icon            =   "frm1.frx":0000
   LinkTopic       =   "Form1"
   ScaleHeight     =   7485
   ScaleWidth      =   11280
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdCompress 
      Caption         =   "Compress Database"
      Height          =   255
      Left            =   9360
      TabIndex        =   8
      Top             =   420
      Width           =   1815
   End
   Begin MSComctlLib.ProgressBar pb 
      Height          =   375
      Left            =   6240
      TabIndex        =   7
      Top             =   0
      Width           =   4985
      _ExtentX        =   8784
      _ExtentY        =   661
      _Version        =   393216
      BorderStyle     =   1
      Appearance      =   0
      Scrolling       =   1
   End
   Begin MSComctlLib.ListView lvSearch 
      Height          =   4335
      Left            =   4680
      TabIndex        =   4
      Top             =   720
      Width           =   6615
      _ExtentX        =   11668
      _ExtentY        =   7646
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      Icons           =   "ImageList1"
      SmallIcons      =   "ImageList1"
      ForeColor       =   -2147483640
      BackColor       =   -2147483643
      Appearance      =   1
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Key             =   "CD"
         Object.Tag             =   "CD"
         Text            =   "CD"
         Object.Width           =   2646
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Key             =   "Results"
         Object.Tag             =   "Results"
         Text            =   "Results"
         Object.Width           =   7938
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Key             =   "FullPath"
         Object.Tag             =   "FullPath"
         Text            =   "FullPath"
         Object.Width           =   7938
      EndProperty
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add CD"
      Height          =   375
      Left            =   4800
      TabIndex        =   1
      Top             =   0
      Width           =   1335
   End
   Begin VB.TextBox txtSearch 
      Height          =   285
      Left            =   4680
      TabIndex        =   2
      Top             =   405
      Width           =   3495
   End
   Begin MSComctlLib.ImageList ImageList1 
      Left            =   1080
      Top             =   2760
      _ExtentX        =   1005
      _ExtentY        =   1005
      BackColor       =   -2147483643
      ImageWidth      =   16
      ImageHeight     =   16
      MaskColor       =   12632256
      _Version        =   393216
      BeginProperty Images {2C247F25-8591-11D1-B16A-00C0F0283628} 
         NumListImages   =   8
         BeginProperty ListImage1 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":0442
            Key             =   ""
         EndProperty
         BeginProperty ListImage2 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":0894
            Key             =   ""
         EndProperty
         BeginProperty ListImage3 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":098E
            Key             =   ""
         EndProperty
         BeginProperty ListImage4 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":0A88
            Key             =   ""
         EndProperty
         BeginProperty ListImage5 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":0B82
            Key             =   ""
         EndProperty
         BeginProperty ListImage6 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":111C
            Key             =   ""
         EndProperty
         BeginProperty ListImage7 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":16B6
            Key             =   ""
         EndProperty
         BeginProperty ListImage8 {2C247F27-8591-11D1-B16A-00C0F0283628} 
            Picture         =   "frm1.frx":1C50
            Key             =   ""
         EndProperty
      EndProperty
   End
   Begin MSComctlLib.TreeView lvCDS 
      Height          =   7455
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   4575
      _ExtentX        =   8070
      _ExtentY        =   13150
      _Version        =   393217
      Indentation     =   265
      LabelEdit       =   1
      LineStyle       =   1
      Sorted          =   -1  'True
      Style           =   7
      HotTracking     =   -1  'True
      SingleSel       =   -1  'True
      ImageList       =   "ImageList1"
      Appearance      =   1
   End
   Begin VB.CommandButton cmdSearch 
      Caption         =   "Search"
      Height          =   255
      Left            =   8280
      TabIndex        =   3
      Top             =   420
      Width           =   975
   End
   Begin VB.TextBox txtInformation 
      BackColor       =   &H8000000F&
      BorderStyle     =   0  'None
      Height          =   525
      Left            =   4680
      Locked          =   -1  'True
      MousePointer    =   1  'Arrow
      TabIndex        =   6
      Top             =   5085
      Width           =   6495
   End
   Begin VB.Label lblInformation 
      Height          =   1755
      Left            =   4680
      TabIndex        =   5
      Top             =   5640
      Width           =   6525
   End
   Begin VB.Menu mnuMain 
      Caption         =   "&Main"
      Begin VB.Menu mnuSideBar1 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Main|FONT:Tahoma|FCOLOR:&H00FFFFFF&|BCOLOR:&H0051B6F2&|FSIZE:12|GRADIENT}"
      End
      Begin VB.Menu mnuMainCategories 
         Caption         =   "{IMG:5}&Category Management"
      End
      Begin VB.Menu mnuMainSep1 
         Caption         =   "-"
      End
      Begin VB.Menu mnuMainExit 
         Caption         =   "{IMG:6}&Exit"
      End
   End
   Begin VB.Menu mnuHelp 
      Caption         =   "&Help"
      Begin VB.Menu mnuSideBar2 
         Caption         =   "{SIDEBAR:TEXT|CAPTION:Help|FONT:Tahoma|FCOLOR:&H00FFFFFF&|BCOLOR:&H0051B6F2&|FSIZE:12|GRADIENT}"
      End
      Begin VB.Menu mnuHelpHelp 
         Caption         =   "{IMG:8}&Help"
      End
      Begin VB.Menu mnuSep2 
         Caption         =   "-"
      End
      Begin VB.Menu mnuHelpAbout 
         Caption         =   "{IMG:7}&About"
      End
   End
   Begin VB.Menu mnuAddCd 
      Caption         =   "mnuAddCD"
      Visible         =   0   'False
      Begin VB.Menu mnuCancel 
         Caption         =   "Cancel"
      End
      Begin VB.Menu sep01 
         Caption         =   "-"
      End
      Begin VB.Menu mnuDrive 
         Caption         =   "mnuDrive"
         Index           =   0
      End
   End
   Begin VB.Menu mnuCDOptions 
      Caption         =   "mnuCDOptions"
      Visible         =   0   'False
      Begin VB.Menu mnuInfo2 
         Caption         =   ""
      End
      Begin VB.Menu mnuInfo 
         Caption         =   ""
      End
      Begin VB.Menu fdshg 
         Caption         =   "-"
      End
      Begin VB.Menu mnuRemCD 
         Caption         =   "Remove CD From Archive"
      End
      Begin VB.Menu mnuReassignCD 
         Caption         =   "Re-Assign to another category"
      End
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Dim mstrFileName As String
Dim mobjRSFiles As New Recordset
Dim mobjRSCat As New Recordset
Dim mobjRSCDs As New Recordset
Dim mstrToClick As String
'Dim mobjLastNode As Node
Dim mobjLastItem As ListItem

Public Sub EmptyNode(ByRef pobjNode As Node)
    On Local Error Resume Next
    
    Do While pobjNode.Children <> 0
        pobjNode.Expanded = False
        EmptyNode pobjNode.Child
        lvCDS.Nodes.Remove pobjNode.Child.Key
    Loop
End Sub

Private Sub lvCDS_MouseDown(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim pobjNode As Node
    Dim plngLoop As Long
    Dim pobjNode2 As Node
    Dim pobjNewNode As Node
    
    Set pobjNode = lvCDS.HitTest(x, y)
    
    ' check to see that a node was clicked on
    If Not pobjNode Is Nothing Then
        ' test for \, if it doesnt have one its a file, if it does, its a directory.
        If InStr(1, pobjNode.Key, "\") = 0 Then
            ' if it was a left click, button = 1
            If Button = 1 Then
                If Left(pobjNode.Tag, 4) <> "CAT-" Then
                    FillNode pobjNode
                Else
                    pobjNode.Expanded = True
                End If
            ElseIf Button = 2 Then
                If Left(pobjNode.Tag, 4) <> "CAT-" Then
                    mnuInfo.Caption = "CDID: " & pobjNode.Text
                    mnuInfo2.Caption = "CDLabel: " & pobjNode.Tag
                    mnuInfo.Tag = pobjNode.Key
                    PopupMenu mnuCDOptions, , x, y, mnuRemCD
                End If
            End If
        Else
            lblInformation.Caption = pobjNode.Tag
        End If
    Else
        lvCDS.ToolTipText = ""
    End If
End Sub

Private Sub lvCDS_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim pobjNode As Node

    Set pobjNode = lvCDS.HitTest(x, y)
    If Not pobjNode Is Nothing Then
        On Local Error Resume Next
        If InStr(1, pobjNode.Key, "\") = 0 Then
            If Left(pobjNode.Tag, 4) <> "CAT-" Then
                lvCDS.ToolTipText = "CDLabel: " & pobjNode.Tag
            Else
                lvCDS.ToolTipText = pobjNode.Text
            End If
        Else
            lblInformation = pobjNode.Tag
        End If
    Else
        lvCDS.ToolTipText = ""
    End If

End Sub

Private Sub cmdAdd_Click()

    Dim pstrDrives() As String
    Dim plngLoop As Long

    Me.Enabled = False

    pstrDrives = GetDriveListing()
    mnuDrive(0).Caption = pstrDrives(0) & ":" & GetDriveSerial(pstrDrives(0))
    mnuDrive(0).Tag = pstrDrives(0)

    If GetDriveSerial(mnuDrive(0).Tag) = "0" Then mnuDrive(0).Visible = False
    If Len(pstrDrives(0)) = 0 Then mnuDrive(0).Visible = False
    
    For plngLoop = 1 To UBound(pstrDrives)
        DoEvents
        
        Load mnuDrive(plngLoop)
        mnuDrive(plngLoop).Caption = pstrDrives(plngLoop) & ":" & GetDriveSerial(pstrDrives(plngLoop))
        mnuDrive(plngLoop).Tag = pstrDrives(plngLoop)
        
        If Len(pstrDrives(plngLoop)) <> 0 Then mnuDrive(plngLoop).Visible = True
        If GetDriveSerial(mnuDrive(plngLoop).Tag) = "0" Then mnuDrive(plngLoop).Visible = False
        
    Next plngLoop

    PopupMenu mnuAddCd, , cmdAdd.Left, cmdAdd.Top + cmdAdd.Height

    For plngLoop = 1 To mnuDrive.UBound
        DoEvents
        Unload mnuDrive(plngLoop)
    Next plngLoop

    Me.Enabled = True
    
End Sub

Private Sub cmdSearch_Click()

    Dim pstrQuery As String
    Dim pcurTotal As Currency
    Dim pobjItem As ListItem

    Enabled = False
    lvSearch.ListItems.Clear
    pstrQuery = "SELECT f.*,c.* FROM Files f, CDs c WHERE Name LIKE ""%" & txtSearch.Text & "%"" AND c.CDID=f.CDID;"
    mobjRSFiles.Open pstrQuery, gobjDB, adOpenStatic, adLockReadOnly

    On Local Error Resume Next
    pb.Max = mobjRSFiles.RecordCount
    txtInformation = "Found " & mobjRSFiles.RecordCount & " results."
    Do While Not mobjRSFiles.EOF
        txtInformation = "Gathering search results..." & pb.Value + 1 & " of " & pb.Max & ", " & FormatPercent((pb.Value + 1) / pb.Max, 0) & ", " & MakeSize(CDbl(pcurTotal))
        pb.Value = mobjRSFiles.AbsolutePosition
        Set pobjItem = lvSearch.ListItems.Add(, mobjRSFiles("f.CDID") & mobjRSFiles("FullPath"), mobjRSFiles("CDName"), mobjRSFiles("Type") + 3, mobjRSFiles("Type") + 3)
        pobjItem.SubItems(1) = mobjRSFiles("Name")
        pobjItem.SubItems(2) = mobjRSFiles("FullPath")
        pobjItem.Tag = "File Name: " & mobjRSFiles("Name") & vbCrLf
        pobjItem.Tag = pobjItem.Tag & "Date: " & mobjRSFiles("Date") & vbCrLf
        If mobjRSFiles("Type") = 1 Then
            pobjItem.Tag = pobjItem.Tag & "Size: " & MakeSize(mobjRSFiles("Size")) & vbCrLf
            pobjItem.Tag = pobjItem.Tag & "Type: File" & vbCrLf
        Else
            pobjItem.Tag = pobjItem.Tag & "Type: Directory" & vbCrLf
        End If
     pobjItem.Tag = pobjItem.Tag & "Path: " & mobjRSFiles("FullPath") & vbCrLf
     pobjItem.Tag = pobjItem.Tag & "FileID: " & mobjRSFiles("ID") & vbCrLf
     pobjItem.Tag = pobjItem.Tag & "CDID: " & mobjRSFiles("c.CDID") & vbCrLf
     pobjItem.Tag = pobjItem.Tag & "CD Serial: " & mobjRSFiles("CDSerial") & vbCrLf
     pobjItem.Tag = pobjItem.Tag & "CD Label: " & mobjRSFiles("CDLabel") & vbCrLf
     pobjItem.Key = mobjRSFiles("c.CDID") & mobjRSFiles("FullPath")
     
     pcurTotal = pcurTotal + mobjRSFiles("Size")
    
     If lvSearch.ListItems.Count Mod 200 = 0 Then DoEvents
     
     mobjRSFiles.MoveNext
    Loop

    pb.Value = 0
    mobjRSFiles.Close
    Enabled = True
End Sub

Private Sub cmdCompress_Click()
    Dim pobjJet As New JetEngine
    
    gobjDB.Close
    Kill mstrFileName & ".bak"
    Name mstrFileName As mstrFileName & ".bak"
    pobjJet.CompactDatabase "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mstrFileName & ".bak", _
    "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & mstrFileName
    gobjDB.Open
End Sub

Private Sub Form_Load()
    Dim pstrDir As String
    
    pstrDir = App.Path
    If Left$(pstrDir, 1) <> "\" Then pstrDir = pstrDir & "\"
    gobjDB.Open "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=" & pstrDir & "archive.mdb"
    mstrFileName = pstrDir & "archive.mdb"
    RefreshCDs

    SetMenus Me.hwnd, ImageList1
    
End Sub

Private Function DoesNodeExsist(ByVal pstrKey As String) As Boolean
    Dim mstrTest As String
    
    DoesNodeExsist = False
    On Local Error GoTo errhand
    mstrTest = TypeName(lvCDS.Nodes(pstrKey))
    DoesNodeExsist = True
    Exit Function

errhand:
End Function

Public Function RefreshCDs()
    On Local Error Resume Next
    
    Dim pobjCatNode As Node
    Dim pobjNode As Node
    Dim pstrTotal As String
    
    lvCDS.Nodes.Clear
    mobjRSCat.Open "SELECT * FROM Categories ORDER BY CategoryName", gobjDB, adOpenForwardOnly, adLockReadOnly
    
    Do While Not mobjRSCat.EOF
        Set pobjCatNode = lvCDS.Nodes.Add(, tvwAutomatic, mobjRSCat("CategoryName"), mobjRSCat("CategoryName"), 5)
        pobjCatNode.Tag = "CAT-" & mobjRSCat("CategoryID")
        
        mobjRSCDs.Open "SELECT * FROM CDs WHERE CategoryID = " & mobjRSCat("CategoryID") & " ORDER BY CDName", gobjDB, adOpenForwardOnly, adLockReadOnly
        Do While Not mobjRSCDs.EOF
            Set pobjNode = lvCDS.Nodes.Add(pobjCatNode, tvwChild, mobjRSCDs("CDID"), mobjRSCDs("CDName") & " [" & mobjRSCDs("CDSerial") & "]", 1)
            pobjNode.Tag = mobjRSCDs("CDLabel")
            mobjRSCDs.MoveNext
        Loop
        mobjRSCDs.Close
        mobjRSCat.MoveNext
    Loop
    
    mobjRSCDs.Open "SELECT count(*) FROM CDs", gobjDB, adOpenForwardOnly, adLockReadOnly
    pstrTotal = mobjRSCDs(0).Value
    mobjRSCDs.Close
    mobjRSCat.Close
    mobjRSCDs.Open "DBInfo"
    Caption = "CDArchive [" & pstrTotal & " discs] [" & MakeSize(mobjRSCDs("AllSize")) & " bytes in " & mobjRSCDs("FileCount") & " files and directories]"
    mobjRSCDs.Close
End Function

Private Sub Form_Unload(Cancel As Integer)
    On Error Resume Next
    ReleaseMenus Me.hwnd
    Set mobjRSFiles = Nothing
    Set mobjRSCDs = Nothing
End Sub

Private Sub lvSearch_ItemClick(ByVal Item As MSComctlLib.ListItem)
    On Error Resume Next
    Dim pobjNode As Node
    FillNode lvCDS.Nodes(GetCD(Item.Key))
    Set pobjNode = lvCDS.Nodes(Item.Key)
    lvCDS.SetFocus
    pobjNode.Expanded = True
    pobjNode.Selected = True
    pobjNode.EnsureVisible
End Sub

Private Sub lvSearch_MouseMove(Button As Integer, Shift As Integer, x As Single, y As Single)
    Dim pobjItem As ListItem
    
    Set pobjItem = lvSearch.HitTest(x, y)
    
    If Not pobjItem Is Nothing Then
    
        On Local Error Resume Next

        lvSearch.ToolTipText = "CDID: " & Mid(pobjItem.Key, 1, InStr(1, pobjItem.Key, "\") - 1)
        lblInformation = pobjItem.Tag
        If DoesNodeExsist(pobjItem.Key) = True Then
            lvCDS.SetFocus
            lvCDS.Nodes(pobjItem.Key).Expanded = True
            lvCDS.Nodes(pobjItem.Key).Selected = True
            lvCDS.Nodes(pobjItem.Key).EnsureVisible
        End If
    Else
        lvSearch.ToolTipText = ""
    End If
End Sub

Private Sub mnuDrive_Click(Index As Integer)
    Dim pstrName As String
    Dim plngCat As Long
    
    Set mobjRSCDs = gobjDB.Execute("SELECT COUNT(*) FROM CDs WHERE CDSerial=" & GetDriveSerial(mnuDrive(Index).Tag) & ";")
    If mobjRSCDs.Fields(0).Value <> 0 Then
        MsgBox "This CD has already been cataloged.", vbCritical, "Error"
        mobjRSCDs.Close
        Exit Sub
    End If
    
    plngCat = frmCategories.Display()
    Unload frmCategories
    
    If plngCat <> 0 Then
retry:
        mobjRSCDs.Close
        pstrName = InputBox$("Enter a simple name for this CD", "Name")
        If Len(pstrName) > 0 Then
            Set mobjRSCDs = gobjDB.Execute("SELECT COUNT(*) FROM CDs WHERE CDName='" & pstrName & "' and CategoryID=" & plngCat & ";")
            If mobjRSCDs.Fields(0).Value <> 0 Then
                MsgBox "A CD in this category with that name has already been cataloged.", vbCritical, "Error"
                GoTo retry
            End If
            
            mobjRSCDs.Close
            mobjRSCDs.Open "CDs", gobjDB, adOpenDynamic, adLockOptimistic
            mobjRSCDs.AddNew
            mobjRSCDs("CDSerial") = GetDriveSerial(mnuDrive(Index).Tag)
            mobjRSCDs("CDLabel") = GetVolumeLabel(mnuDrive(Index).Tag)
            If mobjRSCDs("CDLabel") = "" Then
                mobjRSCDs("CDLabel") = "[No Label]"
            End If
            mobjRSCDs("CategoryID") = plngCat
            mobjRSCDs.Update
            mobjRSCDs.Close
            Set mobjRSCDs = gobjDB.Execute("SELECT CDID FROM CDs WHERE CDSerial=" & GetDriveSerial(mnuDrive(Index).Tag) & ";")
            AddFiles mobjRSCDs("CDID"), mnuDrive(Index).Tag, pstrName
            Set mobjRSCDs = gobjDB.Execute("UPDATE CDs SET CDName='" & pstrName & "' WHERE CDSerial=" & GetDriveSerial(mnuDrive(Index).Tag) & ";")
            RefreshCDs
        Else
            MsgBox "No name given, action canceled.", vbCritical, "Canceled"
        End If
    Else
        MsgBox "Cant add without a category.", vbCritical, "Error"
    End If
End Sub

Private Sub AddFiles(ByVal pstrcdGUID As String, ByVal pstrDriveRoot As String, ByVal pstrName As String)
    Dim plngDir As Long
    Dim plngLoop As Long
    Dim pobjDirQUE()
    Dim pobjDirInfo() As DirInfo

    
    ReDim pobjDirQUE(0 To 0)
    Enabled = False
    txtInformation.Text = "Reading CD..."
    mobjRSCDs.Close
    mobjRSFiles.Open "Files", gobjDB, adOpenDynamic, adLockOptimistic
    pobjDirQUE(0) = pstrDriveRoot
    plngDir = 0

    Do While (UBound(pobjDirQUE) <> plngDir - 1)
        DoEvents
        Call EnumDirs("*", pobjDirQUE(plngDir), pobjDirInfo())
        plngDir = plngDir + 1
        For plngLoop = 0 To UBound(pobjDirInfo)
        If pobjDirInfo(plngLoop).FullPath = "" Then GoTo cowboy
            If pobjDirInfo(plngLoop).fSize = "-" Then
                ReDim Preserve pobjDirQUE(LBound(pobjDirQUE) To UBound(pobjDirQUE) + 1)
                pobjDirQUE(UBound(pobjDirQUE)) = pobjDirInfo(plngLoop).FullPath & "\"
           
                mobjRSFiles.AddNew
                mobjRSFiles("CDID") = pstrcdGUID
                mobjRSFiles("Name") = pobjDirInfo(plngLoop).fNAME
                mobjRSFiles("Date") = pobjDirInfo(plngLoop).fDATE
                mobjRSFiles("FullPath") = Mid(pobjDirInfo(plngLoop).FullPath, 3)
                mobjRSFiles("Type") = 0
                mobjRSFiles("Size") = 0
                mobjRSFiles.Update
            Else
                'file
                mobjRSFiles.AddNew
                mobjRSFiles("CDID") = pstrcdGUID
                mobjRSFiles("Name") = pobjDirInfo(plngLoop).fNAME
                mobjRSFiles("Date") = pobjDirInfo(plngLoop).fDATE
                mobjRSFiles("FullPath") = Mid(pobjDirInfo(plngLoop).FullPath, 3)
                mobjRSFiles("Type") = 1
                mobjRSFiles("Size") = pobjDirInfo(plngLoop).fSize
                mobjRSFiles.Update
            End If
cowboy:
        DoEvents
        txtInformation.Text = "Reading CD..." & pobjDirInfo(plngLoop).FullPath
        Next plngLoop
        ReDim pobjDirInfo(0 To 0)
    Loop
    mobjRSFiles.Close
    lblInformation = "Reading CD...Done.."
    lblInformation = lblInformation & "Label this CD: " & pstrName & vbCrLf & vbCrLf
    lblInformation = lblInformation & plngLoop & " files and directories read."
    Enabled = True
End Sub

Private Function MakeSize(ByVal pdblBytes As Double) As String
    If pdblBytes < 1024 Then
        MakeSize = FormatNumber(pdblBytes, 0, 0, 0, True) & " b"
        Exit Function
    End If
    
    If (pdblBytes / 1024) < 1024 Then
        MakeSize = FormatNumber((pdblBytes / 1024), 2, 0, 0, True) & " KB"
        Exit Function
    End If
    If ((pdblBytes / 1024) / 1024) < 1024 Then
        MakeSize = FormatNumber(((pdblBytes / 1024) / 1024), 2, 0, 0, True) & " MB"
        Exit Function
    End If
    If (((pdblBytes / 1024) / 1024) / 1024) < 1024 Then
        MakeSize = FormatNumber((((pdblBytes / 1024) / 1024) / 1024), 2, 0, 0, True) & " GB"
        Exit Function
    End If
End Function

Private Sub mnuHelpAbout_Click()
    frmAbout.Show
End Sub

Private Sub mnuHelpHelp_Click()
    MsgBox "No help file yet.", vbInformation, "Help"
End Sub

Private Sub mnuMainCategories_Click()
    frmCategoryManagement.Display
    RefreshCDs
End Sub

Private Sub mnuMainExit_Click()
    Unload Me
End Sub

Private Sub mnuReassignCD_Click()
    Dim plngCat As Long
    
        plngCat = frmCategories.Display()
        Unload frmCategories
        If plngCat <> 0 Then
            Set mobjRSCDs = gobjDB.Execute("UPDATE CDs SET CategoryID=" & plngCat & " WHERE CDID='" & mnuInfo.Tag & "';")
            RefreshCDs
        Else
            MsgBox "You did not select a category for re-assignment.", vbCritical, "Error"
        End If
 
End Sub

Private Sub mnuRemCD_Click()
    If MsgBox("Removing a CD is a irreversable action, the CD Can Be re-added but only if you re-add it manually!" & vbCrLf & "Are you sure you want to delete it?", vbYesNo Or vbInformation, "Yes/No?") = vbYes Then
        gobjDB.Execute "DELETE FROM CDs WHERE CDID='" & mnuInfo.Tag & "'"
        gobjDB.Execute "DELETE FROM Files WHERE CDID='" & mnuInfo.Tag & "'"
        RefreshCDs
    End If
End Sub

Private Sub txtSearch_KeyUp(KeyCode As Integer, Shift As Integer)
    If KeyCode = 13 And Len(txtSearch.Text) > 0 Then
        cmdSearch_Click
    End If
End Sub

Private Function FillNode(ByRef pobjNode As Node)
    Dim plngLoop As Long
    Dim pobjNode2 As Node
    Dim pobjNewNode As Node
   
   ' if it is empty, we need to fill it
   If pobjNode.Children = 0 Then
       ' make sure its not expanded while we fill
       pobjNode.Expanded = False
       ' disable evertyhing while we fill
       lblInformation.Caption = ""
       Enabled = False
       lvCDS.Enabled = False
              
       Set pobjNode2 = lvCDS.Nodes(pobjNode.Key)

       ' get a list of all the files and directories for this cd.
       mobjRSFiles.Open "SELECT * FROM Files WHERE CDID='" & pobjNode.Key & "' Order by FullPath;", gobjDB, adOpenStatic, adLockReadOnly
       pb.Max = mobjRSFiles.RecordCount
       
       ' loop through all the files and directories adding them to the treeview
       Do While Not mobjRSFiles.EOF
           pb.Value = mobjRSFiles.AbsolutePosition
           txtInformation.Text = "Loading directory structure..." & pb.Value & " of " & pb.Max & ", " & FormatPercent(pb.Value / pb.Max, 0)
           Select Case mobjRSFiles("Type")
               Case 0
                   'directory
                   If DoesNodeExsist(pobjNode.Key & GetDir(mobjRSFiles("FullPath"))) = False Then
                       Set pobjNewNode = lvCDS.Nodes.Add(pobjNode2, tvwChild, pobjNode.Key & mobjRSFiles("FullPath"), mobjRSFiles("Name"), 3)
                       pobjNewNode.ExpandedImage = 2
                   Else
                       Set pobjNewNode = lvCDS.Nodes.Add(pobjNode.Key & GetDir(mobjRSFiles("FullPath")), tvwChild, pobjNode.Key & mobjRSFiles("FullPath"), mobjRSFiles("Name"), 3)
                       pobjNewNode.ExpandedImage = 2
                   End If
               Case 1
                   'file
                   If DoesNodeExsist(pobjNode.Key & GetDir(mobjRSFiles("FullPath"))) = False Then
                       Set pobjNewNode = lvCDS.Nodes.Add(pobjNode2, tvwChild, pobjNode.Key & mobjRSFiles("FullPath"), mobjRSFiles("Name"), 4)
                   Else
                       Set pobjNewNode = lvCDS.Nodes.Add(pobjNode.Key & GetDir(mobjRSFiles("FullPath")), tvwChild, pobjNode.Key & mobjRSFiles("FullPath"), mobjRSFiles("Name"), 4)
                   End If
           End Select
           
           ' build the information for the lblInformation display.
           pobjNewNode.Tag = "File Name: " & mobjRSFiles("Name") & vbCrLf
           pobjNewNode.Tag = pobjNewNode.Tag & "Date: " & mobjRSFiles("Date") & vbCrLf
           
           If mobjRSFiles("Type") = 1 Then
                 pobjNewNode.Tag = pobjNewNode.Tag & "Size: " & MakeSize(mobjRSFiles("Size")) & vbCrLf
                 pobjNewNode.Tag = pobjNewNode.Tag & "Type: File" & vbCrLf
           Else
                 pobjNewNode.Tag = pobjNewNode.Tag & "Type: Directory" & vbCrLf
           End If
           pobjNewNode.Tag = pobjNewNode.Tag & "Path: " & mobjRSFiles("FullPath") & vbCrLf
           pobjNewNode.Tag = pobjNewNode.Tag & "FileID: " & mobjRSFiles("ID") & vbCrLf
           pobjNewNode.Tag = pobjNewNode.Tag & "CDID: " & pobjNode.Key & vbCrLf
           pobjNewNode.Tag = pobjNewNode.Tag & "CD Serial: " & pobjNode.Tag & vbCrLf
           pobjNewNode.Tag = pobjNewNode.Tag & "CD Label: " & pobjNode.Text & vbCrLf
           
           mobjRSFiles.MoveNext
           ' allow update of the display during long loops
           If lvCDS.Nodes.Count Mod 200 = 0 Then DoEvents
       Loop

       mobjRSFiles.Close
       pb.Value = 0
       pobjNode.Expanded = True
       Enabled = True
       lvCDS.Enabled = True
       txtInformation.Text = ""
   End If
End Function

Private Function GetCD(ByVal pstrKey As String)
    If InStr(1, pstrKey, "\") <> 0 Then
        GetCD = Left(pstrKey, InStr(1, pstrKey, "\") - 1)
    Else
        GetCD = pstrKey
    End If
End Function
