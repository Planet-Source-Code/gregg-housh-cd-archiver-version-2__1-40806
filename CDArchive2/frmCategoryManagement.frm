VERSION 5.00
Begin VB.Form frmCategoryManagement 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Select Category"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3435
   Icon            =   "frmCategoryManagement.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdDelete 
      Caption         =   "Delete"
      Height          =   495
      Left            =   1200
      TabIndex        =   3
      Top             =   2280
      Width           =   975
   End
   Begin VB.CommandButton cmdAdd 
      Caption         =   "Add"
      Height          =   495
      Left            =   120
      TabIndex        =   2
      Top             =   2280
      Width           =   975
   End
   Begin VB.ListBox lstCategories 
      Height          =   2010
      Left            =   120
      TabIndex        =   1
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdClose 
      Caption         =   "Close"
      Height          =   495
      Left            =   2280
      TabIndex        =   0
      Top             =   2280
      Width           =   975
   End
End
Attribute VB_Name = "frmCategoryManagement"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjRSCat As New Recordset

Public Sub Display()
    
    FillCategories
    
    Me.Show
    
End Sub

Private Sub cmdAdd_Click()
    
    Dim pstrCat As String
    
    pstrCat = InputBox$("Enter a simple name for this CD", "Name")

    If Len(pstrCat) > 0 Then
        If InStr(1, pstrCat, "\") = 0 Then
            Set mobjRSCat = gobjDB.Execute("SELECT COUNT(*) FROM Categories WHERE CategoryName='" & pstrCat & "';")
            If mobjRSCat.Fields(0).Value <> 0 Then
                MsgBox "A category with that name already exists.", vbCritical, "Error"
            Else
                mobjRSCat.Close
                mobjRSCat.Open "Categories", gobjDB, adOpenDynamic, adLockOptimistic
                mobjRSCat.AddNew
                mobjRSCat("CategoryName") = pstrCat
                mobjRSCat.Update
                mobjRSCat.Close
                FillCategories
            End If
        Else
            MsgBox "Illegal name.  Names cannot contain the \ character.", vbCritical, "Error"
        End If
    Else
        MsgBox "No name entered, canceling.", vbCritical, "Error"
    End If
      
End Sub

Private Sub cmdClose_Click()
    frmMain.RefreshCDs
    Unload Me
End Sub

Private Sub FillCategories()
    mobjRSCat.Open "SELECT * FROM Categories", gobjDB, adOpenForwardOnly, adLockReadOnly
    lstCategories.Clear
    Do While Not mobjRSCat.EOF
        lstCategories.AddItem mobjRSCat("CategoryName").Value
        lstCategories.ItemData(lstCategories.NewIndex) = mobjRSCat("CategoryID").Value
        mobjRSCat.MoveNext
    Loop
    
    mobjRSCat.Close
End Sub

Private Sub cmdDelete_Click()
    Dim plngCat As Long
    Dim mobjRSCatLoop As Recordset
    
    If lstCategories.ListIndex <> -1 Then
        Set mobjRSCat = gobjDB.Execute("DELETE * FROM Categories WHERE CategoryID=" & lstCategories.ItemData(lstCategories.ListIndex) & ";")
        If MsgBox("Assign all cd's from this category to another category?" & vbCrLf & "(If you do not, they will be deleted)", vbYesNo, "Re-Assign") = vbYes Then
            're-assign
retry:
            plngCat = frmCategories.Display()
            Unload frmCategories
            If plngCat <> 0 Then
                Set mobjRSCat = gobjDB.Execute("UPDATE CDs SET CategoryID=" & plngCat & " WHERE CategoryID=" & lstCategories.ItemData(lstCategories.ListIndex) & ";")
            Else
                MsgBox "You did not select a category for re-assignment." & vbCrLf & "Please select again.", vbExclamation, "Error"
                GoTo retry
            End If
        Else
            'delete
            Set mobjRSCatLoop = New Recordset
            mobjRSCatLoop.Open "SELECT * FROM CDs WHERE CategoryID=" & lstCategories.ItemData(lstCategories.ListIndex) & ";", gobjDB, adOpenStatic, adLockReadOnly
            
            Do While Not mobjRSCatLoop.EOF
                Set mobjRSCat = gobjDB.Execute("DELETE * FROM Files WHERE CDID='" & mobjRSCatLoop("CDID") & "';")
                mobjRSCatLoop.MoveNext
            Loop
            
            Set mobjRSCat = gobjDB.Execute("DELETE * FROM CDs WHERE CategoryID=" & lstCategories.ItemData(lstCategories.ListIndex) & ";")
        End If
        FillCategories
    Else
        MsgBox "No category selected", vbCritical, "Error"
    End If
    Set mobjRSCatLoop = Nothing
End Sub
