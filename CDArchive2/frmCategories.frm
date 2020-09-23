VERSION 5.00
Begin VB.Form frmCategories 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Categories"
   ClientHeight    =   2865
   ClientLeft      =   45
   ClientTop       =   435
   ClientWidth     =   3435
   Icon            =   "frmCategories.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2865
   ScaleWidth      =   3435
   StartUpPosition =   3  'Windows Default
   Begin VB.ListBox lstCategories 
      Height          =   2010
      Left            =   120
      TabIndex        =   2
      Top             =   120
      Width           =   3135
   End
   Begin VB.CommandButton cmdCancel 
      Caption         =   "Cancel"
      Height          =   495
      Left            =   1800
      TabIndex        =   1
      Top             =   2280
      Width           =   1095
   End
   Begin VB.CommandButton cmdOk 
      Caption         =   "Ok"
      Height          =   495
      Left            =   480
      TabIndex        =   0
      Top             =   2280
      Width           =   1095
   End
End
Attribute VB_Name = "frmCategories"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

Private mobjRSCat As New Recordset
Private mblnDone As Boolean
Private mlngCat As Long

Public Function Display() As Long
    
    mblnDone = False
    mlngCat = 0
    lstCategories.ListIndex = -1
    mobjRSCat.Open "SELECT * FROM Categories", gobjDB, adOpenForwardOnly, adLockReadOnly
    
    Do While Not mobjRSCat.EOF
        lstCategories.AddItem mobjRSCat("CategoryName").Value
        lstCategories.ItemData(lstCategories.NewIndex) = mobjRSCat("CategoryID").Value
        mobjRSCat.MoveNext
    Loop
    
    mobjRSCat.Close
    Me.Show
    
    Do Until mblnDone = True
        DoEvents
    Loop
    Display = mlngCat
    
    Unload Me
End Function

Private Sub cmdCancel_Click()
    mlngCat = 0
    mblnDone = True
End Sub

Private Sub cmdOk_Click()
    If mlngCat <> 0 Then
        mblnDone = True
    Else
        MsgBox "Nothing selected", vbCritical, "Error"
    End If
End Sub

Private Sub lstCategories_Click()
    If lstCategories.ListIndex <> -1 Then
        mlngCat = lstCategories.ItemData(lstCategories.ListIndex)
    End If
End Sub

Private Sub lstCategories_DblClick()
    cmdOk_Click
End Sub
