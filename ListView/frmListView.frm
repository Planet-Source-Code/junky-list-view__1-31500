VERSION 5.00
Object = "{831FDD16-0C5C-11D2-A9FC-0000F8754DA1}#2.0#0"; "MSCOMCTL.OCX"
Begin VB.Form frmListView 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "List View Tutorial"
   ClientHeight    =   5025
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4695
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   5025
   ScaleWidth      =   4695
   StartUpPosition =   2  'CenterScreen
   Begin VB.Frame Frame2 
      Height          =   1215
      Left            =   120
      TabIndex        =   8
      Top             =   3720
      Width           =   4455
      Begin VB.CommandButton Command4 
         Caption         =   "E&xit"
         Height          =   315
         Left            =   1560
         TabIndex        =   12
         Top             =   720
         Width           =   1335
      End
      Begin VB.CommandButton Command3 
         Caption         =   "A&bout"
         Height          =   315
         Left            =   3000
         TabIndex        =   11
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command2 
         Caption         =   "&Delete Item"
         Height          =   315
         Left            =   1560
         TabIndex        =   10
         Top             =   240
         Width           =   1335
      End
      Begin VB.CommandButton Command1 
         Caption         =   "&Add Item"
         Height          =   315
         Left            =   120
         TabIndex        =   9
         Top             =   240
         Width           =   1335
      End
   End
   Begin MSComctlLib.ListView lstListing 
      Height          =   2175
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4455
      _ExtentX        =   7858
      _ExtentY        =   3836
      View            =   3
      LabelEdit       =   1
      Sorted          =   -1  'True
      LabelWrap       =   -1  'True
      HideSelection   =   -1  'True
      FullRowSelect   =   -1  'True
      GridLines       =   -1  'True
      _Version        =   393217
      ForeColor       =   -2147483640
      BackColor       =   16777215
      BorderStyle     =   1
      Appearance      =   1
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      NumItems        =   3
      BeginProperty ColumnHeader(1) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         Text            =   "Name"
         Object.Width           =   4939
      EndProperty
      BeginProperty ColumnHeader(2) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   1
         Text            =   "Age"
         Object.Width           =   970
      EndProperty
      BeginProperty ColumnHeader(3) {BDD1F052-858B-11D1-B16A-00C0F0283628} 
         SubItemIndex    =   2
         Text            =   "Sex"
         Object.Width           =   1323
      EndProperty
   End
   Begin VB.Frame Frame1 
      Height          =   1455
      Left            =   120
      TabIndex        =   1
      Top             =   2280
      Width           =   4455
      Begin VB.ComboBox Combo1 
         Height          =   315
         ItemData        =   "frmListView.frx":0000
         Left            =   960
         List            =   "frmListView.frx":000A
         Style           =   2  'Dropdown List
         TabIndex        =   4
         Top             =   960
         Width           =   1575
      End
      Begin VB.TextBox Text2 
         Height          =   285
         Left            =   960
         TabIndex        =   3
         Top             =   600
         Width           =   735
      End
      Begin VB.TextBox Text1 
         Height          =   285
         Left            =   960
         TabIndex        =   2
         Top             =   240
         Width           =   2775
      End
      Begin VB.Label Label3 
         AutoSize        =   -1  'True
         Caption         =   "Sex"
         Height          =   195
         Left            =   240
         TabIndex        =   7
         Top             =   960
         Width           =   270
      End
      Begin VB.Label Label2 
         AutoSize        =   -1  'True
         Caption         =   "Age"
         Height          =   195
         Left            =   240
         TabIndex        =   6
         Top             =   600
         Width           =   285
      End
      Begin VB.Label Label1 
         AutoSize        =   -1  'True
         Caption         =   "Name"
         Height          =   195
         Left            =   240
         TabIndex        =   5
         Top             =   240
         Width           =   420
      End
   End
End
Attribute VB_Name = "frmListView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim Index As Long
Dim DeleteIndex As Long

Private Sub Command1_Click()
    Dim ShowItem As ListItem
    
    If Not Text1.Text = "" And Not Text2.Text = "" And Not Combo1.Text = "" Then
        Index = Index + 1
        Set ShowItem = lstListing.ListItems.Add(Index, , Text1.Text)
        ShowItem.SubItems(1) = Text2.Text
        ShowItem.SubItems(2) = Combo1.Text
    Else
        Text1.SetFocus
    End If
End Sub

Private Sub Command2_Click()
    Dim ShowItem As ListItem
    
    If DeleteIndex = 0 Then
        Exit Sub
    End If
    
    lstListing.ListItems.Remove DeleteIndex
    
    DeleteIndex = 0
End Sub

Private Sub Command3_Click()
    frmAbout.Show 1
End Sub

Private Sub Command4_Click()
    End
End Sub

Private Sub lstListing_ColumnClick(ByVal ColumnHeader As MSComctlLib.ColumnHeader)
    Select Case ColumnHeader
    Case Is = "Name"
        lstListing.SortKey = 0
    Case Is = "Age"
        lstListing.SortKey = 1
    Case Is = "Sex"
        lstListing.SortKey = 2
    End Select
End Sub

Private Sub lstListing_DblClick()
    Call Command2_Click
End Sub

Private Sub lstListing_ItemClick(ByVal Item As MSComctlLib.ListItem)
    DeleteIndex = Item.Index
End Sub
