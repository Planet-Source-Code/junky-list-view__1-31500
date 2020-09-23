VERSION 5.00
Begin VB.Form frmAbout 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "List View Demo"
   ClientHeight    =   3780
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   3540
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Moveable        =   0   'False
   ScaleHeight     =   3780
   ScaleWidth      =   3540
   ShowInTaskbar   =   0   'False
   StartUpPosition =   1  'CenterOwner
   Begin VB.CommandButton cmdOK 
      Caption         =   "OK"
      Height          =   495
      Left            =   840
      TabIndex        =   0
      Top             =   3120
      Width           =   1935
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Feel free to use any part of the code for educational purpose."
      Height          =   495
      Left            =   330
      TabIndex        =   6
      Top             =   1920
      Width           =   2895
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "E-mail:"
      Height          =   255
      Left            =   510
      TabIndex        =   5
      Top             =   2640
      Width           =   495
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "This is an example on how to use List View to display items."
      Height          =   435
      Left            =   330
      TabIndex        =   4
      Top             =   1320
      Width           =   2880
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "cool_junkman@yahoo.com"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FF0000&
      Height          =   195
      Left            =   1110
      TabIndex        =   3
      Top             =   2640
      Width           =   1935
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      BackStyle       =   0  'Transparent
      Caption         =   "Created by JunKy @ JunKy Technology 2002"
      Height          =   375
      Left            =   323
      TabIndex        =   2
      Top             =   720
      Width           =   2895
   End
   Begin VB.Label Label2 
      Alignment       =   2  'Center
      AutoSize        =   -1  'True
      BackStyle       =   0  'Transparent
      Caption         =   "List View Demo"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   300
      Left            =   825
      TabIndex        =   1
      Top             =   240
      Width           =   1905
   End
End
Attribute VB_Name = "frmAbout"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdOK_Click()
    Unload Me
End Sub

Private Sub Label4_Click()
    Shell "c:\program files\internet explorer\iexplore.exe mailto:cool_junkman@yahoo.com", vbMaximizedFocus
End Sub

