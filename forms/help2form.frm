VERSION 5.00
Begin VB.Form help2form 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "User Guide"
   ClientHeight    =   3015
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8850
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "help2form.frx":0000
   ScaleHeight     =   3015
   ScaleWidth      =   8850
   StartUpPosition =   2  'CenterScreen
   Begin VB.Image Image1 
      Height          =   480
      Left            =   7800
      Picture         =   "help2form.frx":1DD3
      Stretch         =   -1  'True
      Top             =   120
      Width           =   810
   End
   Begin VB.Image Image2 
      Height          =   480
      Left            =   7800
      Picture         =   "help2form.frx":3123
      Stretch         =   -1  'True
      Top             =   720
      Width           =   810
   End
   Begin VB.Image Image3 
      Height          =   480
      Left            =   7800
      Picture         =   "help2form.frx":4277
      Stretch         =   -1  'True
      Top             =   1320
      Width           =   810
   End
   Begin VB.Image Image4 
      Height          =   480
      Left            =   7800
      Picture         =   "help2form.frx":539F
      Stretch         =   -1  'True
      Top             =   1920
      Width           =   810
   End
   Begin VB.Label Label1 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "My Sections Form will allow user to view the assigned sections. It includes forms for encoding student's grades."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   4
      Top             =   120
      Width           =   7455
   End
   Begin VB.Label Label2 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   $"help2form.frx":6073
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   3
      Top             =   720
      Width           =   7455
   End
   Begin VB.Label Label3 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "My Account Form will allow user to change his/her own personal account."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   2
      Top             =   1320
      Width           =   7455
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      BorderStyle     =   1  'Fixed Single
      Caption         =   "Help Form will allow user to view user's guide and information about the system."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   1
      Top             =   1920
      Width           =   7455
   End
   Begin VB.Label lbl_restore 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "About Form 137 and Promotion Report Generation System of Manuel S. Rojas Elementary School"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   0
      TabIndex        =   0
      Top             =   2640
      Width           =   8655
   End
End
Attribute VB_Name = "help2form"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lbl_restore_Click()
     Call load_form(aboutform, True)
End Sub
