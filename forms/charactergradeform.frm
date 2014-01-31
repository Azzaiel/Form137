VERSION 5.00
Begin VB.Form charactergradeform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Character Grade"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8805
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "charactergradeform.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   8805
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmd_cancel 
      Height          =   615
      Left            =   6840
      Picture         =   "charactergradeform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   17
      Top             =   5280
      Width           =   1095
   End
   Begin VB.ComboBox cmb_14 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C149
      Left            =   7800
      List            =   "charactergradeform.frx":1C159
      TabIndex        =   15
      Top             =   4320
      Width           =   735
   End
   Begin VB.ComboBox cmb_13 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C169
      Left            =   7800
      List            =   "charactergradeform.frx":1C179
      TabIndex        =   14
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox cmb_12 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C189
      Left            =   7800
      List            =   "charactergradeform.frx":1C199
      TabIndex        =   13
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox cmb_11 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C1A9
      Left            =   7800
      List            =   "charactergradeform.frx":1C1B9
      TabIndex        =   12
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox cmb_10 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C1C9
      Left            =   7800
      List            =   "charactergradeform.frx":1C1D9
      TabIndex        =   11
      Top             =   1920
      Width           =   735
   End
   Begin VB.ComboBox cmb_9 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C1E9
      Left            =   7800
      List            =   "charactergradeform.frx":1C1F9
      TabIndex        =   10
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox cmb_8 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C209
      Left            =   7800
      List            =   "charactergradeform.frx":1C219
      TabIndex        =   9
      Top             =   720
      Width           =   735
   End
   Begin VB.ComboBox cmb_7 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C229
      Left            =   3360
      List            =   "charactergradeform.frx":1C239
      TabIndex        =   8
      Top             =   4320
      Width           =   735
   End
   Begin VB.ComboBox cmb_6 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C249
      Left            =   3360
      List            =   "charactergradeform.frx":1C259
      TabIndex        =   7
      Top             =   3720
      Width           =   735
   End
   Begin VB.ComboBox cmb_5 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C269
      Left            =   3360
      List            =   "charactergradeform.frx":1C279
      TabIndex        =   6
      Top             =   3120
      Width           =   735
   End
   Begin VB.ComboBox cmb_4 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C289
      Left            =   3360
      List            =   "charactergradeform.frx":1C299
      TabIndex        =   5
      Top             =   2520
      Width           =   735
   End
   Begin VB.ComboBox cmb_3 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C2A9
      Left            =   3360
      List            =   "charactergradeform.frx":1C2B9
      TabIndex        =   4
      Top             =   1920
      Width           =   735
   End
   Begin VB.ComboBox cmb_2 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C2C9
      Left            =   3360
      List            =   "charactergradeform.frx":1C2D9
      TabIndex        =   3
      Top             =   1320
      Width           =   735
   End
   Begin VB.ComboBox cmb_1 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   360
      ItemData        =   "charactergradeform.frx":1C2E9
      Left            =   3360
      List            =   "charactergradeform.frx":1C2F9
      TabIndex        =   2
      Top             =   720
      Width           =   735
   End
   Begin VB.Frame Frame2 
      BackColor       =   &H00808080&
      Caption         =   "Legend"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1215
      Left            =   0
      TabIndex        =   40
      Top             =   4800
      Width           =   4335
      Begin VB.Label Label21 
         BackStyle       =   0  'Transparent
         Caption         =   "D - Needs Improvement"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   2280
         TabIndex        =   44
         Top             =   720
         Width           =   1815
      End
      Begin VB.Label Label20 
         BackStyle       =   0  'Transparent
         Caption         =   "C - Satisfactory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2280
         TabIndex        =   43
         Top             =   360
         Width           =   1815
      End
      Begin VB.Label Label19 
         BackStyle       =   0  'Transparent
         Caption         =   "B - Very Satisfactory"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   495
         Left            =   240
         TabIndex        =   42
         Top             =   720
         Width           =   1935
      End
      Begin VB.Label Label18 
         BackStyle       =   0  'Transparent
         Caption         =   "A - Outstanding"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   240
         TabIndex        =   41
         Top             =   360
         Width           =   1815
      End
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   5520
      Picture         =   "charactergradeform.frx":1C309
      Style           =   1  'Graphical
      TabIndex        =   16
      Top             =   5280
      Width           =   1095
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   615
      Left            =   0
      TabIndex        =   0
      Top             =   -120
      Width           =   9015
      Begin VB.Label lbl_period2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   7080
         TabIndex        =   47
         Top             =   240
         Width           =   1815
      End
      Begin VB.Label lbl_name2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   3120
         TabIndex        =   46
         Top             =   240
         Width           =   3015
      End
      Begin VB.Label lbl_id2 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   720
         TabIndex        =   45
         Top             =   240
         Width           =   1455
      End
      Begin VB.Label lbl_subject_name 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   5280
         TabIndex        =   25
         Top             =   600
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Name:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2400
         TabIndex        =   24
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label1 
         BackStyle       =   0  'Transparent
         Caption         =   "LRN:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   120
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "Period:"
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   6240
         TabIndex        =   23
         Top             =   240
         Width           =   975
      End
      Begin VB.Label lbl_id 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   22
         Top             =   240
         Width           =   735
      End
      Begin VB.Label lbl_name 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   21
         Top             =   600
         Width           =   735
      End
      Begin VB.Label lbl_code 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   20
         Top             =   960
         Width           =   735
      End
      Begin VB.Label lbl_sub_name 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   19
         Top             =   1320
         Width           =   735
      End
      Begin VB.Label lbl_period 
         BackStyle       =   0  'Transparent
         BeginProperty Font 
            Name            =   "MS Sans Serif"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         ForeColor       =   &H00FFFFFF&
         Height          =   255
         Left            =   2040
         TabIndex        =   18
         Top             =   1680
         Width           =   735
      End
   End
   Begin VB.Label lbl_view_attendance 
      BackStyle       =   0  'Transparent
      Caption         =   "View student's attendance report."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   -1  'True
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   5040
      TabIndex        =   48
      Top             =   4920
      Visible         =   0   'False
      Width           =   3495
   End
   Begin VB.Label Label17 
      BackStyle       =   0  'Transparent
      Caption         =   "Patriotism and Love of Country:"
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
      Left            =   4680
      TabIndex        =   39
      Top             =   4320
      Width           =   2535
   End
   Begin VB.Label Label16 
      BackStyle       =   0  'Transparent
      Caption         =   "Love of God:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   38
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label15 
      BackStyle       =   0  'Transparent
      Caption         =   "Sense of Responsibility:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   37
      Top             =   3240
      Width           =   2655
   End
   Begin VB.Label Label14 
      BackStyle       =   0  'Transparent
      Caption         =   "Promptness and Punctuality:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   36
      Top             =   2640
      Width           =   2895
   End
   Begin VB.Label Label13 
      BackStyle       =   0  'Transparent
      Caption         =   "Cleanliness and Orderliness:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   4680
      TabIndex        =   35
      Top             =   2040
      Width           =   2895
   End
   Begin VB.Label Label12 
      BackStyle       =   0  'Transparent
      Caption         =   "Industry:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   34
      Top             =   1440
      Width           =   1695
   End
   Begin VB.Label Label11 
      BackStyle       =   0  'Transparent
      Caption         =   "Self-Reliance:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   4680
      TabIndex        =   33
      Top             =   840
      Width           =   1695
   End
   Begin VB.Label Label10 
      BackStyle       =   0  'Transparent
      Caption         =   "Obedience:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   32
      Top             =   4440
      Width           =   1695
   End
   Begin VB.Label Label9 
      BackStyle       =   0  'Transparent
      Caption         =   "Sportmanship:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   31
      Top             =   3840
      Width           =   1695
   End
   Begin VB.Label Label8 
      BackStyle       =   0  'Transparent
      Caption         =   "Consideration of Others:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   30
      Top             =   3120
      Width           =   2175
   End
   Begin VB.Label Label7 
      BackStyle       =   0  'Transparent
      Caption         =   "Resourcefulness and Creativity:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   120
      TabIndex        =   29
      Top             =   2520
      Width           =   2295
   End
   Begin VB.Label Label6 
      BackStyle       =   0  'Transparent
      Caption         =   "Helpfulness and Cooperation:"
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
      Left            =   120
      TabIndex        =   28
      Top             =   1920
      Width           =   2055
   End
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Courtesy:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   27
      Top             =   1440
      Width           =   1095
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "Honesty:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   120
      TabIndex        =   26
      Top             =   840
      Width           =   1695
   End
End
Attribute VB_Name = "charactergradeform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_grade As New ADODB.Recordset
Dim sql_string As String
Public period As String
Private char_grade_rs As New ADODB.Recordset



Private Sub cmb_10_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_10.Text = ""
End Sub

Private Sub cmb_11_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_11.Text = ""
End Sub

Private Sub cmb_12_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_12.Text = ""
End Sub

Private Sub cmb_13_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_13.Text = ""
End Sub

Private Sub cmb_14_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_14.Text = ""
End Sub

Private Sub cmb_2_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_2.Text = ""
End Sub

Private Sub cmb_3_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_3.Text = ""
End Sub

Private Sub cmb_4_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_4.Text = ""
End Sub

Private Sub cmb_5_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_5.Text = ""
End Sub

Private Sub cmb_6_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_6.Text = ""
End Sub

Private Sub cmb_7_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_7.Text = ""
End Sub

Private Sub cmb_8_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_8.Text = ""
End Sub

Private Sub cmb_9_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select an item from the list."
    cmb_9.Text = ""
End Sub

Private Sub cmd_cancel_Click()
   Unload Me
End Sub
Private Sub load_grade()
      Call set_datagrid(characterencodeform.dg_grades, rs_grade, _
                                            "SELECT " _
                                                & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name, b.Honesty,b.Courtesy,b.Helpfulness_and_Cooperation,b.Resourcefulness_and_Creativity,b.Consideration_for_Others,b.Sportsmanship,b.Obedience,b.Self_Reliance,b.Industry,b.Cleanliness_and_Orderliness,b.Promptness_and_Punctuality,b.Sense_of_Responsibility,b.Love_of_God,b.Patriotism_and_Love_of_Country  FROM tbl_student a LEFT JOIN tbl_character_grade b ON a.student_id = b.ID WHERE b.section_name = '" & section & "' AND b.period = '" & lbl_period2.Caption & "'")
       Call characterencodeform.cmb_period_Click
       Unload Me
End Sub

Private Sub cmd_save_Click()
  
  Dim ans As String
  ans = MsgBox("Are you sure you want to update student's character building grade?", vbYesNo, "Character Building")
  If ans = vbNo Then
    Exit Sub
  End If
  
  char_grade_rs!Honesty = cmb_1
  char_grade_rs!Courtesy = cmb_2
  char_grade_rs!Helpfulness_and_Cooperation = cmb_3
  char_grade_rs!Resourcefulness_and_Creativity = cmb_4
  char_grade_rs!Consideration_for_Others = cmb_5
  char_grade_rs!Sportsmanship = cmb_6
  char_grade_rs!Obedience = cmb_7
  char_grade_rs!Self_Reliance = cmb_8
  char_grade_rs!Industry = cmb_9
  char_grade_rs!Cleanliness_and_Orderliness = cmb_10
  char_grade_rs!Promptness_and_Punctuality = cmb_11
  char_grade_rs!Sense_of_Responsibility = cmb_12
  char_grade_rs!Love_of_God = cmb_13
  char_grade_rs!Patriotism_and_Love_of_Country = cmb_14
  
  char_grade_rs.Update
  
  MsgBox "Record Updated", vbInformation
  
   If lbl_period2.Caption = "Final" Then
          Call mysql_select(public_rs, "SELECT * FROM tbl_attendance WHERE SY = '" & mainteacherform.cmb_sy.Text & "' and ID = '" & lbl_id2.Caption & "'")
          If public_rs.RecordCount = 0 Then
            attendanceform.lbl_id2.Caption = lbl_id2.Caption
            attendanceform.lbl_name2.Caption = lbl_name2.Caption
            Call load_form(attendanceform, True)
          End If
  End If
  
End Sub

Private Sub Form_Load()

  sql_string = "Select * " & _
               "FROM tbl_character_grade " & _
               "Where ID = '" & masterlistadvisoriesform.sel_lrn & "' " & _
               "      And Period = '" & period & "' " & _
               "      And SY = '" & mainteacherform.cmb_sy.Text & "' " & _
               "      And section_name = '" & masterlistadvisoriesform.lbl_section & "' "
               
   
  Call mysql_select(char_grade_rs, sql_string)
  
  lbl_id2 = masterlistadvisoriesform.sel_lrn
  lbl_name2 = masterlistadvisoriesform.sel_student_name
  lbl_period2 = period
  If (char_grade_rs.RecordCount > 0) Then
    cmb_1 = CommonHelper.extractStringValue(char_grade_rs!Honesty)
    cmb_2 = CommonHelper.extractStringValue(char_grade_rs!Courtesy)
    cmb_3 = CommonHelper.extractStringValue(char_grade_rs!Helpfulness_and_Cooperation)
    cmb_4 = CommonHelper.extractStringValue(char_grade_rs!Resourcefulness_and_Creativity)
    cmb_5 = CommonHelper.extractStringValue(char_grade_rs!Consideration_for_Others)
    cmb_6 = CommonHelper.extractStringValue(char_grade_rs!Sportsmanship)
    cmb_7 = CommonHelper.extractStringValue(char_grade_rs!Obedience)
    cmb_8 = CommonHelper.extractStringValue(char_grade_rs!Self_Reliance)
    cmb_9 = CommonHelper.extractStringValue(char_grade_rs!Industry)
    cmb_10 = CommonHelper.extractStringValue(char_grade_rs!Cleanliness_and_Orderliness)
    cmb_11 = CommonHelper.extractStringValue(char_grade_rs!Promptness_and_Punctuality)
    cmb_12 = CommonHelper.extractStringValue(char_grade_rs!Sense_of_Responsibility)
    cmb_13 = CommonHelper.extractStringValue(char_grade_rs!Love_of_God)
    cmb_14 = CommonHelper.extractStringValue(char_grade_rs!Patriotism_and_Love_of_Country)
  
    Call mysql_select(public_rs, "SELECT * FROM tbl_attendance WHERE ID = '" & masterlistadvisoriesform.sel_lrn & "' And SY = '" & mainteacherform.cmb_sy.Text & "' AND section_name='" & masterlistadvisoriesform.lbl_section & "'")
    If public_rs.RecordCount = 0 Then
      lbl_view_attendance.Visible = False
    Else
      lbl_view_attendance.Visible = True
    End If
  
  Else
    MsgBox "Please Encode C.E. Grade first", vbCritical
    Unload Me
  End If
    
End Sub
Private Sub oldSave()
Dim ans As String
     ans = MsgBox("Are you sure you want to update student's character building grade?", vbYesNo, "Character Building")
                    If ans = vbNo Then
                        Exit Sub
                    Else
    sql_string = "UPDATE tbl_character_grade SET Honesty='" & cmb_1.Text & "',Courtesy='" & cmb_2.Text & "',Helpfulness_and_Cooperation='" & cmb_3.Text & "',Resourcefulness_and_Creativity='" & cmb_4.Text & "',Consideration_for_Others='" & cmb_5.Text & "',Sportsmanship='" & cmb_6.Text & "',Obedience='" & cmb_7.Text & "',Self_Reliance='" & cmb_8.Text & "',Industry='" & cmb_9.Text & "',Cleanliness_and_Orderliness='" & cmb_10.Text & "',Promptness_and_Punctuality='" & cmb_11.Text & "',Sense_of_Responsibility='" & cmb_12.Text & "',Love_of_God='" & cmb_13.Text & "',Patriotism_and_Love_of_Country='" & cmb_14.Text & "' WHERE ID= '" & lbl_id2.Caption & "'AND section_name= '" & section & "' AND period= '" & lbl_period2.Caption & "'"
    Call mysql_select(charactergradeform.rs_grade, sql_string)
    MsgBox "Character grade encoded."
    End If
    If lbl_period2.Caption = "Final" Then
          Call mysql_select(public_rs, "SELECT * FROM tbl_attendance WHERE ID = '" & lbl_id2.Caption & "'")
          If public_rs.RecordCount = 0 Then
            attendanceform.lbl_id2.Caption = lbl_id2.Caption
            attendanceform.lbl_name2.Caption = lbl_name2.Caption
            Call load_form(attendanceform, True)
          Else
            Call load_grade
          End If
        
    Else
        Call load_grade
    End If

End Sub

Private Sub lbl_view_attendance_Click()
    attendanceform.lbl_id2.Caption = lbl_id2
    attendanceform.lbl_name2.Caption = lbl_name2
    
    Call mysql_select(public_rs, "SELECT * FROM tbl_attendance WHERE ID = '" & lbl_id2 & "' And SY ='" & mainteacherform.cmb_sy & "' ")
     If public_rs.RecordCount = 0 Then
        MsgBox "No attendance record for this student. Please complete first the final grade of character building."
        Exit Sub
    Else
     attendanceform.txt_school_days.Text = public_rs.Fields("no_school_days")
     attendanceform.txt_days_absent.Text = public_rs.Fields("no_days_absent")
     attendanceform.txt_days_tardy.Text = public_rs.Fields("no_days_tardiness")
     attendanceform.txt_days_present.Text = public_rs.Fields("no_days_present")
    Call load_form(attendanceform, True)
    End If
End Sub
