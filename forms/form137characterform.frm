VERSION 5.00
Begin VB.Form form137characterform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Student Form 137 Character Grades"
   ClientHeight    =   5760
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "form137characterform.frx":0000
   ScaleHeight     =   5760
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
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
      Height          =   2175
      Left            =   2160
      TabIndex        =   14
      Top             =   2640
      Width           =   4575
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
         Left            =   480
         TabIndex        =   18
         Top             =   480
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
         Left            =   480
         TabIndex        =   17
         Top             =   840
         Width           =   2655
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
         Left            =   480
         TabIndex        =   16
         Top             =   1200
         Width           =   1815
      End
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
         Left            =   480
         TabIndex        =   15
         Top             =   1560
         Width           =   3975
      End
   End
   Begin VB.CommandButton cmd_print 
      Height          =   615
      Left            =   3360
      Picture         =   "form137characterform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   5040
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   1695
      Left            =   120
      TabIndex        =   0
      Top             =   600
      Width           =   8775
      Begin VB.ComboBox cmb_period 
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
         ItemData        =   "form137characterform.frx":1C324
         Left            =   5880
         List            =   "form137characterform.frx":1C337
         TabIndex        =   2
         Top             =   1200
         Width           =   2535
      End
      Begin VB.ComboBox cmb_sy 
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
         ItemData        =   "form137characterform.frx":1C376
         Left            =   1680
         List            =   "form137characterform.frx":1C378
         TabIndex        =   1
         Top             =   1200
         Width           =   2535
      End
      Begin VB.Label Label6 
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
         Left            =   4440
         TabIndex        =   13
         Top             =   1320
         Width           =   1455
      End
      Begin VB.Label lbl_section 
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
         Left            =   5880
         TabIndex        =   12
         Top             =   960
         Width           =   2775
      End
      Begin VB.Label lbl_level 
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
         Left            =   1080
         TabIndex        =   11
         Top             =   960
         Width           =   3135
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
         Left            =   960
         TabIndex        =   10
         Top             =   600
         Width           =   7455
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
         Left            =   960
         TabIndex        =   9
         Top             =   240
         Width           =   5055
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
         Left            =   240
         TabIndex        =   8
         Top             =   240
         Width           =   615
      End
      Begin VB.Label Label2 
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
         Left            =   240
         TabIndex        =   7
         Top             =   600
         Width           =   1455
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Level:"
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
         TabIndex        =   6
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label4 
         BackStyle       =   0  'Transparent
         Caption         =   "Section:"
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
         Left            =   4440
         TabIndex        =   5
         Top             =   960
         Width           =   1455
      End
      Begin VB.Label Label5 
         BackStyle       =   0  'Transparent
         Caption         =   "School Year:"
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
         TabIndex        =   4
         Top             =   1320
         Width           =   1455
      End
   End
   Begin VB.Label lbl_export 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "Export Form-137"
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
      Left            =   7080
      TabIndex        =   19
      Top             =   5400
      Width           =   1815
   End
End
Attribute VB_Name = "form137characterform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_grade As New ADODB.Recordset
Public public_rs2 As New ADODB.Recordset
Public rs_grade1 As New ADODB.Recordset
Public rs_grade2 As New ADODB.Recordset
Public rs_grade3 As New ADODB.Recordset
Public rs_grade4 As New ADODB.Recordset
Public rs_grade5 As New ADODB.Recordset
Public rs_grade6 As New ADODB.Recordset
Public rs_other As New ADODB.Recordset
Public temp As String
Dim excelApp As Excel.Application
Dim ExcelWorkbook As Excel.Workbook
Dim ExcelSheet As Excel.Worksheet
Dim MyMonth As String
Dim MyYear As String
Dim Mydirectory As String
Dim MyFileName As String
Dim sql_string As String
Dim average As Double
Dim remark As String

Private Sub cmb_period_Click()
     If cmb_sy.Text = "" Then
        MsgBox "Please select a school year first."
        Exit Sub
    End If
End Sub

Private Sub cmb_period_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select a period from the list."
    cmb_period.Text = ""
End Sub

Private Sub cmb_sy_KeyUp(KeyCode As Integer, Shift As Integer)
     MsgBox "Please select school year from the list."
    cmb_sy.Text = ""
End Sub

Private Sub cmd_print_Click()
    If cmb_sy.Text = "" Then
        MsgBox "Please select a school year."
        Exit Sub
    End If
    If cmb_period.Text = "" Then
        MsgBox "Please select a period."
        Exit Sub
    End If
   
      dr_character.Sections(2).Controls("lbl_sy").Caption = cmb_sy.Text
      dr_character.Sections(2).Controls("lbl_date").Caption = Format(Now, "mmmm dd, yyyy") & "/ " & Time
        dr_character.Sections(2).Controls("lbl_level").Caption = lbl_level.Caption
        dr_character.Sections(2).Controls("lbl_section").Caption = lbl_section.Caption
        dr_character.Sections(2).Controls("lbl_id").Caption = lbl_id.Caption
        dr_character.Sections(2).Controls("lbl_name").Caption = lbl_name.Caption
        dr_character.Sections(2).Controls("lbl_period").Caption = cmb_period.Text
        
        Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE SY = '" & cmb_sy.Text & "' AND ID = '" & lbl_id.Caption & "' AND Period = '" & cmb_period.Text & "'")
            If public_rs.RecordCount = 0 Then
                MsgBox "No character grade."
            Else
            dr_character.Sections(2).Controls("lbl_1").Caption = public_rs.Fields("Honesty").value
            dr_character.Sections(2).Controls("lbl_2").Caption = public_rs.Fields("Courtesy").value
            dr_character.Sections(2).Controls("lbl_3").Caption = public_rs.Fields("Helpfulness_and_Cooperation").value
            dr_character.Sections(2).Controls("lbl_4").Caption = public_rs.Fields("Resourcefulness_and_Creativity").value
            dr_character.Sections(2).Controls("lbl_5").Caption = public_rs.Fields("Consideration_for_Others").value
            dr_character.Sections(2).Controls("lbl_6").Caption = public_rs.Fields("Sportsmanship").value
            dr_character.Sections(2).Controls("lbl_7").Caption = public_rs.Fields("Obedience").value
            dr_character.Sections(2).Controls("lbl_8").Caption = public_rs.Fields("Self_Reliance").value
            dr_character.Sections(2).Controls("lbl_9").Caption = public_rs.Fields("Industry").value
            dr_character.Sections(2).Controls("lbl_10").Caption = public_rs.Fields("Cleanliness_and_Orderliness").value
            dr_character.Sections(2).Controls("lbl_11").Caption = public_rs.Fields("Promptness_and_Punctuality").value
            dr_character.Sections(2).Controls("lbl_12").Caption = public_rs.Fields("Sense_of_Responsibility").value
            dr_character.Sections(2).Controls("lbl_13").Caption = public_rs.Fields("Love_of_God").value
            dr_character.Sections(2).Controls("lbl_14").Caption = public_rs.Fields("Patriotism_and_Love_of_Country").value
            Set dr_character.DataSource = public_rs
             
            dr_character.Show vbModal, Me
            End If
    
End Sub
Public Sub next_prod()
     Dim sy_4, sy_5, sy_6, sy_7 As String
     
      Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & temp & "' AND lvl_name = 'Grade 4'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(23, 5).value = ""
         ExcelSheet.Cells(23, 3).value = ""
        sy_4 = ""
    Else
        ExcelSheet.Cells(25, 5).value = public_rs.Fields("SY").value
         ExcelSheet.Cells(25, 3).value = "IV"
        sy_4 = public_rs.Fields("SY").value
    End If
   Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='1st Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 2).value = ""
        ExcelSheet.Cells(26, 2).value = ""
        ExcelSheet.Cells(27, 2).value = ""
        ExcelSheet.Cells(28, 2).value = ""
        ExcelSheet.Cells(29, 2).value = ""
        ExcelSheet.Cells(30, 2).value = ""
        ExcelSheet.Cells(31, 2).value = ""
        ExcelSheet.Cells(32, 2).value = ""
        ExcelSheet.Cells(33, 2).value = ""
        ExcelSheet.Cells(34, 2).value = ""
        ExcelSheet.Cells(35, 2).value = ""
        ExcelSheet.Cells(36, 2).value = ""
        ExcelSheet.Cells(37, 2).value = ""
        ExcelSheet.Cells(38, 2).value = ""
    Else
        ExcelSheet.Cells(25, 2).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 2).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 2).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 2).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 2).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 2).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 2).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 2).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 2).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 2).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 2).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 2).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 2).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 2).value = public_rs.Fields(18)
    End If
    
    Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='2nd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 3).value = ""
        ExcelSheet.Cells(26, 3).value = ""
        ExcelSheet.Cells(27, 3).value = ""
        ExcelSheet.Cells(28, 3).value = ""
        ExcelSheet.Cells(29, 3).value = ""
        ExcelSheet.Cells(30, 3).value = ""
        ExcelSheet.Cells(31, 3).value = ""
        ExcelSheet.Cells(32, 3).value = ""
        ExcelSheet.Cells(33, 3).value = ""
        ExcelSheet.Cells(34, 3).value = ""
        ExcelSheet.Cells(35, 3).value = ""
        ExcelSheet.Cells(36, 3).value = ""
        ExcelSheet.Cells(37, 3).value = ""
        ExcelSheet.Cells(38, 3).value = ""
    Else
        ExcelSheet.Cells(25, 3).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 3).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 3).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 3).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 3).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 3).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 3).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 3).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 3).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 3).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 3).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 3).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 3).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 3).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='3rd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 4).value = ""
        ExcelSheet.Cells(26, 4).value = ""
        ExcelSheet.Cells(27, 4).value = ""
        ExcelSheet.Cells(28, 4).value = ""
        ExcelSheet.Cells(29, 4).value = ""
        ExcelSheet.Cells(30, 4).value = ""
        ExcelSheet.Cells(31, 4).value = ""
        ExcelSheet.Cells(32, 4).value = ""
        ExcelSheet.Cells(33, 4).value = ""
        ExcelSheet.Cells(34, 4).value = ""
        ExcelSheet.Cells(35, 4).value = ""
        ExcelSheet.Cells(36, 4).value = ""
        ExcelSheet.Cells(37, 4).value = ""
        ExcelSheet.Cells(38, 4).value = ""
    Else
        ExcelSheet.Cells(25, 4).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 4).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 4).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 4).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 4).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 4).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 4).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 4).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 4).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 4).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 4).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 4).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 4).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 4).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='4th Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 5).value = ""
        ExcelSheet.Cells(26, 5).value = ""
        ExcelSheet.Cells(27, 5).value = ""
        ExcelSheet.Cells(28, 5).value = ""
        ExcelSheet.Cells(29, 5).value = ""
        ExcelSheet.Cells(30, 5).value = ""
        ExcelSheet.Cells(31, 5).value = ""
        ExcelSheet.Cells(32, 5).value = ""
        ExcelSheet.Cells(33, 5).value = ""
        ExcelSheet.Cells(34, 5).value = ""
        ExcelSheet.Cells(35, 5).value = ""
        ExcelSheet.Cells(36, 5).value = ""
        ExcelSheet.Cells(37, 5).value = ""
        ExcelSheet.Cells(38, 5).value = ""
    Else
        ExcelSheet.Cells(25, 5).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 5).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 5).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 5).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 5).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 5).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 5).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 5).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 5).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 5).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 5).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 5).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 5).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 5).value = public_rs.Fields(18)
    End If
      Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_4 & "' AND Period='Final'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 6).value = ""
        ExcelSheet.Cells(26, 6).value = ""
        ExcelSheet.Cells(27, 6).value = ""
        ExcelSheet.Cells(28, 6).value = ""
        ExcelSheet.Cells(29, 6).value = ""
        ExcelSheet.Cells(30, 6).value = ""
        ExcelSheet.Cells(31, 6).value = ""
        ExcelSheet.Cells(32, 6).value = ""
        ExcelSheet.Cells(33, 6).value = ""
        ExcelSheet.Cells(34, 6).value = ""
        ExcelSheet.Cells(35, 6).value = ""
        ExcelSheet.Cells(36, 6).value = ""
        ExcelSheet.Cells(37, 6).value = ""
        ExcelSheet.Cells(38, 6).value = ""
    Else
        ExcelSheet.Cells(25, 6).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 6).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 6).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 6).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 6).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 6).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 6).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 6).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 6).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 6).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 6).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 6).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 6).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 6).value = public_rs.Fields(18)
    End If
    
           Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 5'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(23, 10).value = ""
         ExcelSheet.Cells(23, 8).value = ""
        sy_5 = ""
    Else
        ExcelSheet.Cells(23, 10).value = public_rs.Fields("SY").value
         ExcelSheet.Cells(23, 8).value = "IV"
        sy_5 = public_rs.Fields("SY").value
    End If
    Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='1st Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 7).value = ""
        ExcelSheet.Cells(26, 7).value = ""
        ExcelSheet.Cells(27, 7).value = ""
        ExcelSheet.Cells(28, 7).value = ""
        ExcelSheet.Cells(29, 7).value = ""
        ExcelSheet.Cells(30, 7).value = ""
        ExcelSheet.Cells(31, 7).value = ""
        ExcelSheet.Cells(32, 7).value = ""
        ExcelSheet.Cells(33, 7).value = ""
        ExcelSheet.Cells(34, 7).value = ""
        ExcelSheet.Cells(35, 7).value = ""
        ExcelSheet.Cells(36, 7).value = ""
        ExcelSheet.Cells(37, 7).value = ""
        ExcelSheet.Cells(38, 7).value = ""
    Else
        ExcelSheet.Cells(25, 7).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 7).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 7).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 7).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 7).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 7).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 7).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 7).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 7).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 7).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 7).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 7).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 7).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 7).value = public_rs.Fields(18)
    End If
    
    Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='2nd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 8).value = ""
        ExcelSheet.Cells(26, 8).value = ""
        ExcelSheet.Cells(27, 8).value = ""
        ExcelSheet.Cells(28, 8).value = ""
        ExcelSheet.Cells(29, 8).value = ""
        ExcelSheet.Cells(30, 8).value = ""
        ExcelSheet.Cells(31, 8).value = ""
        ExcelSheet.Cells(32, 8).value = ""
        ExcelSheet.Cells(33, 8).value = ""
        ExcelSheet.Cells(34, 8).value = ""
        ExcelSheet.Cells(35, 8).value = ""
        ExcelSheet.Cells(36, 8).value = ""
        ExcelSheet.Cells(37, 8).value = ""
        ExcelSheet.Cells(38, 8).value = ""
    Else
        ExcelSheet.Cells(25, 8).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 8).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 8).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 8).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 8).value = public_rs.Fields(9)
        ExcelSheet.Cells(30).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 8).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 8).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 8).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 8).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 8).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 8).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 8).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 8).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='3rd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 9).value = ""
        ExcelSheet.Cells(26, 9).value = ""
        ExcelSheet.Cells(27, 9).value = ""
        ExcelSheet.Cells(28, 9).value = ""
        ExcelSheet.Cells(29, 9).value = ""
        ExcelSheet.Cells(30, 9).value = ""
        ExcelSheet.Cells(31, 9).value = ""
        ExcelSheet.Cells(32, 9).value = ""
        ExcelSheet.Cells(33, 9).value = ""
        ExcelSheet.Cells(34, 9).value = ""
        ExcelSheet.Cells(35, 9).value = ""
        ExcelSheet.Cells(36, 9).value = ""
        ExcelSheet.Cells(37, 9).value = ""
        ExcelSheet.Cells(38, 9).value = ""
    Else
        ExcelSheet.Cells(25, 9).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 9).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 9).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 9).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 9).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 9).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 9).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 9).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 9).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 9).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 9).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 9).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 9).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 9).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='4th Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 10).value = ""
        ExcelSheet.Cells(26, 10).value = ""
        ExcelSheet.Cells(27, 10).value = ""
        ExcelSheet.Cells(28, 10).value = ""
        ExcelSheet.Cells(29, 10).value = ""
        ExcelSheet.Cells(30, 10).value = ""
        ExcelSheet.Cells(31, 10).value = ""
        ExcelSheet.Cells(32, 10).value = ""
        ExcelSheet.Cells(33, 10).value = ""
        ExcelSheet.Cells(34, 10).value = ""
        ExcelSheet.Cells(35, 10).value = ""
        ExcelSheet.Cells(36, 10).value = ""
        ExcelSheet.Cells(37, 10).value = ""
        ExcelSheet.Cells(38, 10).value = ""
    Else
        ExcelSheet.Cells(25, 10).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 10).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 10).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 10).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 10).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 10).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 10).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 10).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 10).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 10).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 10).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 10).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 10).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 10).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_5 & "' AND Period='Final'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 11).value = ""
        ExcelSheet.Cells(26, 11).value = ""
        ExcelSheet.Cells(27, 11).value = ""
        ExcelSheet.Cells(28, 11).value = ""
        ExcelSheet.Cells(29, 11).value = ""
        ExcelSheet.Cells(30, 11).value = ""
        ExcelSheet.Cells(31, 11).value = ""
        ExcelSheet.Cells(32, 11).value = ""
        ExcelSheet.Cells(33, 11).value = ""
        ExcelSheet.Cells(34, 11).value = ""
        ExcelSheet.Cells(35, 11).value = ""
        ExcelSheet.Cells(36, 11).value = ""
        ExcelSheet.Cells(37, 11).value = ""
        ExcelSheet.Cells(38, 11).value = ""
    Else
        ExcelSheet.Cells(25, 11).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 11).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 11).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 11).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 11).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 11).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 11).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 11).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 11).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 11).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 11).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 11).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 11).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 11).value = public_rs.Fields(18)
    End If
    
           Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 6'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(23, 15).value = ""
        ExcelSheet.Cells(23, 13).value = ""
        sy_6 = ""
    Else
        ExcelSheet.Cells(23, 15).value = public_rs.Fields("SY").value
        ExcelSheet.Cells(23, 13).value = "VI"
        sy_6 = public_rs.Fields("SY").value
    End If
   Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='1st Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 12).value = ""
        ExcelSheet.Cells(26, 12).value = ""
        ExcelSheet.Cells(27, 12).value = ""
        ExcelSheet.Cells(28, 12).value = ""
        ExcelSheet.Cells(29, 12).value = ""
        ExcelSheet.Cells(30, 12).value = ""
        ExcelSheet.Cells(31, 12).value = ""
        ExcelSheet.Cells(32, 12).value = ""
        ExcelSheet.Cells(33, 12).value = ""
        ExcelSheet.Cells(34, 12).value = ""
        ExcelSheet.Cells(35, 12).value = ""
        ExcelSheet.Cells(36, 12).value = ""
        ExcelSheet.Cells(37, 12).value = ""
        ExcelSheet.Cells(38, 12).value = ""
    Else
        ExcelSheet.Cells(25, 12).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 12).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 12).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 12).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 12).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 12).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 12).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 12).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 12).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 12).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 12).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 12).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 12).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 12).value = public_rs.Fields(18)
    End If
    
    Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='2nd Grading'")
   
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 13).value = ""
        ExcelSheet.Cells(26, 13).value = ""
        ExcelSheet.Cells(27, 13).value = ""
        ExcelSheet.Cells(28, 13).value = ""
        ExcelSheet.Cells(29, 13).value = ""
        ExcelSheet.Cells(30, 13).value = ""
        ExcelSheet.Cells(31, 13).value = ""
        ExcelSheet.Cells(32, 13).value = ""
        ExcelSheet.Cells(33, 13).value = ""
        ExcelSheet.Cells(34, 13).value = ""
        ExcelSheet.Cells(35, 13).value = ""
        ExcelSheet.Cells(36, 13).value = ""
        ExcelSheet.Cells(37, 13).value = ""
        ExcelSheet.Cells(38, 13).value = ""
    Else
        ExcelSheet.Cells(25, 13).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 13).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 13).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 13).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 13).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 13).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 13).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 13).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 13).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 13).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 13).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 13).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 13).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 13).value = public_rs.Fields(18)
    End If
      Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='3rd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 14).value = ""
        ExcelSheet.Cells(26, 14).value = ""
        ExcelSheet.Cells(27, 14).value = ""
        ExcelSheet.Cells(28, 14).value = ""
        ExcelSheet.Cells(29, 14).value = ""
        ExcelSheet.Cells(30, 14).value = ""
        ExcelSheet.Cells(31, 14).value = ""
        ExcelSheet.Cells(32, 14).value = ""
        ExcelSheet.Cells(33, 14).value = ""
        ExcelSheet.Cells(34, 14).value = ""
        ExcelSheet.Cells(35, 14).value = ""
        ExcelSheet.Cells(36, 14).value = ""
        ExcelSheet.Cells(37, 14).value = ""
        ExcelSheet.Cells(38, 14).value = ""
    Else
        ExcelSheet.Cells(25, 14).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 14).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 14).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 14).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 14).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 14).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 14).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 14).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 14).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 14).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 14).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 14).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 14).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 14).value = public_rs.Fields(18)
    End If
      Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='4th Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 15).value = ""
        ExcelSheet.Cells(26, 15).value = ""
        ExcelSheet.Cells(27, 15).value = ""
        ExcelSheet.Cells(28, 15).value = ""
        ExcelSheet.Cells(29, 15).value = ""
        ExcelSheet.Cells(30, 15).value = ""
        ExcelSheet.Cells(31, 15).value = ""
        ExcelSheet.Cells(32, 15).value = ""
        ExcelSheet.Cells(33, 15).value = ""
        ExcelSheet.Cells(34, 15).value = ""
        ExcelSheet.Cells(35, 15).value = ""
        ExcelSheet.Cells(36, 15).value = ""
        ExcelSheet.Cells(37, 15).value = ""
        ExcelSheet.Cells(38, 15).value = ""
    Else
        ExcelSheet.Cells(25, 15).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 15).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 15).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 15).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 15).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 15).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 15).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 15).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 15).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 15).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 15).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 15).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 15).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 15).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_6 & "' AND Period='Final'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(25, 16).value = ""
        ExcelSheet.Cells(26, 16).value = ""
        ExcelSheet.Cells(27, 16).value = ""
        ExcelSheet.Cells(28, 16).value = ""
        ExcelSheet.Cells(29, 16).value = ""
        ExcelSheet.Cells(30, 16).value = ""
        ExcelSheet.Cells(31, 16).value = ""
        ExcelSheet.Cells(32, 16).value = ""
        ExcelSheet.Cells(33, 16).value = ""
        ExcelSheet.Cells(34, 16).value = ""
        ExcelSheet.Cells(35, 16).value = ""
        ExcelSheet.Cells(36, 16).value = ""
        ExcelSheet.Cells(37, 16).value = ""
        ExcelSheet.Cells(38, 16).value = ""
    Else
        ExcelSheet.Cells(25, 16).value = public_rs.Fields(5)
        ExcelSheet.Cells(26, 16).value = public_rs.Fields(6)
        ExcelSheet.Cells(27, 16).value = public_rs.Fields(7)
        ExcelSheet.Cells(28, 16).value = public_rs.Fields(8)
        ExcelSheet.Cells(29, 16).value = public_rs.Fields(9)
        ExcelSheet.Cells(30, 16).value = public_rs.Fields(10)
        ExcelSheet.Cells(31, 16).value = public_rs.Fields(11)
        ExcelSheet.Cells(32, 16).value = public_rs.Fields(12)
        ExcelSheet.Cells(33, 16).value = public_rs.Fields(13)
        ExcelSheet.Cells(34, 16).value = public_rs.Fields(14)
        ExcelSheet.Cells(35, 16).value = public_rs.Fields(15)
        ExcelSheet.Cells(36, 16).value = public_rs.Fields(16)
        ExcelSheet.Cells(37, 16).value = public_rs.Fields(17)
        ExcelSheet.Cells(38, 16).value = public_rs.Fields(18)
    End If
   
    
         
   
End Sub

Private Sub lbl_export_Click()
      If lbl_id.Caption = "" Then
        MsgBox "No record selected."
        Exit Sub
    End If
    MyFileName = App.Path & "\Form-137\" & lbl_id.Caption & "-" & lbl_name.Caption & "-Character Building.xls"
    On Error Resume Next
    Set excelApp = CreateObject("Excel.Application")
'if file exists, place file name in FileCheck
FileCheck = Dir$(MyFileName)
  If FileCheck = MyMonth + "_" + MyYear + MyExtension Then
    'Workbook exists, open it
    Set ExcelWorkbook = excelApp.Workbooks.Open(MyFileName)
    Set ExcelSheet = ExcelWorkbook.Worksheets(1)
  Else
'create Excel object
Set excelApp = CreateObject("Excel.Application")
    Set ExcelWorkbook = excelApp.Workbooks.Add
    Set ExcelSheet = ExcelWorkbook.Worksheets(1)
    ExcelSheet.Name = "Character Building Grades"
        
   excelApp.Range("A1:P1").Merge
   excelApp.Range("A2:P2").Merge
   excelApp.Range("A3:A4").Merge
   excelApp.Range("A23:A24").Merge
   excelApp.Range("A2:P2").Font.Bold = True
   excelApp.Range("A2:P2").Font.Size = 16
   excelApp.Range("A3:A3").Font.Bold = True
    excelApp.Range("A3:A3").Font.Size = 16
    excelApp.Range("C3:D3").Merge
    excelApp.Range("C23:D23").Merge
    excelApp.Range("G3:H3").Merge
    excelApp.Range("L3:M3").Merge
    excelApp.Range("G23:H23").Merge
    excelApp.Range("L23:M23").Merge
    excelApp.Range("A23:A23").Font.Size = 16
   
    ExcelSheet.Cells(2, 1).value = "CHARACTER BUILDING"
    ExcelSheet.Cells(3, 1).value = "TRAITS"
    ExcelSheet.Cells(5, 1).value = "1. Honesty"
    ExcelSheet.Cells(6, 1).value = "2. Courtesy"
    ExcelSheet.Cells(7, 1).value = "3. Helpfulness & Cooperation"
    ExcelSheet.Cells(8, 1).value = "4. Resourcefulness and Creativity"
    ExcelSheet.Cells(9, 1).value = "5. Consideration for Others"
    ExcelSheet.Cells(10, 1).value = "6. Sportsmanship"
    ExcelSheet.Cells(11, 1).value = "7. Obedience"
    ExcelSheet.Cells(12, 1).value = "8. Self-Reliance"
    ExcelSheet.Cells(13, 1).value = "9. Industry"
    ExcelSheet.Cells(14, 1).value = "10. Cleanliness & Orderliness"
    ExcelSheet.Cells(15, 1).value = "11. Promptness and Punctuality"
    ExcelSheet.Cells(16, 1).value = "12. Sense of Responisibility"
    ExcelSheet.Cells(17, 1).value = "13. Love of God"
    ExcelSheet.Cells(18, 1).value = "14. Patriotism and Love of Country"
    
    ExcelSheet.Cells(23, 1).value = "TRAITS"
    ExcelSheet.Cells(25, 1).value = "1. Honesty"
    ExcelSheet.Cells(26, 1).value = "2. Courtesy"
    ExcelSheet.Cells(27, 1).value = "3. Helpfulness & Cooperation"
    ExcelSheet.Cells(28, 1).value = "4. Resourcefulness and Creativity"
    ExcelSheet.Cells(29, 1).value = "5. Consideration for Others"
    ExcelSheet.Cells(30, 1).value = "6. Sportsmanship"
    ExcelSheet.Cells(31, 1).value = "7. Obedience"
    ExcelSheet.Cells(32, 1).value = "8. Self-Reliance"
    ExcelSheet.Cells(33, 1).value = "9. Industry"
    ExcelSheet.Cells(34, 1).value = "10. Cleanliness & Orderliness"
    ExcelSheet.Cells(35, 1).value = "11. Promptness and Punctuality"
    ExcelSheet.Cells(36, 1).value = "12. Sense of Responisibility"
    ExcelSheet.Cells(37, 1).value = "13. Love of God"
    ExcelSheet.Cells(38, 1).value = "14. Patriotism and Love of Country"
    ExcelSheet.Cells(3, 2).value = "Gr"
    ExcelSheet.Cells(3, 7).value = "Gr"
    ExcelSheet.Cells(3, 12).value = "Gr"
    ExcelSheet.Cells(23, 2).value = "Gr"
    ExcelSheet.Cells(23, 7).value = "Gr"
    ExcelSheet.Cells(23, 12).value = "Gr"
    ExcelSheet.Cells(4, 2).value = "1"
    ExcelSheet.Cells(4, 3).value = "2"
    ExcelSheet.Cells(4, 4).value = "3"
     ExcelSheet.Cells(4, 5).value = "4"
      ExcelSheet.Cells(4, 6).value = "F.R"
     ExcelSheet.Cells(4, 7).value = "1"
    ExcelSheet.Cells(4, 8).value = "2"
    ExcelSheet.Cells(4, 9).value = "3"
     ExcelSheet.Cells(4, 10).value = "4"
      ExcelSheet.Cells(4, 11).value = "F.R"
       ExcelSheet.Cells(4, 12).value = "1"
    ExcelSheet.Cells(4, 13).value = "2"
    ExcelSheet.Cells(4, 14).value = "3"
     ExcelSheet.Cells(4, 15).value = "4"
      ExcelSheet.Cells(4, 16).value = "F.R"
       ExcelSheet.Cells(24, 2).value = "1"
    ExcelSheet.Cells(24, 3).value = "2"
    ExcelSheet.Cells(24, 4).value = "3"
     ExcelSheet.Cells(24, 5).value = "4"
      ExcelSheet.Cells(24, 6).value = "F.R"
     ExcelSheet.Cells(24, 7).value = "1"
    ExcelSheet.Cells(24, 8).value = "2"
    ExcelSheet.Cells(24, 9).value = "3"
     ExcelSheet.Cells(24, 10).value = "4"
      ExcelSheet.Cells(24, 11).value = "F.R"
       ExcelSheet.Cells(24, 12).value = "1"
    ExcelSheet.Cells(24, 13).value = "2"
    ExcelSheet.Cells(24, 14).value = "3"
     ExcelSheet.Cells(24, 15).value = "4"
      ExcelSheet.Cells(24, 16).value = "F.R"
      excelApp.Range("B42:F42").Merge
      excelApp.Range("B43:F43").Merge
      excelApp.Range("G42:K42").Merge
      excelApp.Range("G43:K43").Merge
      ExcelSheet.Cells(42, 2).value = "A - Outstanding"
      ExcelSheet.Cells(43, 2).value = "B - Very Satisfactory"
      ExcelSheet.Cells(42, 7).value = "C - Satisfactory"
      ExcelSheet.Cells(43, 7).value = "D - Needs Improvement"
        excelApp.Range("B42:P42").Font.Bold = True
        excelApp.Range("B43:P43").Font.Bold = True
        excelApp.Range("B:P").ColumnWidth = 9
        excelApp.Range("A:A").ColumnWidth = 35
        excelApp.Range("B:P").HorizontalAlignment = xlCenter
         excelApp.Range("A23:A23").Font.Bold = True
    
    Dim sy_1, sy_2, sy_3 As String
    Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 1'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(3, 5).value = ""
        ExcelSheet.Cells(3, 3).value = ""
        sy_1 = ""
    Else
        ExcelSheet.Cells(3, 5).value = public_rs.Fields("SY").value
        ExcelSheet.Cells(3, 3).value = "I"
        sy_1 = public_rs.Fields("SY").value
    End If
    
    Call mysql_select(public_rs2, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='1st Grading'")
    If public_rs2.RecordCount = 0 Then
        ExcelSheet.Cells(5, 2).value = ""
        ExcelSheet.Cells(6, 2).value = ""
        ExcelSheet.Cells(7, 2).value = ""
        ExcelSheet.Cells(8, 2).value = ""
        ExcelSheet.Cells(9, 2).value = ""
        ExcelSheet.Cells(10, 2).value = ""
        ExcelSheet.Cells(11, 2).value = ""
        ExcelSheet.Cells(12, 2).value = ""
        ExcelSheet.Cells(13, 2).value = ""
        ExcelSheet.Cells(14, 2).value = ""
        ExcelSheet.Cells(15, 2).value = ""
        ExcelSheet.Cells(16, 2).value = ""
        ExcelSheet.Cells(17, 2).value = ""
        ExcelSheet.Cells(18, 2).value = ""
    Else
        ExcelSheet.Cells(5, 2).value = public_rs2.Fields(5)
        ExcelSheet.Cells(6, 2).value = public_rs2.Fields(6)
        ExcelSheet.Cells(7, 2).value = public_rs2.Fields(7)
        ExcelSheet.Cells(8, 2).value = public_rs2.Fields(8)
        ExcelSheet.Cells(9, 2).value = public_rs2.Fields(9)
        ExcelSheet.Cells(10, 2).value = public_rs2.Fields(10)
        ExcelSheet.Cells(11, 2).value = public_rs2.Fields(11)
        ExcelSheet.Cells(12, 2).value = public_rs2.Fields(12)
        ExcelSheet.Cells(13, 2).value = public_rs2.Fields(13)
        ExcelSheet.Cells(14, 2).value = public_rs2.Fields(14)
        ExcelSheet.Cells(15, 2).value = public_rs2.Fields(15)
        ExcelSheet.Cells(16, 2).value = public_rs2.Fields(16)
        ExcelSheet.Cells(17, 2).value = public_rs2.Fields(17)
        ExcelSheet.Cells(18, 2).value = public_rs2.Fields(18)
    End If
    
   Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='2nd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 3).value = ""
        ExcelSheet.Cells(6, 3).value = ""
        ExcelSheet.Cells(7, 3).value = ""
        ExcelSheet.Cells(8, 3).value = ""
        ExcelSheet.Cells(9, 3).value = ""
        ExcelSheet.Cells(10, 3).value = ""
        ExcelSheet.Cells(11, 3).value = ""
        ExcelSheet.Cells(12, 3).value = ""
        ExcelSheet.Cells(13, 3).value = ""
        ExcelSheet.Cells(14, 3).value = ""
        ExcelSheet.Cells(15, 3).value = ""
        ExcelSheet.Cells(16, 3).value = ""
        ExcelSheet.Cells(17, 3).value = ""
        ExcelSheet.Cells(18, 3).value = ""
    Else
        ExcelSheet.Cells(5, 3).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 3).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 3).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 3).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 3).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 3).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 3).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 3).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 3).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 3).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 3).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 3).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 3).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 3).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='3rd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 4).value = ""
        ExcelSheet.Cells(6, 4).value = ""
        ExcelSheet.Cells(7, 4).value = ""
        ExcelSheet.Cells(8, 4).value = ""
        ExcelSheet.Cells(9, 4).value = ""
        ExcelSheet.Cells(10, 4).value = ""
        ExcelSheet.Cells(11, 4).value = ""
        ExcelSheet.Cells(12, 4).value = ""
        ExcelSheet.Cells(13, 4).value = ""
        ExcelSheet.Cells(14, 4).value = ""
        ExcelSheet.Cells(15, 4).value = ""
        ExcelSheet.Cells(16, 4).value = ""
        ExcelSheet.Cells(17, 4).value = ""
        ExcelSheet.Cells(18, 4).value = ""
    Else
        ExcelSheet.Cells(5, 4).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 4).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 4).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 4).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 4).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 4).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 4).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 4).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 4).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 4).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 4).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 4).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 4).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 4).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='4th Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 5).value = ""
        ExcelSheet.Cells(6, 5).value = ""
        ExcelSheet.Cells(7, 5).value = ""
        ExcelSheet.Cells(8, 5).value = ""
        ExcelSheet.Cells(9, 5).value = ""
        ExcelSheet.Cells(10, 5).value = ""
        ExcelSheet.Cells(11, 5).value = ""
        ExcelSheet.Cells(12, 5).value = ""
        ExcelSheet.Cells(13, 5).value = ""
        ExcelSheet.Cells(14, 5).value = ""
        ExcelSheet.Cells(15, 5).value = ""
        ExcelSheet.Cells(16, 5).value = ""
        ExcelSheet.Cells(17, 5).value = ""
        ExcelSheet.Cells(18, 5).value = ""
    Else
        ExcelSheet.Cells(5, 5).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 5).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 5).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 5).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 5).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 5).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 5).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 5).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 5).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 5).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 5).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 5).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 5).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 5).value = public_rs.Fields(18)
    End If
      Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='Final'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 6).value = ""
        ExcelSheet.Cells(6, 6).value = ""
        ExcelSheet.Cells(7, 6).value = ""
        ExcelSheet.Cells(8, 6).value = ""
        ExcelSheet.Cells(9, 6).value = ""
        ExcelSheet.Cells(10, 6).value = ""
        ExcelSheet.Cells(11, 6).value = ""
        ExcelSheet.Cells(12, 6).value = ""
        ExcelSheet.Cells(13, 6).value = ""
        ExcelSheet.Cells(14, 6).value = ""
        ExcelSheet.Cells(15, 6).value = ""
        ExcelSheet.Cells(16, 6).value = ""
        ExcelSheet.Cells(17, 6).value = ""
        ExcelSheet.Cells(18, 6).value = ""
    Else
        ExcelSheet.Cells(5, 6).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 6).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 6).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 6).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 6).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 6).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 6).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 6).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 6).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 6).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 6).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 6).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 6).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 6).value = public_rs.Fields(18)
    End If
    
     Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 2'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(3, 10).value = ""
        ExcelSheet.Cells(3, 8).value = ""
        sy_2 = ""
    Else
        ExcelSheet.Cells(3, 10).value = public_rs.Fields("SY").value
        ExcelSheet.Cells(3, 8).value = "II"
        sy_2 = public_rs.Fields("SY").value
    End If
    Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='1st Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 7).value = ""
        ExcelSheet.Cells(6, 7).value = ""
        ExcelSheet.Cells(7, 7).value = ""
        ExcelSheet.Cells(8, 7).value = ""
        ExcelSheet.Cells(9, 7).value = ""
        ExcelSheet.Cells(10, 7).value = ""
        ExcelSheet.Cells(11, 7).value = ""
        ExcelSheet.Cells(12, 7).value = ""
        ExcelSheet.Cells(13, 7).value = ""
        ExcelSheet.Cells(14, 7).value = ""
        ExcelSheet.Cells(15, 7).value = ""
        ExcelSheet.Cells(16, 7).value = ""
        ExcelSheet.Cells(17, 7).value = ""
        ExcelSheet.Cells(18, 7).value = ""
    Else
        ExcelSheet.Cells(5, 7).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 7).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 7).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 7).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 7).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 7).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 7).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 7).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 7).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 7).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 7).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 7).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 7).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 7).value = public_rs.Fields(18)
    End If
    
    Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='2nd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 8).value = ""
        ExcelSheet.Cells(6, 8).value = ""
        ExcelSheet.Cells(7, 8).value = ""
        ExcelSheet.Cells(8, 8).value = ""
        ExcelSheet.Cells(9, 8).value = ""
        ExcelSheet.Cells(10, 8).value = ""
        ExcelSheet.Cells(11, 8).value = ""
        ExcelSheet.Cells(12, 8).value = ""
        ExcelSheet.Cells(13, 8).value = ""
        ExcelSheet.Cells(14, 8).value = ""
        ExcelSheet.Cells(15, 8).value = ""
        ExcelSheet.Cells(16, 8).value = ""
        ExcelSheet.Cells(17, 8).value = ""
        ExcelSheet.Cells(18, 8).value = ""
    Else
        ExcelSheet.Cells(5, 8).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 8).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 8).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 8).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 8).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 8).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 8).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 8).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 8).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 8).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 8).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 8).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 8).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 8).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='3rd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 9).value = ""
        ExcelSheet.Cells(6, 9).value = ""
        ExcelSheet.Cells(7, 9).value = ""
        ExcelSheet.Cells(8, 9).value = ""
        ExcelSheet.Cells(9, 9).value = ""
        ExcelSheet.Cells(10, 9).value = ""
        ExcelSheet.Cells(11, 9).value = ""
        ExcelSheet.Cells(12, 9).value = ""
        ExcelSheet.Cells(13, 9).value = ""
        ExcelSheet.Cells(14, 9).value = ""
        ExcelSheet.Cells(15, 9).value = ""
        ExcelSheet.Cells(16, 9).value = ""
        ExcelSheet.Cells(17, 9).value = ""
        ExcelSheet.Cells(18, 9).value = ""
    Else
        ExcelSheet.Cells(5, 9).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 9).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 9).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 9).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 9).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 9).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 9).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 9).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 9).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 9).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 9).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 9).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 9).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 9).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='4th Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 10).value = ""
        ExcelSheet.Cells(6, 10).value = ""
        ExcelSheet.Cells(7, 10).value = ""
        ExcelSheet.Cells(8, 10).value = ""
        ExcelSheet.Cells(9, 10).value = ""
        ExcelSheet.Cells(10, 10).value = ""
        ExcelSheet.Cells(11, 10).value = ""
        ExcelSheet.Cells(12, 10).value = ""
        ExcelSheet.Cells(13, 10).value = ""
        ExcelSheet.Cells(14, 10).value = ""
        ExcelSheet.Cells(15, 10).value = ""
        ExcelSheet.Cells(16, 10).value = ""
        ExcelSheet.Cells(17, 10).value = ""
        ExcelSheet.Cells(18, 10).value = ""
    Else
        ExcelSheet.Cells(5, 10).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 10).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 10).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 10).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 10).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 10).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 10).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 10).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 10).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 10).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 10).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 10).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 10).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 10).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_2 & "' AND Period='Final'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 11).value = ""
        ExcelSheet.Cells(6, 11).value = ""
        ExcelSheet.Cells(7, 11).value = ""
        ExcelSheet.Cells(8, 11).value = ""
        ExcelSheet.Cells(9, 11).value = ""
        ExcelSheet.Cells(10, 11).value = ""
        ExcelSheet.Cells(11, 11).value = ""
        ExcelSheet.Cells(12, 11).value = ""
        ExcelSheet.Cells(13, 11).value = ""
        ExcelSheet.Cells(14, 11).value = ""
        ExcelSheet.Cells(15, 11).value = ""
        ExcelSheet.Cells(16, 11).value = ""
        ExcelSheet.Cells(17, 11).value = ""
        ExcelSheet.Cells(18, 11).value = ""
    Else
        ExcelSheet.Cells(5, 11).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 11).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 11).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 11).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 11).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 11).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 11).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 11).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 11).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 11).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 11).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 11).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 11).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 11).value = public_rs.Fields(18)
    End If
    
      Call mysql_select(public_rs, "SELECT * FROM tbl_student_level WHERE ID = '" & lbl_id.Caption & "' AND lvl_name = 'Grade 3'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(3, 15).value = ""
        ExcelSheet.Cells(3, 13).value = ""
        sy_3 = ""
    Else
        ExcelSheet.Cells(3, 15).value = public_rs.Fields("SY").value
        ExcelSheet.Cells(3, 13).value = "III"
        sy_3 = public_rs.Fields("SY").value
    End If
    Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='1st Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 12).value = ""
        ExcelSheet.Cells(6, 12).value = ""
        ExcelSheet.Cells(7, 12).value = ""
        ExcelSheet.Cells(8, 12).value = ""
        ExcelSheet.Cells(9, 12).value = ""
        ExcelSheet.Cells(10, 12).value = ""
        ExcelSheet.Cells(11, 12).value = ""
        ExcelSheet.Cells(12, 12).value = ""
        ExcelSheet.Cells(13, 12).value = ""
        ExcelSheet.Cells(14, 12).value = ""
        ExcelSheet.Cells(15, 12).value = ""
        ExcelSheet.Cells(16, 12).value = ""
        ExcelSheet.Cells(17, 12).value = ""
        ExcelSheet.Cells(18, 12).value = ""
    Else
        ExcelSheet.Cells(5, 12).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 12).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 12).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 12).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 12).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 12).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 12).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 12).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 12).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 12).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 12).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 12).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 12).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 12).value = public_rs.Fields(18)
    End If
    
    Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='2nd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 13).value = ""
        ExcelSheet.Cells(6, 13).value = ""
        ExcelSheet.Cells(7, 13).value = ""
        ExcelSheet.Cells(8, 13).value = ""
        ExcelSheet.Cells(9, 13).value = ""
        ExcelSheet.Cells(10, 13).value = ""
        ExcelSheet.Cells(11, 13).value = ""
        ExcelSheet.Cells(12, 13).value = ""
        ExcelSheet.Cells(13, 13).value = ""
        ExcelSheet.Cells(14, 13).value = ""
        ExcelSheet.Cells(15, 13).value = ""
        ExcelSheet.Cells(16, 13).value = ""
        ExcelSheet.Cells(17, 13).value = ""
        ExcelSheet.Cells(18, 13).value = ""
    Else
        ExcelSheet.Cells(5, 13).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 13).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 13).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 13).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 13).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 13).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 13).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 13).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 13).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 13).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 13).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 13).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 13).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 13).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='3rd Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 14).value = ""
        ExcelSheet.Cells(6, 14).value = ""
        ExcelSheet.Cells(7, 14).value = ""
        ExcelSheet.Cells(8, 14).value = ""
        ExcelSheet.Cells(9, 14).value = ""
        ExcelSheet.Cells(10, 14).value = ""
        ExcelSheet.Cells(11, 14).value = ""
        ExcelSheet.Cells(12, 14).value = ""
        ExcelSheet.Cells(13, 14).value = ""
        ExcelSheet.Cells(14, 14).value = ""
        ExcelSheet.Cells(15, 14).value = ""
        ExcelSheet.Cells(16, 14).value = ""
        ExcelSheet.Cells(17, 14).value = ""
        ExcelSheet.Cells(18, 14).value = ""
    Else
        ExcelSheet.Cells(5, 14).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 14).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 14).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 14).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 14).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 14).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 14).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 14).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 14).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 14).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 14).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 14).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 14).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 14).value = public_rs.Fields(18)
    End If
     Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='4th Grading'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 15).value = ""
        ExcelSheet.Cells(6, 15).value = ""
        ExcelSheet.Cells(7, 15).value = ""
        ExcelSheet.Cells(8, 15).value = ""
        ExcelSheet.Cells(9, 15).value = ""
        ExcelSheet.Cells(10, 15).value = ""
        ExcelSheet.Cells(11, 15).value = ""
        ExcelSheet.Cells(12, 15).value = ""
        ExcelSheet.Cells(13, 15).value = ""
        ExcelSheet.Cells(14, 15).value = ""
        ExcelSheet.Cells(15, 15).value = ""
        ExcelSheet.Cells(16, 15).value = ""
        ExcelSheet.Cells(17, 15).value = ""
        ExcelSheet.Cells(18, 15).value = ""
    Else
        ExcelSheet.Cells(5, 15).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 15).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 15).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 15).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 15).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 15).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 15).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 15).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 15).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 15).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 15).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 15).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 15).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 15).value = public_rs.Fields(18)
    End If
      Call mysql_select(public_rs, "SELECT * FROM tbl_character_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_3 & "' AND Period='Final'")
    If public_rs.RecordCount = 0 Then
        ExcelSheet.Cells(5, 16).value = ""
        ExcelSheet.Cells(6, 16).value = ""
        ExcelSheet.Cells(7, 16).value = ""
        ExcelSheet.Cells(8, 16).value = ""
        ExcelSheet.Cells(9, 16).value = ""
        ExcelSheet.Cells(10, 16).value = ""
        ExcelSheet.Cells(11, 16).value = ""
        ExcelSheet.Cells(12, 16).value = ""
        ExcelSheet.Cells(13, 16).value = ""
        ExcelSheet.Cells(14, 16).value = ""
        ExcelSheet.Cells(15, 16).value = ""
        ExcelSheet.Cells(16, 16).value = ""
        ExcelSheet.Cells(17, 16).value = ""
        ExcelSheet.Cells(18, 16).value = ""
    Else
        ExcelSheet.Cells(5, 16).value = public_rs.Fields(5)
        ExcelSheet.Cells(6, 16).value = public_rs.Fields(6)
        ExcelSheet.Cells(7, 16).value = public_rs.Fields(7)
        ExcelSheet.Cells(8, 16).value = public_rs.Fields(8)
        ExcelSheet.Cells(9, 16).value = public_rs.Fields(9)
        ExcelSheet.Cells(10, 16).value = public_rs.Fields(10)
        ExcelSheet.Cells(11, 16).value = public_rs.Fields(11)
        ExcelSheet.Cells(12, 16).value = public_rs.Fields(12)
        ExcelSheet.Cells(13, 16).value = public_rs.Fields(13)
        ExcelSheet.Cells(14, 16).value = public_rs.Fields(14)
        ExcelSheet.Cells(15, 16).value = public_rs.Fields(15)
        ExcelSheet.Cells(16, 16).value = public_rs.Fields(16)
        ExcelSheet.Cells(17, 16).value = public_rs.Fields(17)
        ExcelSheet.Cells(18, 16).value = public_rs.Fields(18)
    End If
    
   With ExcelSheet.PageSetup
     .PaperSize = xlPaperLegal
     .FitToPagesWide = 1
   End With
    
   Call next_prod
    If FileCheck = MyMonth + "_" + MyYear + MyExtension Then
        'Save existing workbook
        ExcelWorkbook.Save
    Else
        'Save new workbook
        ExcelWorkbook.SaveAs MyFileName
    End If
    
        'Close Excel
     excelApp.Visible = True
        'ExcelWorkbook.Close savechanges:=False
        'excelApp.Quit
        'Set excelApp = Nothing
        'Set ExcelWorkbook = Nothing
        'Set ExcelSheet = Nothing
    'MsgBox "Form 137 for Character Building Grade has been exported to an excel file."
    End If
    Call attendance
End Sub
Public Sub attendance()
    MyFileName = App.Path & "\Form-137\" & lbl_id.Caption & "-" & lbl_name.Caption & "-Attendance.xls"
    On Error Resume Next
    Set excelApp = CreateObject("Excel.Application")
'if file exists, place file name in FileCheck
FileCheck = Dir$(MyFileName)
  If FileCheck = MyMonth + "_" + MyYear + MyExtension Then
    'Workbook exists, open it
    Set ExcelWorkbook = excelApp.Workbooks.Open(MyFileName)
    Set ExcelSheet = ExcelWorkbook.Worksheets(1)
  Else
'create Excel object
Set excelApp = CreateObject("Excel.Application")
    Set ExcelWorkbook = excelApp.Workbooks.Add
    Set ExcelSheet = ExcelWorkbook.Worksheets(1)
    ExcelSheet.Name = "Attendance"
    
    excelApp.Range("A1:G1").Merge
    excelApp.Range("A1:G1").Font.Size = 16
    excelApp.Range("A1:G1").Font.Bold = True
    excelApp.Range("A2:A2").ColumnWidth = 12
    excelApp.Range("B2:B2").ColumnWidth = 12
    excelApp.Range("C2:C2").ColumnWidth = 15
    excelApp.Range("D2:D2").ColumnWidth = 20
    excelApp.Range("E2:E2").ColumnWidth = 12
    excelApp.Range("F2:F2").ColumnWidth = 20
    excelApp.Range("G2:G2").ColumnWidth = 15
    excelApp.Range("A2:A2").RowHeight = 30
    excelApp.Range("B2:B2").RowHeight = 30
    excelApp.Range("C2:C2").RowHeight = 30
    excelApp.Range("D2:D2").RowHeight = 30
    excelApp.Range("E2:E2").RowHeight = 30
    excelApp.Range("F2:F2").RowHeight = 30
    excelApp.Range("G2:G2").RowHeight = 30
   
    ExcelSheet.Cells(1, 1).value = "ATTENDANCE RECORD"
    ExcelSheet.Cells(2, 1).value = "Grade"
    ExcelSheet.Cells(2, 2).value = "No. of School Days"
    ExcelSheet.Cells(2, 3).value = "No. of School Days Absent"
    ExcelSheet.Cells(2, 4).value = "Cause"
    ExcelSheet.Cells(2, 5).value = "No. of Times Tardy"
    ExcelSheet.Cells(2, 6).value = "Cause"
    ExcelSheet.Cells(2, 7).value = "No. of School Days Present"
    excelApp.Range("A2:G2").HorizontalAlignment = xlCenter
    excelApp.Range("A2:G2").Font.Bold = True
    
     Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND SY = '" & sy_1 & "' AND Period='1st Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            
        End If
    
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND Period='2nd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND  Period='3rd Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND Period='4th Grading'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            
        End If
        Call mysql_select(public_rs, "SELECT subject_code as Subject, period, grade, remark FROM tbl_grade WHERE ID = '" & lbl_id.Caption & "' AND  Period='Final'")
        no_subj = public_rs.RecordCount
        If no_subj = 0 Then
            
            average = 0
            remark = "No grades"
        Else
            average = 0
            While Not public_rs.EOF
                average = val(public_rs.Fields("grade")) + average
                public_rs.MoveNext
            Wend
            average = average / no_subj
            average = Round(average, 2)
            If average >= 90 Then
              remark = "A"
              ElseIf average >= 85 Then
              remark = "P"
              ElseIf average >= 80 Then
                  remark = "AP"
              ElseIf average >= 74 Then
                  remark = "D"
              Else
                  remark = "B"
            End If
            
          
            
            
        End If
        
    ExcelSheet.Cells(3, 1).value = remark
    
    
    
    
    Call mysql_select(public_rs, "SELECT * FROM tbl_attendance WHERE ID = '" & lbl_id.Caption & "'")
    ExcelSheet.Cells(3, 2).value = public_rs.Fields("no_school_days").value
    ExcelSheet.Cells(3, 3).value = public_rs.Fields("no_days_absent").value
    ExcelSheet.Cells(3, 3).value = public_rs.Fields("causes_of_absences").value
    ExcelSheet.Cells(3, 4).value = public_rs.Fields("no_days_tardiness").value
    ExcelSheet.Cells(3, 5).value = public_rs.Fields("causes_of_tardiness").value
    ExcelSheet.Cells(3, 6).value = public_rs.Fields("no_days_present").value
    
    
     ExcelSheet.Cells(15, 1).value = "CERTIFICATE OF TRANSFER"
    excelApp.Range("A15:G15").Merge
    excelApp.Range("A15:G15").Font.Size = 16
    excelApp.Range("A15:G15").Font.Bold = True
      excelApp.Range("A15:G15").HorizontalAlignment = xlCenter
      excelApp.Range("A17:G17").Merge
    ExcelSheet.Cells(17, 1).value = "TO WHOM IT MAY CONCERN:"
    excelApp.Range("A17:G17").Font.Size = 16
    excelApp.Range("A17:G17").Font.Bold = True
    excelApp.Range("A18:G18").Merge
    excelApp.Range("A19:G19").Merge
    excelApp.Range("F25:G25").Merge
    excelApp.Range("F29:G29").Merge
    excelApp.Range("A32:B32").Merge
    Dim sy_new As String
    sy_new = Format(Date, "yyyy") & "-" & Left(Format(Date, "yyyy"), 3) & Trim(Str(val(Right(Format(Date, "yyyy"), 1) + 1)))
    ExcelSheet.Cells(18, 1).value = "         This is to certify that this is a true record of the Elementary School Permanent Record of "
    ExcelSheet.Cells(19, 1).value = lbl_name.Caption & ". He/She is eligible for transfer and admission to Grade/Year II."
    excelApp.Range("A18:G18").Font.Size = 12
    excelApp.Range("A19:G19").Font.Size = 12
    ExcelSheet.Cells(25, 6).value = "Signature"
    ExcelSheet.Cells(29, 6).value = "Designation"
    ExcelSheet.Cells(32, 1).value = "Date"
    excelApp.Range("F25:G25").HorizontalAlignment = xlCenter
    excelApp.Range("F29:G29").HorizontalAlignment = xlCenter
    excelApp.Range("A32:B32").HorizontalAlignment = xlCenter
    
    If FileCheck = MyMonth + "_" + MyYear + MyExtension Then
        'Save existing workbook
        ExcelWorkbook.Save
    Else
        'Save new workbook
        ExcelWorkbook.SaveAs MyFileName
    End If
    
        'Close Excel
        excelApp.Visible = True
        'ExcelWorkbook.Close savechanges:=False
        'excelApp.Quit
        'Set excelApp = Nothing
        'Set ExcelWorkbook = Nothing
        'Set ExcelSheet = Nothing
   
    End If
    
End Sub
