VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form masterlistform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Masterlist"
   ClientHeight    =   6045
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   9000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "masterlistform.frx":0000
   ScaleHeight     =   6045
   ScaleWidth      =   9000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.ComboBox cmb_category 
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
      ItemData        =   "masterlistform.frx":1B3CE
      Left            =   3000
      List            =   "masterlistform.frx":1B3DB
      TabIndex        =   12
      Top             =   960
      Width           =   3135
   End
   Begin VB.CommandButton cmd_print 
      Height          =   615
      Left            =   3840
      Picture         =   "masterlistform.frx":1B3FB
      Style           =   1  'Graphical
      TabIndex        =   0
      Top             =   5400
      Width           =   1575
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      Height          =   975
      Left            =   120
      TabIndex        =   3
      Top             =   -120
      Width           =   8775
      Begin VB.Label Label1 
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
         TabIndex        =   11
         Top             =   240
         Width           =   735
      End
      Begin VB.Label Label2 
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
         Left            =   4320
         TabIndex        =   10
         Top             =   240
         Width           =   855
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Code:"
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
         TabIndex        =   9
         Top             =   600
         Width           =   975
      End
      Begin VB.Label Label4 
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
         Left            =   4320
         TabIndex        =   8
         Top             =   600
         Width           =   855
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
         TabIndex        =   7
         Top             =   240
         Width           =   2895
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
         Left            =   5400
         TabIndex        =   6
         Top             =   240
         Width           =   3135
      End
      Begin VB.Label lbl_subject_code 
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
         TabIndex        =   5
         Top             =   600
         Width           =   2895
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
         Left            =   5400
         TabIndex        =   4
         Top             =   600
         Width           =   3255
      End
   End
   Begin MSDataGridLib.DataGrid dg_masterlist 
      Height          =   3855
      Left            =   240
      TabIndex        =   2
      Top             =   1440
      Width           =   8535
      _ExtentX        =   15055
      _ExtentY        =   6800
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   19
      BeginProperty HeadFont {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ColumnCount     =   2
      BeginProperty Column00 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      BeginProperty Column01 
         DataField       =   ""
         Caption         =   ""
         BeginProperty DataFormat {6D835690-900B-11D0-9484-00A0C91110ED} 
            Type            =   0
            Format          =   ""
            HaveTrueFalseNull=   0
            FirstDayOfWeek  =   0
            FirstWeekOfYear =   0
            LCID            =   1033
            SubFormatType   =   0
         EndProperty
      EndProperty
      SplitCount      =   1
      BeginProperty Split0 
         BeginProperty Column00 
         EndProperty
         BeginProperty Column01 
         EndProperty
      EndProperty
   End
   Begin VB.Label Label5 
      BackStyle       =   0  'Transparent
      Caption         =   "Sort by:"
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
      Left            =   1560
      TabIndex        =   13
      Top             =   1080
      Width           =   1335
   End
   Begin VB.Label lbl_set_grade 
      BackStyle       =   0  'Transparent
      Caption         =   "Encode grade."
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
      Left            =   7200
      TabIndex        =   1
      Top             =   5640
      Width           =   1575
   End
End
Attribute VB_Name = "masterlistform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Public rs_masterlist As New ADODB.Recordset
Public rs_grade As New ADODB.Recordset

Private Sub cmb_category_Click()
     Dim col_order As String
    Select Case (cmb_category.Text)
        Case "LRN"
            col_order = "a.student_id"
        Case "Last Name"
            col_order = "a.last_name"
        Case "First Name"
            col_order = "a.first_name"
    End Select
    
    Call set_datagrid(dg_masterlist, rs_masterlist, _
                                            "SELECT @index := @index + 1 as No," _
                                                & "masterlist.* " _
                                            & "FROM " _
                                                & "(SELECT " _
                                                    & "a.student_id as LRN, " _
                                                    & "a.last_name as Last_Name, a.First_Name,a.Middle_Name " _
                                                & "FROM " _
                                                    & "tbl_student a " _
                                                & "LEFT JOIN " _
                                                    & "tbl_student_level b " _
                                                & "ON " _
                                                    & "a.student_id = b.ID " _
                                                & "WHERE " _
                                                    & "   b.section_name = '" & mysectionform.rs_section.Fields("Section").Value & "' ORDER BY " & col_order & " ASC) masterlist" _
                                            & " JOIN " _
                                                & "(SELECT @index :=0) c ")
    dg_masterlist.Columns(0).Width = 400
End Sub

Private Sub cmb_category_KeyUp(KeyCode As Integer, Shift As Integer)
    MsgBox "Please select a category from the list."
    cmb_category.Text = ""
End Sub

Private Sub cmd_print_Click()
    If rs_masterlist.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
      dr_masterlist.Sections(2).Controls("lbl_sy").Caption = mainteacherform.lbl_sy.Caption
      dr_masterlist.Sections(2).Controls("lbl_date").Caption = Format(Now, "mmmm dd, yyyy") & "/ " & Time
        dr_masterlist.Sections(2).Controls("lbl_section").Caption = lbl_level.Caption & " - " & lbl_section.Caption
        dr_masterlist.Sections(2).Controls("lbl_subject").Caption = lbl_subject_name.Caption
        dr_masterlist.Sections(2).Controls("lbl_no").Caption = rs_masterlist.RecordCount
         Set dr_masterlist.DataSource = rs_masterlist
        dr_masterlist.Show vbModal, Me
End Sub

Public Sub Form_Load()
     
     Call set_datagrid(dg_masterlist, rs_masterlist, _
                                        "SELECT @index := @index + 1 as No," _
                                            & "a.student_id as LRN, a.last_name as Last_Name, a.First_Name,a.Middle_Name FROM tbl_student a LEFT JOIN tbl_student_level b ON a.student_id = b.ID JOIN(SELECT @index :=0) c WHERE b.section_name = '" & mysectionform.rs_section.Fields("Section").Value & "'")

    

     dg_masterlist.Columns(0).Width = 400

End Sub

Private Sub lbl_set_grade_Click()
     If rs_masterlist.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
    gradeencodeform.lbl_level.Caption = lbl_level.Caption
    gradeencodeform.lbl_section.Caption = lbl_section.Caption
    gradeencodeform.lbl_subject_code.Caption = lbl_subject_code.Caption
    gradeencodeform.lbl_subject_name.Caption = lbl_subject_name.Caption
         
    Call load_form(gradeencodeform, True)
End Sub
