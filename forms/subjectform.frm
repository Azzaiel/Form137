VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form subjectform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Add Subject Per Level"
   ClientHeight    =   6000
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7815
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "subjectform.frx":0000
   ScaleHeight     =   6000
   ScaleWidth      =   7815
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_oldname 
      Height          =   375
      Left            =   7440
      TabIndex        =   12
      Top             =   2040
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.CommandButton cmd_settings 
      Height          =   495
      Left            =   5640
      Picture         =   "subjectform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   5
      Top             =   5400
      Width           =   1815
   End
   Begin VB.TextBox txt_op 
      Height          =   375
      Left            =   7440
      TabIndex        =   10
      Top             =   1440
      Visible         =   0   'False
      Width           =   1335
   End
   Begin VB.TextBox txt_subject_name 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   3000
      MaxLength       =   35
      TabIndex        =   1
      Top             =   3960
      Width           =   3375
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H00808080&
      BorderStyle     =   0  'None
      Height          =   375
      Left            =   120
      TabIndex        =   7
      Top             =   120
      Width           =   7695
      Begin VB.ComboBox cmb_level 
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
         Left            =   2640
         TabIndex        =   0
         Top             =   0
         Width           =   3375
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Level Name:"
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
         TabIndex        =   8
         Top             =   120
         Width           =   1455
      End
   End
   Begin VB.CommandButton cmb_clear 
      Height          =   615
      Left            =   4680
      Picture         =   "subjectform.frx":1C591
      Style           =   1  'Graphical
      TabIndex        =   3
      Top             =   4560
      Width           =   1095
   End
   Begin VB.CommandButton cmd_save 
      Height          =   615
      Left            =   3360
      Picture         =   "subjectform.frx":1D30C
      Style           =   1  'Graphical
      TabIndex        =   2
      Top             =   4560
      Width           =   1095
   End
   Begin VB.TextBox txt_oldcode 
      Height          =   375
      Left            =   7440
      TabIndex        =   6
      Top             =   840
      Visible         =   0   'False
      Width           =   1335
   End
   Begin MSDataGridLib.DataGrid dg_subjects 
      Height          =   2655
      Left            =   360
      TabIndex        =   4
      Top             =   1080
      Width           =   5655
      _ExtentX        =   9975
      _ExtentY        =   4683
      _Version        =   393216
      AllowUpdate     =   0   'False
      HeadLines       =   1
      RowHeight       =   24
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
         Size            =   12
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
   Begin VB.Label Label4 
      BackStyle       =   0  'Transparent
      Caption         =   "Double click to edit a subject."
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   -1  'True
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   480
      TabIndex        =   11
      Top             =   2400
      Width           =   3735
   End
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "*Subject Name:"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   255
      Left            =   960
      TabIndex        =   9
      Top             =   4080
      Width           =   1815
   End
End
Attribute VB_Name = "subjectform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_subject As New ADODB.Recordset
Private Sub cmb_clear_Click()
    txt_op.Text = "add"
    'txt_subject_code.Text = ""
    txt_subject_name.Text = ""
End Sub
Private Sub cmb_level_Change()
    'txt_subject_code.Text = ""
    txt_subject_name.Text = ""
         Call set_datagrid(dg_subjects, rs_subject, _
                                        "SELECT " _
                                            & "subject_code as Subject_Code,subject_name as Subject_Name, last_mod_date " _
                                        & "FROM " _
                                            & "tbl_subject  " _
                                        & "WHERE " _
                                            & "lvl_name = '" & cmb_level.Text & "' " _
                                        & "Order by last_mod_date desc")
   Call formatDataGrid
End Sub
Private Sub cmb_level_Click()
  txt_subject_code.Text = ""
    txt_subject_name.Text = ""
    Call set_datagrid(dg_subjects, rs_subject, _
                                        "SELECT " _
                                            & "subject_code as Subject_Code,subject_name as Subject_Name " _
                                        & "FROM " _
                                            & "tbl_subject  " _
                                        & "WHERE " _
                                            & "lvl_name = '" & cmb_level.Text & "'")
End Sub

Private Sub cmd_save_Click()
    Dim ans As String
    If txt_op.Text = "add" Then
        If txt_subject_name.Text <> "" Then
            Call mysql_select(public_rs, "SELECT * FROM tbl_subject WHERE lvl_name = '" & cmb_level.Text & "'AND subject_code = '" & txt_subject_code.Text & "'")
            If public_rs.RecordCount = 0 Then
                 ans = MsgBox("Are you sure you want to add the subject", vbYesNo, "Add Subject")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                Call mysql_select(rs_subject, "INSERT INTO tbl_subject(lvl_name,subject_code,subject_name, last_mod_date) VALUES ('" & cmb_level.Text & "','" & txt_subject_code.Text & "', '" & txt_subject_name.Text & "', sysdate())")
                MsgBox "Subject successfully added!"
                txt_subject_code.Text = ""
                txt_subject_name.Text = ""
                level = cmb_level.Text
                Call Form_Load
                End If
            Else
                MsgBox "Subject code already exists."
            End If
        Else
            MsgBox "Please complete all fields."
        End If
    Else
        If txt_subject_name.Text <> "" Then
                Call mysql_select(public_rs, "SELECT * FROM tbl_subject WHERE lvl_name = '" & cmb_level.Text & "'AND subject_code = '" & cmb_level.Text & "'")
                If public_rs.RecordCount = 0 Then
                     ans = MsgBox("Are you sure you want to update the subject?", vbYesNo, "Update Subject")
                    If ans = vbNo Then
                        Exit Sub
                    Else
                    Call mysql_select(rs_subject, "UPDATE tbl_subject SET last_mod_date = sysdate(), subject_code='" & txt_subject_name.Text & "',subject_name='" & txt_subject_name.Text & "' WHERE lvl_name = '" & cmb_level.Text & "'AND subject_code = '" & txt_oldcode.Text & "'")
                    MsgBox "Subject successfully updated!"
                    'txt_subject_code.Text = ""
                    txt_subject_name.Text = ""
                    txt_op.Text = "add"
                    level = cmb_level.Text
                    Call Form_Load
                    End If
                Else
                    MsgBox "Subject code already exists."
                End If
        Else
              MsgBox "Please complete all fields."
        End If
End If
End Sub

Private Sub cmd_settings_Click()
    level = cmb_level.Text
    Call load_form(sectionform, True)
    Unload Me
End Sub

Private Sub dg_subjects_DblClick()
    txt_op.Text = "edit"
    'txt_subject_code.Text = rs_subject.Fields("Subject_Code")
    txt_oldcode.Text = rs_subject.Fields("Subject_Code")
    txt_subject_name.Text = rs_subject.Fields("Subject_Name")
    txt_oldname.Text = rs_subject.Fields("Subject_Name")
End Sub

Private Sub Form_Load()
    Call mysql_select(public_rs, "SELECT * FROM tbl_level ")
    cmb_level.Clear
    While Not public_rs.EOF
        cmb_level.AddItem (public_rs.Fields("lvl_name"))
        public_rs.MoveNext
    Wend
    If Not level = "" Then
        cmb_level.Text = level
        level = ""
        Call cmb_level_Change
    End If
    txt_op.Text = "add"
     Call set_datagrid(dg_subjects, rs_subject, _
                                        "SELECT " _
                                            & "subject_code as Subject_Code,subject_name as Subject_Name, last_mod_date " _
                                        & "FROM " _
                                            & "tbl_subject  " _
                                        & "WHERE " _
                                            & "lvl_name = '" & cmb_level.Text & "' " _
                                        & "Order by last_mod_date desc")
    Call formatDataGrid
End Sub

Private Sub formatDataGrid()
   dg_subjects.Columns(0).Visible = False
   dg_subjects.Columns(2).Visible = False
End Sub
