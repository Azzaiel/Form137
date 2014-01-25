VERSION 5.00
Object = "{CDE57A40-8B86-11D0-B3C6-00A0C90AEA82}#1.0#0"; "MSDATGRD.OCX"
Begin VB.Form userform 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "List of Users"
   ClientHeight    =   5985
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   8940
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   Picture         =   "userform.frx":0000
   ScaleHeight     =   5985
   ScaleWidth      =   8940
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox txt_search 
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   13.5
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   960
      TabIndex        =   0
      Top             =   240
      Width           =   5775
   End
   Begin VB.CommandButton cmd_search 
      Height          =   615
      Left            =   6960
      Picture         =   "userform.frx":1B3CE
      Style           =   1  'Graphical
      TabIndex        =   1
      Top             =   240
      Width           =   1095
   End
   Begin VB.CommandButton cmd_edit 
      Height          =   615
      Left            =   3720
      Picture         =   "userform.frx":1C21D
      Style           =   1  'Graphical
      TabIndex        =   4
      Top             =   5280
      Width           =   1095
   End
   Begin MSDataGridLib.DataGrid dg_users 
      Height          =   3975
      Left            =   360
      TabIndex        =   5
      Top             =   1200
      Width           =   8295
      _ExtentX        =   14631
      _ExtentY        =   7011
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
   Begin VB.Label Label1 
      BackStyle       =   0  'Transparent
      Caption         =   "Security Questions"
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
      Left            =   360
      TabIndex        =   2
      Top             =   5400
      Width           =   2895
   End
   Begin VB.Label lbl_default_username 
      Alignment       =   1  'Right Justify
      BackStyle       =   0  'Transparent
      Caption         =   "View user log history."
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
      Left            =   6000
      TabIndex        =   3
      Top             =   5400
      Width           =   2535
   End
   Begin VB.Label Label2 
      BackStyle       =   0  'Transparent
      Caption         =   "List of Users."
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
      TabIndex        =   6
      Top             =   960
      Width           =   3735
   End
End
Attribute VB_Name = "userform"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Dim rs_user As New ADODB.Recordset
Dim usertype, id As String

Private Sub cmd_edit_Click()
If rs_user.RecordCount = 0 Then
    MsgBox "No record selected."
Else
    usereditform.txt_username = rs_user.Fields("Username")
    usereditform.txt_usertype = rs_user.Fields("User_Type")
    usertype = rs_user.Fields("User_Type")
    Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & rs_user.Fields("Username") & "'")
    id = public_rs.Fields("ID")
    usereditform.txt_id.Text = id
    usereditform.txt_oldid.Text = id
    usereditform.txt_oldusername.Text = rs_user.Fields("Username")
    If usertype = "Administrator" Then
        Call mysql_select(public_rs, "SELECT * FROM tbl_user WHERE Username = '" & rs_user.Fields("Username") & "'")
        usereditform.txt_lastname.Text = public_rs.Fields("Lastname")
        usereditform.txt_firstname.Text = public_rs.Fields("Firstname")
        usereditform.txt_middlename.Text = public_rs.Fields("Middlename")
         usereditform.txt_id.Enabled = True
        usereditform.txt_lastname.Enabled = True
         usereditform.txt_firstname.Enabled = True
          usereditform.txt_middlename.Enabled = True
          usereditform.lbl_password1.Visible = True
          usereditform.lbl_password2.Visible = True
          usereditform.lbl_password3.Visible = True
          usereditform.txt_confirm_password.Visible = True
          usereditform.txt_new_password.Visible = True
          usereditform.txt_old_password.Visible = True
          
    Else
        Call mysql_select(public_rs, "SELECT * FROM tbl_teacher WHERE teacher_id = '" & id & "'")
        usereditform.txt_lastname.Text = public_rs.Fields("last_name")
        usereditform.txt_firstname.Text = public_rs.Fields("first_name")
        usereditform.txt_middlename.Text = public_rs.Fields("middle_name")
        usereditform.txt_id.Enabled = False
         usereditform.txt_lastname.Enabled = False
         usereditform.txt_firstname.Enabled = False
          usereditform.txt_middlename.Enabled = False
          usereditform.txt_username.Enabled = False
          usereditform.txt_confirm_password.Enabled = False
          usereditform.txt_new_password.Enabled = False
          usereditform.txt_old_password.Enabled = False
          usereditform.lbl_password1.Visible = True
          usereditform.lbl_password2.Visible = False
          usereditform.lbl_password3.Visible = False
          usereditform.txt_confirm_password.Visible = False
          usereditform.txt_new_password.Visible = False
          usereditform.txt_old_password.Visible = True
          
    End If
     Call load_form(usereditform, True)
End If
End Sub

Private Sub cmd_new_Click()
    useraccountform.txt_id.Text = ""
    useraccountform.txt_firstname.Text = ""
    useraccountform.txt_lastname.Text = ""
    useraccountform.txt_middlename.Text = ""
    useraccountform.txt_usertype.Text = "Administrator"
    useraccountform.txt_username.Text = ""
    useraccountform.txt_password.Text = ""
    Call load_form(useraccountform, True)
End Sub

Private Sub cmd_search_Click()
     Call set_datagrid(dg_users, rs_user, _
                                        "SELECT " _
                                            & "Usertype as User_Type, Username FROM tbl_user WHERE Usertype = '" & txt_search.Text & "' OR Username = '" & txt_search.Text & "' ORDER BY Usertype ASC ")
                                        
                    
      If rs_user.RecordCount = 0 Then
        MsgBox "Record not found."
      End If
End Sub

Public Sub Form_Load()
      Call set_datagrid(dg_users, rs_user, _
                                        "SELECT " _
                                            & "Usertype as User_Type, Username FROM tbl_user ORDER BY Usertype ASC")
                                        
                    
                                       
End Sub

Private Sub Label1_Click()
     Call mysql_select(public_rs, "SELECT * FROM tbl_security WHERE Username = '" & user & "'")
     If public_rs.RecordCount = 0 Then
        securityquestions.txt_pet.Text = ""
        securityquestions.txt_place.Text = ""
        securityquestions.txt_author.Text = ""
    Else
        securityquestions.txt_pet.Text = public_rs.Fields("Pet").Value
        securityquestions.txt_place.Text = public_rs.Fields("Place").Value
        securityquestions.txt_author.Text = public_rs.Fields("Author").Value
    End If
     Call load_form(securityquestions, True)
End Sub

Private Sub lbl_default_username_Click()
    If rs_user.RecordCount = 0 Then
        MsgBox "No record."
        Exit Sub
    End If
    Call load_form(userlogform, True)
End Sub

Private Sub txt_search_KeyUp(KeyCode As Integer, Shift As Integer)
      Call set_datagrid(dg_users, rs_user, _
                                        "SELECT " _
                                            & "Usertype as User_Type, Username FROM tbl_user WHERE Usertype LIKE '%" & txt_search.Text & "%' OR Username LIKE '%" & txt_search.Text & "%' ORDER BY Usertype ASC ")
                                        
                    
                                   
End Sub
