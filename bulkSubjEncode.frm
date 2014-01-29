VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bulkSubjEncode 
   BackColor       =   &H8000000E&
   ClientHeight    =   6510
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   12825
   LinkTopic       =   "Form1"
   ScaleHeight     =   6510
   ScaleWidth      =   12825
   StartUpPosition =   2  'CenterScreen
   Begin MSFlexGridLib.MSFlexGrid flexGrade 
      Height          =   4815
      Left            =   240
      TabIndex        =   9
      Top             =   840
      Width           =   12375
      _ExtentX        =   21828
      _ExtentY        =   8493
      _Version        =   393216
      BeginProperty Font {0BE35203-8F91-11CE-9DE3-00AA004BB851} 
         Name            =   "MS Sans Serif"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
   End
   Begin VB.CommandButton Command1 
      Caption         =   "Close"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   6960
      TabIndex        =   8
      Top             =   5760
      Width           =   1215
   End
   Begin VB.CommandButton cmd_add 
      Caption         =   "Save"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   8.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   4320
      TabIndex        =   7
      Top             =   5760
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Height          =   735
      Left            =   0
      TabIndex        =   0
      Top             =   0
      Width           =   12855
      Begin VB.Label lbl_subject 
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
         Left            =   9600
         TabIndex        =   6
         Top             =   240
         Width           =   5775
      End
      Begin VB.Label Label3 
         BackStyle       =   0  'Transparent
         Caption         =   "Subject:"
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
         Left            =   8640
         TabIndex        =   5
         Top             =   240
         Width           =   855
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
         TabIndex        =   4
         Top             =   240
         Width           =   3135
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
         TabIndex        =   3
         Top             =   240
         Width           =   2895
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
         TabIndex        =   2
         Top             =   240
         Width           =   855
      End
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
         TabIndex        =   1
         Top             =   240
         Width           =   735
      End
   End
End
Attribute VB_Name = "bulkSubjEncode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs_grades As New ADODB.Recordset
Public subj_code As String

Private Sub flexGrade_KeyPress(KeyAscii As Integer)
      With flexGrade
        Select Case KeyAscii
            Case 8
                If Not .Text = "" Then
                    .Text = Left(.Text, Len(.Text) - 1)
                End If
            Case 9 ' Tab
                If .Col + 1 = .Cols Then
                    .Col = 0
                    If .Row + 1 = .Rows Then
                        .Row = 0
                    Else
                        .Row = .Row + 1
                    End If
                Else
                    .Col = .Col + 1
                End If
            Case Else
                If (Len(.Text) < 3 And (CommonHelper.isNumberAscii(KeyAscii) Or CommonHelper.isFunctionAscii(KeyAscii))) Then
                   .Text = .Text & Chr(KeyAscii)
                End If
        End Select
    End With
End Sub
Public Sub populateGrades()
  Dim sql_query As String
  sql_query = "Select a.student_id as LRN, a.GENDER, concat(a.LAST_NAME, ', ', a.FIRST_NAME)  as Name " & _
              "      , " & generateGradePeriodQuery("1st Grading") & "as First_Grading " & _
              "      , " & generateGradePeriodQuery("2nd Grading") & "as Second_Grading " & _
              "      , " & generateGradePeriodQuery("3rd Grading") & "as Third_Grading " & _
              "      , " & generateGradePeriodQuery("4th Grading") & "as Fourth_Grading " & _
              "From tbl_student a, tbl_student_level b " & _
              "Where b.ID = a.STUDENT_ID " & _
              "      And b.SY= '" & mainteacherform.cmb_sy.Text & "' " & _
              "      And b.LVL_NAME = '" & masterlistadvisoriesform.lbl_level & "' " & _
              "      And b.SECTION_NAME = '" & masterlistadvisoriesform.lbl_section & "' " & _
              "ORDER By a.gender desc"
  Call mysql_select(rs_grades, sql_query)
  
  Dim index As String
  index = 1
  With flexGrade
    
    .Rows = rs_grades.RecordCount + 1
    .Cols = 7
    
    .TextMatrix(0, 0) = "LRN"
    .TextMatrix(0, 1) = "GENDER"
    .TextMatrix(0, 2) = "NAME"
    .TextMatrix(0, 3) = "1ST GRADING"
    .TextMatrix(0, 4) = "2ND GRADING"
    .TextMatrix(0, 5) = "3RD GRADING"
    .TextMatrix(0, 6) = "4TH GRADING"

    .ColWidth(0) = 1450
    .ColAlignment(1) = flexAlignCenterCenter
    .ColWidth(1) = 1100
    .ColWidth(2) = 3000
    .ColAlignment(3) = flexAlignCenterCenter
    .ColWidth(3) = 1650
    .ColAlignment(4) = flexAlignCenterCenter
    .ColWidth(4) = 1650
    .ColAlignment(5) = flexAlignCenterCenter
    .ColWidth(5) = 1650
    .ColAlignment(6) = flexAlignCenterCenter
    .ColWidth(6) = 1650

    
    While Not rs_grades.EOF
    
      .TextMatrix(index, 0) = rs_grades!LRN
      .TextMatrix(index, 1) = rs_grades!gender
      .TextMatrix(index, 2) = rs_grades!Name
      .Row = index
      
      .Col = 3
      .Text = CommonHelper.extractStringValue(rs_grades!First_Grading)
      
      .Col = 4
      .Text = CommonHelper.extractStringValue(rs_grades!Second_Grading)
      
      .Col = 5
      .Text = CommonHelper.extractStringValue(rs_grades!Third_Grading)
      
      .Col = 6
      .Text = CommonHelper.extractStringValue(rs_grades!Fourth_Grading)
      
      rs_grades.MoveNext
      index = index + 1
    Wend
  End With
  
End Sub
Private Sub MSFlexGrid1_KeyPress(KeyAscii As Integer)

End Sub

Private Function generateGradePeriodQuery(period As String) As String
  Dim sql_query As String
  sql_query = "(Select GRADE from tbl_grade " & _
              " Where ID = a.student_id and period = '" & period & "' " & _
              "       And SY = '" & mainteacherform.cmb_sy & "'" & _
              "       And subject_code = '" & subj_code & "' " & _
              ") "
  generateGradePeriodQuery = sql_query
End Function

