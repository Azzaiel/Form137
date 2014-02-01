VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form bulkSubjEncode 
   BackColor       =   &H8000000E&
   ClientHeight    =   7635
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   10005
   LinkTopic       =   "Form1"
   ScaleHeight     =   7635
   ScaleWidth      =   10005
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmb_export 
      Caption         =   "Export"
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
      TabIndex        =   1
      Top             =   6960
      Width           =   1215
   End
   Begin MSFlexGridLib.MSFlexGrid flexGradeBoys 
      Height          =   2415
      Left            =   0
      TabIndex        =   10
      Top             =   1440
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4260
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
   Begin VB.CommandButton cmd_close 
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
      Left            =   5640
      TabIndex        =   2
      Top             =   6960
      Width           =   1215
   End
   Begin VB.CommandButton cmd_save 
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
      Left            =   3000
      TabIndex        =   0
      Top             =   6960
      Width           =   1215
   End
   Begin VB.Frame Frame3 
      BackColor       =   &H00808080&
      Height          =   855
      Left            =   0
      TabIndex        =   3
      Top             =   0
      Width           =   9975
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
         Left            =   4200
         TabIndex        =   9
         Top             =   480
         Width           =   8895
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
         Left            =   3240
         TabIndex        =   8
         Top             =   480
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
         Left            =   6840
         TabIndex        =   7
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
         TabIndex        =   6
         Top             =   240
         Width           =   2055
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
         Left            =   5880
         TabIndex        =   5
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
         TabIndex        =   4
         Top             =   240
         Width           =   735
      End
   End
   Begin MSFlexGridLib.MSFlexGrid flexGradeGirls 
      Height          =   2415
      Left            =   0
      TabIndex        =   12
      Top             =   4440
      Width           =   9975
      _ExtentX        =   17595
      _ExtentY        =   4260
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
   Begin VB.Label Label5 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Girls"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   13
      Top             =   4080
      Width           =   9975
   End
   Begin VB.Label Label4 
      Alignment       =   2  'Center
      BackColor       =   &H00808080&
      Caption         =   "Boys"
      BeginProperty Font 
         Name            =   "MS Sans Serif"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00FFFFFF&
      Height          =   375
      Left            =   0
      TabIndex        =   11
      Top             =   1080
      Width           =   9975
   End
End
Attribute VB_Name = "bulkSubjEncode"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Public rs_grades As New ADODB.Recordset
Public rs_tmp As New ADODB.Recordset
Public subj_code As String
Private temp_grades(0 To 3) As Double
Private temp As Integer

Private Sub cmd_add_Click()

End Sub

Private Sub cmb_export_Click()
  
  Dim excelApp As New Excel.Application
  Dim oBook As New Excel.Workbook
  Dim oSheet As New Excel.Worksheet
  
  Set excelApp = CreateObject("Excel.Application")
  Set oBook = excelApp.Workbooks.Open(CommonHelper.getTemplatesPath & "\Stud_Subj_Grade")
  Set oSheet = excelApp.Worksheets(1)
  
  excelApp.DisplayAlerts = False
  oBook.SaveAs CommonHelper.getTempPath & "\tmp.xlsx"

  excelApp.Visible = True

End Sub

Private Sub cmd_close_Click()
  Unload Me
End Sub
Private Sub updateGrade(grade As Double, lrn As String, period As String)

  Dim isKinder As Boolean
  Dim gradeChanged As Boolean
  gradeChanged = False
  
  If (lbl_level = "Kinder") Then
    isKinder = True
  Else
    isKinder = False
  End If
    
    If (rs_tmp.RecordCount > 0) Then
      If (val(rs_tmp!grade) <> val(grade)) Then
        rs_tmp!grade = grade
        rs_tmp!remark = mod_grade.getRemark(val(grade), isKinder)
        rs_tmp.Update
        gradeChanged = True
      End If
    Else
      gradeChanged = True
      rs_tmp.AddNew
      rs_tmp!id = lrn
      rs_tmp!SY = mainteacherform.cmb_sy.Text
      rs_tmp!section_name = lbl_section
      rs_tmp!SUBJECT_CODE = subj_code
      rs_tmp!period = period
      rs_tmp!grade = grade
      rs_tmp!remark = mod_grade.getRemark(val(grade), isKinder)
      rs_tmp.Update
    End If
  
  If (gradeChanged And isKinder = False And subj_code = "Edukasyon sa Pagpapakatao") Then
    Dim sql_query As String
    sql_query = "Select * " & _
                "From tbl_character_grade " & _
                "Where  SY = '" & mainteacherform.cmb_sy.Text & "' " & _
                "       and ID = '" & lrn & "' " & _
                "       and section_name = '" & lbl_section & "' " & _
                "       and period = '" & period & "' "
    Call mysql_select(rs_grades, sql_query)
    
    If (rs_grades.RecordCount = 0) Then
      rs_grades.AddNew
      rs_grades!SY = mainteacherform.cmb_sy.Text
      rs_grades!id = lrn
      rs_grades!section_name = lbl_section
      rs_grades!period = period
    End If
    Dim charRemark  As String
    
    charRemark = mod_grade.getCharacterRemark(val(grade))
    rs_grades!Honesty = charRemark
    rs_grades!Courtesy = charRemark
    rs_grades!Helpfulness_and_Cooperation = charRemark
    rs_grades!Resourcefulness_and_Creativity = charRemark
    rs_grades!Consideration_for_Others = charRemark
    rs_grades!Sportsmanship = charRemark
    rs_grades!Obedience = charRemark
    rs_grades!Self_Reliance = charRemark
    rs_grades!Industry = charRemark
    rs_grades!Cleanliness_and_Orderliness = charRemark
    rs_grades!Promptness_and_Punctuality = charRemark
    rs_grades!Sense_of_Responsibility = charRemark
    rs_grades!Love_of_God = charRemark
    rs_grades!Patriotism_and_Love_of_Country = charRemark
    rs_grades.Update
    
  End If
  
End Sub
Private Sub saveFlexData(flexGrid As MSFlexGrid)
  Dim index As Integer
  Dim cur_lrn As String
  Dim cur_period As String
  
  With flexGrid
    For index = 1 To (flexGrid.Rows - 1)
      
      .Row = index
      cur_lrn = .TextMatrix(index, 1)
      
      cur_period = "1st Grading"
      Call mysql_select(rs_tmp, generatePeriodSelectGradeQuery(cur_lrn, cur_period))
      .Col = 3
      Call updateGrade(val(.Text), cur_lrn, cur_period)
      
      
      cur_period = "2nd Grading"
      Call mysql_select(rs_tmp, generatePeriodSelectGradeQuery(cur_lrn, cur_period))
      .Col = 4
      Call updateGrade(val(.Text), cur_lrn, cur_period)
      
      
      cur_period = "3rd Grading"
      Call mysql_select(rs_tmp, generatePeriodSelectGradeQuery(cur_lrn, cur_period))
      .Col = 5
      Call updateGrade(val(.Text), cur_lrn, cur_period)
      
      cur_period = "4th Grading"
      Call mysql_select(rs_tmp, generatePeriodSelectGradeQuery(cur_lrn, cur_period))
      .Col = 6
      Call updateGrade(val(.Text), cur_lrn, cur_period)
      
     .Col = 7
      Call updateGrade(val(.Text), cur_lrn, "Final")
      
    Next index
  End With
  
End Sub


Private Sub cmd_save_Click()
  Call saveFlexData(flexGradeBoys)
  Call saveFlexData(flexGradeGirls)
  MsgBox "Record Updated", vbInformation
  Call populateGrades
End Sub
Private Function generatePeriodSelectGradeQuery(lrn As String, period As String) As String
  Dim sql_query As String
  sql_query = "Select SY, ID, SECTION_NAME, SUBJECT_CODE, PERIOD, GRADE, REMARK " & _
              "From tbl_grade " & _
              "Where ID = '" & lrn & "' " & _
              "      And SY = '" & mainteacherform.cmb_sy.Text & "' " & _
              "      And SECTION_NAME = '" & lbl_section & "' " & _
              "      And PERIOD = '" & period & "'" & _
              "      aND SUBJECT_CODE = '" & subj_code & "' "
  generatePeriodSelectGradeQuery = sql_query
End Function
Private Sub encodeFlexData(KeyAscii As Integer, flexGrid As MSFlexGrid)
    With flexGrid
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
        temp = .Col
        .Col = 3
        temp_grades(0) = val(.Text)
        .Col = 4
        temp_grades(1) = val(.Text)
        .Col = 5
        temp_grades(2) = val(.Text)
        .Col = 6
        temp_grades(3) = val(.Text)
        .TextMatrix(.Row, 7) = mod_grade.getFinalGrade(temp_grades)
        .Col = temp
    End With

End Sub

Private Sub Command1_Click()

End Sub

Private Sub flexGradeBoys_KeyPress(KeyAscii As Integer)
  Call encodeFlexData(KeyAscii, flexGradeBoys)
End Sub
Public Sub populateGradeFlex(flexGrid As MSFlexGrid, rs As ADODB.Recordset)
    Dim index As String
  index = 1
  
  With flexGrid
    .Clear
    .Rows = rs.RecordCount + 1
    .Cols = 8
    
    .TextMatrix(0, 0) = "NO"
    .TextMatrix(0, 1) = "LRN"
    .TextMatrix(0, 2) = "NAME"
    .TextMatrix(0, 3) = "1"
    .TextMatrix(0, 4) = "2"
    .TextMatrix(0, 5) = "3"
    .TextMatrix(0, 6) = "4"
    .TextMatrix(0, 7) = "FINALS"

    .ColAlignment(0) = flexAlignCenterCenter
    .ColWidth(0) = 500
    .ColAlignment(1) = flexAlignCenterCenter
    .ColWidth(1) = 1650
    .ColWidth(2) = 3000
    .ColAlignment(3) = flexAlignCenterCenter
    .ColWidth(3) = 900
    .ColAlignment(4) = flexAlignCenterCenter
    .ColWidth(4) = 900
    .ColAlignment(5) = flexAlignCenterCenter
    .ColWidth(5) = 900
    .ColAlignment(6) = flexAlignCenterCenter
    .ColWidth(6) = 900
    .ColAlignment(7) = flexAlignCenterCenter
    .ColWidth(7) = 900

    Dim grades(0 To 3) As Double
    
    While Not rs.EOF
     
      .TextMatrix(index, 0) = index
      .TextMatrix(index, 1) = rs!lrn
      .TextMatrix(index, 2) = rs!Name
      .Row = index
      
      .Col = 3
      .Text = CommonHelper.extractStringValue(rs!First_Grading)
      grades(0) = val(.Text)
      
      .Col = 4
      .Text = CommonHelper.extractStringValue(rs!Second_Grading)
      grades(1) = val(.Text)
      
      .Col = 5
      .Text = CommonHelper.extractStringValue(rs!Third_Grading)
      grades(2) = val(.Text)
      
      .Col = 6
      .Text = CommonHelper.extractStringValue(rs!Fourth_Grading)
      grades(3) = val(.Text)
      
      .TextMatrix(index, 7) = mod_grade.getFinalGrade(grades)
      
      rs.MoveNext
      index = index + 1
    Wend
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
              "      And b.LVL_NAME = '" & lbl_level & "' " & _
              "      And b.SECTION_NAME = '" & lbl_section & "' "
  
  Call mysql_select(rs_grades, sql_query & "And a.Gender = 'Male' ")
  Call populateGradeFlex(flexGradeBoys, rs_grades)
  
  Call mysql_select(rs_grades, sql_query & "And a.Gender = 'Female' ")
  Call populateGradeFlex(flexGradeGirls, rs_grades)
  
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

Private Sub flexGradeBoysBoys_Click()

End Sub

Private Sub MSFlexGrid1_Click()

End Sub

Private Sub flexGradeGirls_KeyPress(KeyAscii As Integer)
  Call encodeFlexData(KeyAscii, flexGradeGirls)
End Sub

