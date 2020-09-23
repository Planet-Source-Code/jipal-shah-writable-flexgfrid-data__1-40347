VERSION 5.00
Object = "{5E9E78A0-531B-11CF-91F6-C2863C385E30}#1.0#0"; "MSFLXGRD.OCX"
Begin VB.Form Form1 
   BackColor       =   &H80000004&
   BorderStyle     =   0  'None
   Caption         =   " "
   ClientHeight    =   8505
   ClientLeft      =   0
   ClientTop       =   255
   ClientWidth     =   12000
   ControlBox      =   0   'False
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   8505
   ScaleWidth      =   12000
   ShowInTaskbar   =   0   'False
   StartUpPosition =   3  'Windows Default
   WindowState     =   2  'Maximized
   Begin VB.TextBox Text4 
      Height          =   375
      Left            =   240
      TabIndex        =   9
      Text            =   "Text4"
      Top             =   5880
      Visible         =   0   'False
      Width           =   855
   End
   Begin MSFlexGridLib.MSFlexGrid MSFlexGrid1 
      Height          =   2895
      Left            =   120
      TabIndex        =   4
      Top             =   2640
      Width           =   11775
      _ExtentX        =   20770
      _ExtentY        =   5106
      _Version        =   393216
   End
   Begin VB.Frame Frame2 
      Caption         =   "Method"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   855
      Left            =   120
      TabIndex        =   8
      Top             =   6960
      Width           =   11655
      Begin VB.CommandButton Command2 
         BackColor       =   &H000000FF&
         Caption         =   "ABOUT ME"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   375
         Left            =   10320
         TabIndex        =   19
         Top             =   360
         Width           =   1095
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Delete"
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
         Index           =   8
         Left            =   7800
         TabIndex        =   18
         ToolTipText     =   "TOP RECORD"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "First"
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
         Index           =   0
         Left            =   240
         TabIndex        =   17
         ToolTipText     =   "TOP RECORD"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Next"
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
         Index           =   1
         Left            =   1320
         TabIndex        =   16
         ToolTipText     =   "MOVE RECORD"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Previous"
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
         Index           =   2
         Left            =   2400
         TabIndex        =   15
         ToolTipText     =   "BACK RECORD"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Last"
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
         Index           =   3
         Left            =   3480
         TabIndex        =   14
         ToolTipText     =   "TOP RECORD"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Add"
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
         Index           =   4
         Left            =   4560
         TabIndex        =   13
         ToolTipText     =   "TOP RECORD"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
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
         Height          =   375
         Index           =   5
         Left            =   5640
         TabIndex        =   12
         ToolTipText     =   "TOP RECORD"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Edit"
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
         Index           =   7
         Left            =   6720
         TabIndex        =   11
         ToolTipText     =   "TOP RECORD"
         Top             =   360
         Width           =   975
      End
      Begin VB.CommandButton Command1 
         Caption         =   "Exit"
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
         Index           =   9
         Left            =   9000
         TabIndex        =   10
         ToolTipText     =   "TOP RECORD"
         Top             =   360
         Width           =   975
      End
   End
   Begin VB.Frame Frame1 
      BackColor       =   &H80000004&
      Caption         =   "Invoice"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H00000000&
      Height          =   1335
      Left            =   120
      TabIndex        =   0
      Top             =   480
      Width           =   11775
      Begin VB.TextBox Text3 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   7440
         TabIndex        =   3
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text2 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   3960
         TabIndex        =   2
         Top             =   720
         Width           =   1815
      End
      Begin VB.TextBox Text1 
         Appearance      =   0  'Flat
         Height          =   285
         Left            =   1320
         TabIndex        =   1
         Top             =   720
         Width           =   1335
      End
      Begin VB.Label Label3 
         Caption         =   "Amount : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   6240
         TabIndex        =   7
         Top             =   720
         Width           =   975
      End
      Begin VB.Label Label2 
         Caption         =   "Name :"
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   3120
         TabIndex        =   6
         Top             =   720
         Width           =   735
      End
      Begin VB.Label Label1 
         Caption         =   "Bill No : "
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   5
         Top             =   720
         Width           =   855
      End
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim db As Database
Dim rs As Recordset
Dim rs1 As Recordset
Dim mode As String
Private Sub Command1_Click(Index As Integer)
  If Index = 0 Then
     rs.MoveFirst
     Call dreposition
     Call lproc
  ElseIf Index = 1 Then
     rs.MoveNext
     If rs.EOF Then
        rs.MoveLast
     End If
     Call dreposition
     Call lproc
 ElseIf Index = 2 Then
     rs.MovePrevious
     If rs.BOF Then
        rs.MoveFirst
     End If
     Call dreposition
     Call lproc
 ElseIf Index = 3 Then
     rs.MoveLast
     Call dreposition
     Call lproc
 ElseIf Index = 4 Then
    rs.AddNew
    Call brecord
    MSFlexGrid1.Cols = 6
    MSFlexGrid1.Rows = 2
    Call flexhead
    Call unlproc
    mode = "add"
    Text1.SetFocus
    ElseIf Index = 5 Then
    Dim i As Integer, j As Integer
    rs!invno = Text1 & ""
    rs!Name = Text2 & ""
    rs!amount = Text3 & ""
    rs.Update
    If mode = "edit" Then
        Set rs1 = db.Recordsets("select * from tans where invno = " & CInt(Text1))
        rs1.MoveLast
        rs1.MoveFirst
    End If
    For i = 1 To MSFlexGrid1.Rows - 1
        If mode = "add" Then
            rs1.AddNew
        ElseIf mode = "edit" Then
        rs1.Edit
        End If
        rs1!invno = CInt(Text1)
        For j = 1 To MSFlexGrid1.Cols - 1
            rs1.Fields(j) = MSFlexGrid1.TextMatrix(i, j)
            Next
        rs1.Update
            If mode = "edit" Then rs1.MoveNext
        Next
    Call Command1_Click(1)
    ElseIf Index = 7 Then
    Text4.Locked = False
    MSFlexGrid1.Enabled = True
    Call unlproc
    rs.Edit
    mode = "edit"
    ElseIf Index = 8 Then
    Dim ans As String
    ans = MsgBox("Are You Sure", vbYesNo + vbDefaultButton2, "Confirmation ")
    If ans = vbYes Then
        Dim temp As Integer
        temp = CInt(Text1)
        rs.Delete
    db.Execute "delete * from tans where invno = " & CInt(temp)
    If rs.RecordCount > 0 Then
    Call Command1_Click(1)
    End If
    End If
    ElseIf Index = 9 Then
    End
    End If
End Sub

Private Sub Command2_Click()
Form2.Show
End Sub

Private Sub Form_Load()
  Set db = DBEngine.Workspaces(0).OpenDatabase("c:\windows\desktop\jip\master.mdb")
  Set rs = db.OpenRecordset("master", dbOpenDynaset)
  Set rs1 = db.OpenRecordset("tans", dbOpenDynaset)
End Sub
Private Sub flexhead()
 MSFlexGrid1.Clear
 MSFlexGrid1.Row = 0
 MSFlexGrid1.Col = 1
 MSFlexGrid1.ColWidth(1) = 1500
 MSFlexGrid1.Text = "Product Code"
 MSFlexGrid1.Col = 2
 MSFlexGrid1.ColWidth(2) = 2500
 MSFlexGrid1.Text = "Product Name "
 MSFlexGrid1.Col = 3
 MSFlexGrid1.ColWidth(3) = 1000
 MSFlexGrid1.Text = "Quantity  "
 MSFlexGrid1.Col = 4
 MSFlexGrid1.ColWidth(4) = 1000
 MSFlexGrid1.Text = "Rate  "
 MSFlexGrid1.Col = 5
 MSFlexGrid1.ColWidth(5) = 1000
 MSFlexGrid1.Text = "Amount  "
End Sub
Private Sub dreposition()
      Dim str As String
      Dim i As Integer, j As Integer
      str = "select * from tans where invno = " & IIf(IsNull(rs!invno), 0, rs!invno)
     Set rs1 = db.OpenRecordset(str, dbOpenDynaset)
     If rs1.RecordCount > 0 Then
        rs1.MoveLast
        rs1.MoveFirst
     End If
     Text1 = rs!invno & ""
     Text2 = rs!Name & ""
     Text3 = rs!amount & ""
     MSFlexGrid1.Rows = rs1.RecordCount + 1
     MSFlexGrid1.Cols = rs1.Fields.Count
     Call flexhead
     For i = 1 To rs1.RecordCount
          For j = 1 To rs1.Fields.Count - 1
              MSFlexGrid1.TextMatrix(i, j) = rs1.Fields(j)
          Next
          rs1.MoveNext
     Next
End Sub
Private Sub lproc()
   Text1.Locked = True
   Text2.Locked = True
   Text3.Locked = True
   MSFlexGrid1.Enabled = False
End Sub
Private Sub unlproc()
   Text1.Locked = False
   Text2.Locked = False
   Text3.Locked = False
   MSFlexGrid1.Enabled = True
End Sub
Private Sub brecord()
 Text1 = ""
 Text2 = ""
 Text3 = ""
End Sub

Private Sub MSFlexGrid1_KeyPress(keyascii As Integer)
   eedit keyascii
End Sub
Private Sub eedit(keyascii As Integer)
   Select Case keyascii
           Case 0 To Asc(" ")
                Text4 = MSFlexGrid1
                Text4.SelStart = 1000
           Case Else
                Text4 = Chr(keyascii)
                Text4.SelStart = 1
   End Select
   With Text4
        .FontName = MSFlexGrid1.CellFontName
        .FontSize = MSFlexGrid1.CellFontSize
        .Left = MSFlexGrid1.Left + MSFlexGrid1.CellLeft
        .Top = MSFlexGrid1.Top + MSFlexGrid1.CellTop
        .Width = MSFlexGrid1.CellWidth
        .Height = MSFlexGrid1.CellHeight
        .Visible = True
        .SetFocus
   End With
End Sub
Private Sub Text3_LostFocus()
MSFlexGrid1.Row = 1
MSFlexGrid1.Col = 1
MSFlexGrid1.SetFocus
End Sub
Private Sub Text4_KeyDown(KeyCode As Integer, Shift As Integer)
    Dim X As Integer, i As Integer
    i = MSFlexGrid1.Row
    If KeyCode = vbKeyReturn And MSFlexGrid1.Col < MSFlexGrid1.Cols - 1 Then
       MSFlexGrid1 = Text4
       Text4.Visible = False
       MSFlexGrid1.Col = MSFlexGrid1.Col + 1
       MSFlexGrid1.SetFocus
       If MSFlexGrid1.Col > 4 Then
            MSFlexGrid1.TextMatrix(i, 5) = MSFlexGrid1.TextMatrix(i, 3) * MSFlexGrid1.TextMatrix(i, 4)
       End If
    ElseIf KeyCode = vbKeyReturn And MSFlexGrid1.Col = MSFlexGrid1.Cols - 1 Then
       MSFlexGrid1 = Text4
       Text4.Visible = False
       X = MsgBox("Add More Data", vbYesNo + vbDefaultButton1 + vbQuestion, "Confirmation")
       If X = vbYes Then
          MSFlexGrid1.Rows = MSFlexGrid1.Rows + 1
          MSFlexGrid1.Row = MSFlexGrid1.Row + 1
          MSFlexGrid1.Col = 1
       End If
    End If
End Sub
