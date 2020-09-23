VERSION 5.00
Begin VB.Form calci 
   BackColor       =   &H00F1E065&
   BorderStyle     =   1  'Fixed Single
   Caption         =   "SciFi Claculator"
   ClientHeight    =   4965
   ClientLeft      =   585
   ClientTop       =   1770
   ClientWidth     =   5760
   DrawMode        =   6  'Mask Pen Not
   Icon            =   "frm_calci.frx":0000
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   Picture         =   "frm_calci.frx":030A
   ScaleHeight     =   4965
   ScaleWidth      =   5760
   StartUpPosition =   2  'CenterScreen
   Begin VB.TextBox Text1 
      BackColor       =   &H00000000&
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      Enabled         =   0   'False
      BeginProperty Font 
         Name            =   "Microsoft Sans Serif"
         Size            =   11.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H0080FFFF&
      Height          =   375
      Left            =   120
      Locked          =   -1  'True
      MaxLength       =   25
      MousePointer    =   1  'Arrow
      TabIndex        =   3
      Text            =   "SciFi Calciculator"
      Top             =   120
      Width           =   3015
   End
   Begin VB.TextBox txtpint 
      Appearance      =   0  'Flat
      Enabled         =   0   'False
      Height          =   285
      Left            =   4920
      TabIndex        =   2
      Top             =   600
      Visible         =   0   'False
      Width           =   375
   End
   Begin VB.TextBox txtmem 
      Appearance      =   0  'Flat
      Height          =   375
      Left            =   2880
      TabIndex        =   0
      Top             =   120
      Visible         =   0   'False
      Width           =   2775
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   19
      Left            =   3480
      Top             =   4200
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   18
      Left            =   5040
      Top             =   4200
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   17
      Left            =   4320
      Top             =   4200
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   495
      Index           =   16
      Left            =   2760
      Top             =   4200
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   15
      Left            =   5040
      Top             =   3480
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   14
      Left            =   4320
      Top             =   3480
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   13
      Left            =   3480
      Top             =   3480
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   12
      Left            =   2760
      Top             =   3480
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   11
      Left            =   5040
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   10
      Left            =   4320
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   9
      Left            =   3480
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   8
      Left            =   2760
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   7
      Left            =   5040
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   6
      Left            =   4200
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   5
      Left            =   3480
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   4
      Left            =   2760
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   3
      Left            =   4920
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   2
      Left            =   4200
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   1
      Left            =   3480
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image Image3 
      Height          =   615
      Index           =   0
      Left            =   2760
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   10
      Left            =   1080
      Top             =   3480
      Width           =   1215
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   9
      Left            =   240
      Top             =   3360
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   8
      Left            =   1800
      Top             =   2760
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   7
      Left            =   960
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   6
      Left            =   240
      Top             =   2760
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   5
      Left            =   1800
      Top             =   2040
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   4
      Left            =   960
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   495
      Index           =   3
      Left            =   240
      Top             =   2040
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   2
      Left            =   1800
      Top             =   1320
      Width           =   495
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   1
      Left            =   960
      Top             =   1320
      Width           =   615
   End
   Begin VB.Image Image2 
      Height          =   615
      Index           =   0
      Left            =   240
      Top             =   1320
      Width           =   615
   End
   Begin VB.Label lblmem 
      Appearance      =   0  'Flat
      BackColor       =   &H00FFFFFF&
      BorderStyle     =   1  'Fixed Single
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H000000FF&
      Height          =   375
      Left            =   120
      TabIndex        =   1
      Top             =   600
      Width           =   495
   End
   Begin VB.Menu file 
      Caption         =   "&File"
      Begin VB.Menu exit 
         Caption         =   "E&xit"
      End
   End
   Begin VB.Menu help 
      Caption         =   "&Help"
      Begin VB.Menu abt 
         Caption         =   "A&bout"
      End
   End
End
Attribute VB_Name = "calci"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Option Explicit
Dim op1 As Double, op2 As Double, s As Integer, p As Integer, operator As String, eq As Integer

Private Sub abt_Click()
frmAbout.Show
End Sub

Private Sub exit_Click()
End
End Sub

Private Sub Form_Load()
Text1.Text = ""
p = 1
eq = 1
txtpint = "y"
s = 1
End Sub

Private Sub Image2_Click(Index As Integer)
If Index = 0 Then
MsgBox txtpint.Text & eq
If txtpint = "y" Then
Call equal(Index)
Else
Text1.Text = Text1.Text & Index + 1
End If
End If

If Index = 1 Then
If txtpint = "y" Then
Call equal(Index)
Else
Text1.Text = Text1.Text & Index + 1
End If
End If

If Index = 2 Then
If txtpint = "y" Then
Call equal(Index)
Else
Text1.Text = Text1.Text & Index + 1
End If
End If

If Index = 3 Then
If txtpint = "y" Then
Call equal(Index)
Else
Text1.Text = Text1.Text & Index + 1
End If
End If

If Index = 4 Then
If txtpint = "y" Then
Call equal(Index)
Else
Text1.Text = Text1.Text & Index + 1
End If
End If

If Index = 5 Then
If txtpint = "y" Then
Call equal(Index)
Else
Text1.Text = Text1.Text & Index + 1
End If
End If

If Index = 6 Then
If txtpint = "y" Then
Call equal(Index)
Else
Text1.Text = Text1.Text & Index + 1
End If
End If

If Index = 7 Then
If txtpint = "y" Then
Call equal(Index)
Else
Text1.Text = Text1.Text & Index + 1
End If
End If

If Index = 8 Then
If txtpint = "y" Then
Call equal(Index)
Else
Text1.Text = Text1.Text & Index + 1
End If
End If

If Index = 9 Then
If txtpint = "y" And s = 0 Then
Call equal(Index)
txtpint.Text = "n"
Else
If Text1.Text Like ("0.*") Or s = 1 Or Val(Text1.Text) > 0 Then
Text1.Text = Text1.Text & "0"
s = 0
End If
End If
End If

If Index = 10 Then
If txtpint = "y" Then
Call equal(Index)
Else
Text1.Text = Text1.Text & "00"
End If
End If
End Sub

Public Sub proc()
If (Not IsNumeric(Text1.Text)) Then
MsgBox "Invalid value", 16, "Error"
Else
op1 = CDbl(Text1.Text)
Text1.Text = ""
End If
End Sub


Private Sub Image3_Click(sindex As Integer)
Call enable
If sindex = 0 Then
Call proc
operator = "+"
End If

If sindex = 1 Then
Call proc
operator = "-"
End If

If sindex = 2 Then
Call proc
operator = "*"
End If

If sindex = 3 Then
Call proc
operator = "/"
End If

If sindex = 4 Then
Call proc
Text1.Text = Sin(op1 * ((22 / 7) / 180))
txtpint.Text = "y"
eq = 0
MsgBox txtpint.Text
End If

If sindex = 5 Then
Call proc
If op1 = "90" Then
Text1.Text = "0"
Else
Text1.Text = Cos(op1 * ((22 / 7) / 180))
End If
txtpint.Text = "y"
eq = 0
End If

If sindex = 6 Then
Call proc
If op1 = "90" Then
Text1.Text = "Invalid input or function"
Else
Text1.Text = Tan(op1 * ((22 / 7) / 180))
End If
txtpint.Text = "y"
eq = 0
End If

If sindex = 7 Then
Call proc
Text1.Text = Log(op1) / 2.30258509299405
txtpint.Text = "y"
eq = 0
End If

If sindex = 8 Then
Call proc
If (op1 < 0) Then
MsgBox "Imaginary root", vbExclamation, "Error"
Else
Text1.Text = Sqr(op1)
End If
txtpint.Text = "y"
eq = 0
End If

If sindex = 9 Then
Call proc
If (op1 = 0) Then
MsgBox "Division by 0!", 16, "Error"
Else
Text1.Text = 1 / op1
End If
txtpint.Text = "y"
eq = 0
End If

If sindex = 10 Then
If Text1.Text <> "" And Text1.Text <> "0" And eq <> 0 Then
MsgBox "first"
Text1.Text = Text1.Text & "."
Image3(10).Enabled = False
End If
If txtpint = "y" And Text1.Text = "" And p = 1 Then
s = 1
Text1.Text = 0 & "."
Image3(10).Enabled = False
Else
If eq = 1 Then
If Text1.Text = "0" Then
Text1.Text = 0 & "."
Image3(10).Enabled = False
End If
End If
End If
End If

If sindex = 11 Then
s = 0
p = 1
Image3(10).Enabled = True
On Error Resume Next
Dim i, result, a, ans As Double
result = 1
op2 = Text1.Text
If (Not IsNumeric(Text1.Text)) Then
MsgBox "Invalid value", 16, "Error"
Else
Select Case operator
Case "+"
ans = op1 + op2
Text1.Text = ans
Case "-"
ans = op1 - op2
Text1.Text = ans
Case "*"
ans = op1 * op2
Text1.Text = ans
Case "/"
ans = op1 / op2
Text1.Text = ans
Case "^"
For i = 1 To op2 Step 1
result = op1 * result
Next
Text1.Text = result
End Select
eq = 0
txtpint = "y"
MsgBox " eq = " & eq & " s= " & s & "txtpint.text= " & txtpint.Text
End If
End If

If sindex = 12 Then
Image3(13).Enabled = True
Image3(14).Enabled = True
lblmem.Enabled = True
lblmem.Caption = " M"
lblmem.BackColor = &HD2EAF0
If Trim(Val(txtmem.Text)) = "" Then
txtmem.Text = Val(Text1.Text)
Else
txtmem.Text = Val(txtmem.Text) + Val(Text1.Text)
End If
txtpint.Text = "y"
eq = 0
End If

If sindex = 13 Then
Text1.Text = Val(txtmem.Text)
End If

If sindex = 14 Then
lblmem.Enabled = False
Image3(13).Enabled = False
Image3(14).Enabled = False
txtmem.Text = ""
lblmem.Caption = ""
lblmem.BackColor = vbWhite
End If

If sindex = 15 Then
On Error Resume Next
Text1.Text = Left(Text1.Text, Len(Text1.Text) - 1)
If Text1.Text = "" Then
Image3(15).Enabled = False
End If
End If

If sindex = 16 Then
Call proc
operator = "^"
End If

If sindex = 17 Then
Text1.Text = ""
eq = 1
txtpint = "y"
s = 1
End If

If sindex = 18 Then
End
End If

If sindex = 19 Then
If Not (Text1.Text = "") And Not (Val(Text1.Text) = 0) Then
Text1.Text = -Val(Text1.Text)
End If
End If


End Sub

Private Sub enable()
Image3(10).Enabled = True
End Sub

Private Sub equal(andex)
Dim p As Integer

Select Case andex
Case 0
If eq = 0 Then
Call qe(andex)
Else
If s = 1 Or Not (Text1.Text) = "" Then
Text1.Text = Text1.Text & "1"
Else
If s = 0 Then
Text1.Text = "1"
Call pint
End If
End If
End If

Case 1

If eq = 0 Then
Call qe(andex)
Else
If s = 1 Or Not (Text1.Text) = "" Then
Text1.Text = Text1.Text + "2"
Else
If s = 0 Then
Text1.Text = "2"
Call pint
End If
End If
End If

Case 2

If eq = 0 Then
Call qe(andex)
Else
If s = 1 Or Not (Text1.Text) = "" Then
Text1.Text = Text1.Text + "3"
Else
If s = 0 Then
Text1.Text = "3"
Call pint
End If
End If
End If

Case 3

If eq = 0 Then
Call qe(andex)
Else
If s = 1 Or Not (Text1.Text) = "" Then
Text1.Text = Text1.Text + "4"
Else
If s = 0 Then
Text1.Text = "4"
Call pint
End If
End If
End If

Case 4

If eq = 0 Then
Call qe(andex)
Else
If s = 1 Or Not (Text1.Text) = "" Then
Text1.Text = Text1.Text + "5"
Else
If s = 0 Then
Text1.Text = "5"
Call pint
End If
End If
End If

Case 5

If eq = 0 Then
Call qe(andex)
Else
If s = 1 Or Not (Text1.Text) = "" Then
Text1.Text = Text1.Text + "6"
Else
If s = 0 Then
Text1.Text = "6"
Call pint
End If
End If
End If

Case 6

If eq = 0 Then
Call qe(andex)
Else
If s = 1 Or Not (Text1.Text) = "" Then
Text1.Text = Text1.Text + "7"
Else
If s = 0 Then
Text1.Text = "7"
Call pint
End If
End If
End If

Case 7

If eq = 0 Then
Call qe(andex)
Else
If s = 1 Or Not (Text1.Text) = "" Then
Text1.Text = Text1.Text + "8"
Else
If s = 0 Then
Text1.Text = "8"
Call pint
End If
End If
End If

Case 8

If eq = 0 Then
Call qe(andex)
Else
If s = 1 Or Not (Text1.Text) = "" Then
Text1.Text = Text1.Text + "9"
Else
If s = 0 Then
Text1.Text = "9"
Call pint
End If
End If
End If

Case 9
If eq = 0 Then
Call qe(andex)
Else
's=0 And Not (eq = 0)
If Not (Text1.Text) = "0" And Not (Text1.Text) = "" Then
Text1.Text = Text1.Text & "0"
txtpint = "y"
Call pint
Else
If s = 0 Or eq = 0 Then
Text1.Text = "0"
txtpint = "y"
End If
End If
End If


Case 10
If s = 0 And (Not (Text1.Text) = "") And eq = 1 Then
Text1.Text = Text1.Text & "00"
Else
If Trim(Text1.Text) = "" Then
txtpint.Text = "y"
End If
End If
End Select
s = 0
End Sub

Private Sub Image4_Click(dex As Integer)
If dex = 0 Then
calci.WindowState = 1
Else
If dex = 1 Then
End
End If
End If
End Sub

Private Sub pint()
If txtpint = "y" Then
If s = 0 Then
txtpint = "n"
Else
txtpint = "y"
End If
End If
End Sub
Private Sub qe(andex)
If eq = 0 Then
If Not (Text1.Text) = "" And s = 1 Then
Text1.Text = Text1.Text & andex
Else
If s = 0 And eq = 0 And txtpint.Text = "y" And andex = 9 Then
Text1.Text = 0
Else
Text1.Text = andex + 1
eq = 1
txtpint.Text = "y"
s = 0
End If
End If
End If
End Sub

Private Sub onlyz(andex)
If Text1.Text = "0" Then
Text1.Text = andex + 1
End If
Exit Sub
End Sub
