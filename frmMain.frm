VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Quadratic Equation Calculator"
   ClientHeight    =   3060
   ClientLeft      =   2190
   ClientTop       =   3390
   ClientWidth     =   6615
   Icon            =   "frmMain.frx":0000
   KeyPreview      =   -1  'True
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3060
   ScaleWidth      =   6615
   StartUpPosition =   2  'CenterScreen
   Begin VB.CommandButton cmdClear 
      Caption         =   "&Clear Contents"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   11
      ToolTipText     =   "Click here to clear the values and the roots"
      Top             =   600
      Width           =   3495
   End
   Begin VB.TextBox txtValue 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   2
      Left            =   1140
      TabIndex        =   2
      ToolTipText     =   "Enter the numeric value for C"
      Top             =   1200
      Width           =   1455
   End
   Begin VB.TextBox txtValue 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   1
      Left            =   1140
      TabIndex        =   1
      ToolTipText     =   "Enter the numeric value for B"
      Top             =   840
      Width           =   1455
   End
   Begin VB.CommandButton cmdCompute 
      Caption         =   "Display the &Root(s)"
      Default         =   -1  'True
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   375
      Left            =   2880
      TabIndex        =   3
      ToolTipText     =   "Click here to display the root(s)"
      Top             =   1080
      Width           =   3495
   End
   Begin VB.TextBox txtValue 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   285
      Index           =   0
      Left            =   1140
      TabIndex        =   0
      ToolTipText     =   "Enter the numeric value for A"
      Top             =   480
      Width           =   1455
   End
   Begin VB.Frame fraOutput 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   975
      Left            =   120
      TabIndex        =   7
      Top             =   1680
      Width           =   6375
      Begin VB.Label lblRoot2 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   9
         ToolTipText     =   "Value of 2nd root"
         Top             =   600
         Width           =   5895
      End
      Begin VB.Label lblRoot1 
         BeginProperty Font 
            Name            =   "Tahoma"
            Size            =   9.75
            Charset         =   0
            Weight          =   700
            Underline       =   0   'False
            Italic          =   0   'False
            Strikethrough   =   0   'False
         EndProperty
         Height          =   255
         Left            =   240
         TabIndex        =   8
         ToolTipText     =   "Value of 1st root"
         Top             =   240
         Width           =   5895
      End
   End
   Begin VB.Label Label5 
      AutoSize        =   -1  'True
      Caption         =   "Enter the values of A, B, C from the equation Ax² + Bx + C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   195
      Left            =   120
      TabIndex        =   12
      Top             =   120
      Width           =   4290
   End
   Begin VB.Label Label4 
      Alignment       =   1  'Right Justify
      AutoSize        =   -1  'True
      Caption         =   "©2001 by Marc Christian Saribay"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      ForeColor       =   &H8000000D&
      Height          =   195
      Left            =   4080
      TabIndex        =   10
      ToolTipText     =   "QECalc™ (August 5, 2001)"
      Top             =   2760
      Width           =   2385
   End
   Begin VB.Label Label3 
      Caption         =   "Value of C"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   6
      Top             =   1245
      Width           =   855
   End
   Begin VB.Label Label2 
      Caption         =   "Value of B"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   5
      Top             =   885
      Width           =   855
   End
   Begin VB.Label Label1 
      Caption         =   "Value of A"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   8.25
         Charset         =   0
         Weight          =   400
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   255
      Left            =   240
      TabIndex        =   4
      Top             =   525
      Width           =   855
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

Public Sub ClearContents()
  'Clears the value boxes and root display
  For Index = 0 To 2
    txtValue(Index).Text = ""
  Next
  fraOutput.Caption = ""
  lblRoot1.Caption = ""
  lblRoot2.Caption = ""
  txtValue(0).SetFocus
End Sub

Private Sub cmdClear_Click()
  ClearContents
End Sub

Private Sub cmdCompute_Click()
  'Message Box contents
  strPrompt1 = "Enter the values for A, B, C"
  strPrompt2 = "Invalid Equation!"
  strPrompt3 = "Please enter numeric values!"
  strTitle = "QE Calculator"
  'Analysis of equation
  If txtValue(0).Text = "" Or txtValue(1).Text = "" Or txtValue(2).Text = "" Then
    'All value boxes must be filled
    Msg = MsgBox(strPrompt1, vbOKOnly + vbInformation, strTitle)
    ClearContents
  Else
    On Error GoTo ErrorHandler
    'Assign values to variables
    intA = Int(txtValue(0).Text)
    intB = Int(txtValue(1).Text)
    intC = Int(txtValue(2).Text)
    'Check if equation is Linear or Quadratic
    If txtValue(0).Text = 0 Then
      If txtValue(1).Text = 0 Then
        'Equation is invalid
        Msg = MsgBox(strPrompt2, vbOKOnly + vbExclamation, strTitle)
        ClearContents
      Else
        'Equation is Linear
        intX1 = -intC / intB
        fraOutput.Caption = "Linear Equation"
        lblRoot1.Caption = "X = " & Str(intX1)
        lblRoot2.Caption = ""
      End If
    Else
      'Equation is Quadratic
      intDisc = (intB ^ 2) - (4 * intA * intC) 'Discriminant
      If intDisc = 0 Then
        '1 real root
        intX1 = -intB / (2 * intA)
        fraOutput.Caption = "Quadratic Equation (1 real root)"
        lblRoot1.Caption = "X = " & Str(intX1)
        lblRoot2.Caption = ""
      ElseIf intDisc > 0 Then
        '2 real roots
        intX1 = (-intB + Sqr(intDisc)) / (2 * intA)
        intX2 = (-intB - Sqr(intDisc)) / (2 * intA)
        fraOutput.Caption = "Quadratic Equation (2 real roots)"
        lblRoot1.Caption = "X = " & Str(intX1)
        lblRoot2.Caption = "X = " & Str(intX2)
      Else
        'If Disc < 0, 2 imaginary roots
        intX1 = -intB / (2 * intA)
        intX2 = Sqr(Abs(intDisc)) / (2 * intA)
        fraOutput = "Quadratic Equation (2 imaginary roots)"
        'Imaginary part separated
        lblRoot1.Caption = "X = " & Str(intX1) & " + " & Str(intX2) & " i"
        lblRoot2.Caption = "X = " & Str(intX1) & " - " & Str(intX2) & " i"
      End If
    End If
  End If
Exit Sub
ErrorHandler:
  'Only numeric values will be validated
  If Err.Number > 0 Then
    Msg = MsgBox(strPrompt3, vbOKOnly + vbExclamation, strTitle)
    Err.Clear
    ClearContents
  End If
End Sub

Private Sub Form_KeyDown(KeyCode As Integer, Shift As Integer)
  If KeyCode = vbKeyEscape Then
    'Escape key function to quit the program
    Unload Me
  End If
End Sub

Private Sub txtValue_GotFocus(Index As Integer)
  'Select whole value when the value box is accessed
  txtValue(Index).SelStart = 0
  txtValue(Index).SelLength = Len(txtValue(Index).Text)
End Sub
