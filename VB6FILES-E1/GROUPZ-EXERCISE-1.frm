VERSION 5.00
Begin VB.Form frmGroupZ 
   BackColor       =   &H80000015&
   Caption         =   "Group Z- Exercise1"
   ClientHeight    =   2865
   ClientLeft      =   120
   ClientTop       =   615
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   2865
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton cmdClear 
      Caption         =   "Clear"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   735
      Left            =   120
      TabIndex        =   9
      ToolTipText     =   "Click to clear"
      Top             =   5880
      Width           =   6135
   End
   Begin VB.CommandButton cmdCalculate 
      BackColor       =   &H80000018&
      Caption         =   "CALCULATE"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   14.25
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   120
      TabIndex        =   7
      ToolTipText     =   "Click to calculate"
      Top             =   4440
      Width           =   2340
   End
   Begin VB.TextBox txtTime 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   6
      Text            =   "Enter Time"
      Top             =   3240
      Width           =   2535
   End
   Begin VB.TextBox txtRate 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   1033
         SubFormatType   =   1
      EndProperty
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   2
      Text            =   "Enter Interest (%)"
      Top             =   2280
      Width           =   2535
   End
   Begin VB.TextBox txtPrincipal 
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   9.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   3840
      TabIndex        =   1
      Text            =   "Enter Principal"
      Top             =   1200
      Width           =   2535
   End
   Begin VB.Label lblOutput 
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   1095
      Left            =   3360
      TabIndex        =   8
      Top             =   4440
      Width           =   3975
   End
   Begin VB.Label lblTime 
      Caption         =   "Enter Time in years"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   5
      Top             =   3240
      Width           =   3015
   End
   Begin VB.Label lblRate 
      Caption         =   "Enter Interest Rate (%):"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   615
      Left            =   240
      TabIndex        =   4
      Top             =   2280
      Width           =   3015
   End
   Begin VB.Label lblPrincipal 
      Caption         =   "Enter Principal Amount:"
      BeginProperty Font 
         Name            =   "Arial"
         Size            =   12
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   600
      Left            =   240
      TabIndex        =   3
      Top             =   1200
      Width           =   3000
   End
   Begin VB.Label Label1 
      Alignment       =   2  'Center
      Caption         =   "Interest Rate Calculator"
      BeginProperty Font 
         Name            =   "Tahoma"
         Size            =   15.75
         Charset         =   0
         Weight          =   700
         Underline       =   0   'False
         Italic          =   0   'False
         Strikethrough   =   0   'False
      EndProperty
      Height          =   495
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   7095
   End
End
Attribute VB_Name = "frmGroupZ"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub cmdCommand_Click()

End Sub

Private Sub cmdCalculate_Click()
    ' Declare variables for the inputs
    Dim principal As Double
    Dim rate As Double
    Dim time As Double
    Dim amount As Double
    Dim r As Double ' This will be rate/100 for calculation

    ' Input validation: Check if the inputs are numeric
    If IsNumeric(txtPrincipal.Text) And IsNumeric(txtRate.Text) And IsNumeric(txtTime.Text) Then
        ' Get user inputs
        principal = CDbl(txtPrincipal.Text) ' Convert input text to a number
        rate = CDbl(txtRate.Text) ' Convert rate to number
        time = CDbl(txtTime.Text) ' Convert time to number

        ' Check if time period is at least 3 years
        If time < 3 Then
            MsgBox "Please enter a time period of at least 3 years."
            Exit Sub
            ElseIf time > 100 Then
            MsgBox "Please enter a time period of no more than 100 years."
            Exit Sub
        End If

        ' Calculate the interest rate in decimal
        r = rate / 100

        ' Apply the formula A = P(1 + r)^t, with n = 1 (compounded annually)
        amount = principal * (1 + r) ^ time

        ' Display the result in the Output Label
        lblOutput.Caption = "The final amount is: " & Format(amount, "Currency")
        MsgBox "Try Another Value " & Me.Caption ' Show message with form caption
    Else
        ' Display error message if inputs are invalid
        MsgBox "Please enter valid numbers for Principal, Rate, and Time."
    End If
End Sub

Private Sub cmdClear_Click()

    ' Clear all the input textboxes
    txtPrincipal.Text = ""
    txtRate.Text = ""
    txtTime.Text = ""
    
    ' Optionally clear the output label
    lblOutput.Caption = ""
    
    ' Return focus to the first input field
    txtPrincipal.SetFocus
End Sub


Private Sub Form_Load()

    ' Set default values for Textboxes
    txtPrincipal.Text = "0"     ' Default principal amount to 0
'    txtRate.Text = "0"          ' Default interest rate to 0
    txtTime.Text = "3"          ' Default time to 3 year

    ' Set a default message in the Output Label (optional)
    lblOutput.Caption = "Please enter values and click Calculate."
    
    ' Optionally set focus to the first Textbox
'    txtPrincipal.SetFocus
End Sub






Private Sub lblInterest_Click()

End Sub


