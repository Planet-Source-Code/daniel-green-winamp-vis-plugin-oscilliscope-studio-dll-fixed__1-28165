VERSION 5.00
Begin VB.Form frmFunctions 
   BorderStyle     =   4  'Fixed ToolWindow
   Caption         =   "f u n c t i o n  -  h e l p"
   ClientHeight    =   3375
   ClientLeft      =   45
   ClientTop       =   285
   ClientWidth     =   7575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   3375
   ScaleWidth      =   7575
   ShowInTaskbar   =   0   'False
   StartUpPosition =   2  'CenterScreen
   Begin VB.Label lblCOTAN 
      Caption         =   "cotan(var) : returns the cotangent, or inverse tangent, of var."
      Height          =   255
      Left            =   120
      TabIndex        =   12
      Top             =   3000
      Width           =   7335
   End
   Begin VB.Label lblATAN 
      Caption         =   "atan(var) : returns the arctangent of var."
      Height          =   255
      Left            =   120
      TabIndex        =   11
      Top             =   2760
      Width           =   7335
   End
   Begin VB.Label lblSQRT 
      Caption         =   "sqrt(var) : returns the square-root of var."
      Height          =   255
      Left            =   120
      TabIndex        =   10
      Top             =   2520
      Width           =   7335
   End
   Begin VB.Label lblSQR 
      Caption         =   "sqr(var) : rreturns the square of var."
      Height          =   255
      Left            =   120
      TabIndex        =   9
      Top             =   2280
      Width           =   7335
   End
   Begin VB.Label lblLOG10 
      Caption         =   "log10(var) : returns a double specifying the natural base-10 logarithm of a number, var."
      Height          =   255
      Left            =   120
      TabIndex        =   8
      Top             =   2040
      Width           =   7335
   End
   Begin VB.Label lblLOG 
      Caption         =   "log(var) : returns a double specifying the natural logarithm of a number, var."
      Height          =   255
      Left            =   120
      TabIndex        =   7
      Top             =   1800
      Width           =   7335
   End
   Begin VB.Label lblRADDEG 
      Caption         =   "raddeg(var) : returns var (radians) as degrees."
      Height          =   255
      Left            =   120
      TabIndex        =   6
      Top             =   1560
      Width           =   7335
   End
   Begin VB.Label lblDEGRAD 
      Caption         =   "degrad(var) : returns var (degrees) as radians."
      Height          =   255
      Left            =   120
      TabIndex        =   5
      Top             =   1320
      Width           =   7335
   End
   Begin VB.Label lblTAN 
      Caption         =   "tan(var) : returns the tangent value of var, where var equals the angle in degrees."
      Height          =   255
      Left            =   120
      TabIndex        =   4
      Top             =   1080
      Width           =   7335
   End
   Begin VB.Label lblRAND 
      Caption         =   "rand(var) : returns a random value between zero and var."
      Height          =   255
      Left            =   120
      TabIndex        =   3
      Top             =   840
      Width           =   7335
   End
   Begin VB.Label lblABS 
      Caption         =   "abs(var) : returns the absolute value of var."
      Height          =   255
      Left            =   120
      TabIndex        =   2
      Top             =   600
      Width           =   7335
   End
   Begin VB.Label lblCOS 
      Caption         =   "cos(var) : returns the cosine value of var, where var equals the angle in degrees."
      Height          =   255
      Left            =   120
      TabIndex        =   1
      Top             =   360
      Width           =   7335
   End
   Begin VB.Label lblSIN 
      Caption         =   "sin(var) : returns the sine value of var, where var equals the angle in degrees."
      Height          =   255
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   7335
   End
End
Attribute VB_Name = "frmFunctions"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub lblPI_Click()

End Sub

Private Sub Form_Unload(Cancel As Integer)
Unload Me
End Sub
