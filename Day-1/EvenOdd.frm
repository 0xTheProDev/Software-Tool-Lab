VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "\"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Here"
      Height          =   615
      Left            =   2160
      TabIndex        =   1
      Top             =   480
      Width           =   1335
   End
   Begin VB.TextBox Text1 
      BeginProperty DataFormat 
         Type            =   1
         Format          =   "0;(0)"
         HaveTrueFalseNull=   0
         FirstDayOfWeek  =   0
         FirstWeekOfYear =   0
         LCID            =   16393
         SubFormatType   =   1
      EndProperty
      Height          =   975
      Left            =   240
      TabIndex        =   0
      Top             =   360
      Width           =   1335
   End
   Begin VB.Label Label1 
      Caption         =   "Click the Button"
      Height          =   495
      Left            =   960
      TabIndex        =   2
      Top             =   1800
      Width           =   2055
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
n = CInt(Text1.Text)
If n Mod 2 = 0 Then
    Label1 = "Even"
Else
    Label1 = "Odd"
End If
End Sub
