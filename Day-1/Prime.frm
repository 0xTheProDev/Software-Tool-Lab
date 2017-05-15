VERSION 5.00
Begin VB.Form Form2 
   Caption         =   "Form2"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form2"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Here"
      Height          =   615
      Left            =   2640
      TabIndex        =   2
      Top             =   360
      Width           =   1095
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
      Height          =   855
      Left            =   240
      TabIndex        =   0
      Top             =   240
      Width           =   1815
   End
   Begin VB.Label Label1 
      Caption         =   "Click the Button."
      Height          =   495
      Left            =   960
      TabIndex        =   1
      Top             =   1560
      Width           =   2175
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim n As Integer
n = CInt(Text1.Text)
Dim m As Integer
m = n / 2
For i = 2 To m
    If n Mod i = 0 Then
        Label1 = "Composite"
        Exit Sub
    End If
Next
Label1 = "Prime"
End Sub
