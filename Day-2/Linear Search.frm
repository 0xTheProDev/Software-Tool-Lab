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
      Height          =   855
      Left            =   960
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private Sub Command1_Click()
Dim n As Integer
n = CInt(InputBox("Enter size of Array: "))
If n > 20 Then
    MsgBox ("Size of Array must be less than or equal 20!")
    Exit Sub
End If
Dim A(20) As Integer
Dim key As Integer
For i = 1 To n
    A(i) = CInt(InputBox("Enter a number: "))
Next
key = CInt(InputBox("Enter the key: "))
For i = 1 To n
    If A(i) = key Then
        MsgBox ("Found at: " + CStr(i))
        Exit Sub
    End If
Next
MsgBox ("Key not found!")
End Sub
