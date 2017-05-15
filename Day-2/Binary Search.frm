VERSION 5.00
Begin VB.Form Form1 
   Caption         =   "Form1"
   ClientHeight    =   3030
   ClientLeft      =   120
   ClientTop       =   450
   ClientWidth     =   4560
   LinkTopic       =   "Form1"
   ScaleHeight     =   3030
   ScaleWidth      =   4560
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton Command1 
      Caption         =   "Click Here."
      Height          =   855
      Left            =   840
      TabIndex        =   0
      Top             =   600
      Width           =   2535
   End
End
Attribute VB_Name = "Form1"
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
Dim iStart As Integer
Dim iEnd As Integer
Dim iMid As Integer
iStart = 1
iEnd = n
While iStart <= iEnd
    If A(iMid) = key Then
        MsgBox ("Found at: " + CStr(i))
        Exit Sub
    ElseIf A(iMid) < key Then
        iEnd = iMid - 1
    Else
        iStart = iMid + 1
    End If
MsgBox ("Key not found!")
End Sub
