VERSION 5.00
Begin VB.Form frmMain 
   BorderStyle     =   1  'Fixed Single
   Caption         =   "Simple Drag Reordering"
   ClientHeight    =   3045
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   4575
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   ScaleHeight     =   3045
   ScaleWidth      =   4575
   StartUpPosition =   2  'CenterScreen
   Begin VB.ListBox lst 
      Height          =   2790
      Left            =   120
      TabIndex        =   0
      Top             =   120
      Width           =   4335
   End
End
Attribute VB_Name = "frmMain"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Private StartingText As String, StartingIndex As Integer 'Declare some variables
Private Sub Form_Load()
    Dim i As Integer
    
    For i = 1 To 10
        lst.AddItem "Example" & i 'create some examples for you to test out the code with
    Next i
End Sub

Private Sub lst_MouseDown(Button As Integer, Shift As Integer, X As Single, Y As Single)
    StartingText = lst.Text 'store orginal data
    StartingIndex = lst.ListIndex 'store original data
End Sub

Private Sub lst_MouseUp(Button As Integer, Shift As Integer, X As Single, Y As Single)
    Dim SecondIndex As Integer 'declare a variable to store the list index of the finishing point
    
    If StartingIndex <> lst.ListIndex Then  'if they moved the listitem
        SecondIndex = lst.ListIndex 'store the destination listindex
        lst.RemoveItem StartingIndex 'remove the original row so everything below shifts up
        lst.AddItem StartingText, SecondIndex 'add the original row back into the new place
        lst.ListIndex = SecondIndex 'making it look better by having the selection be on the destination row
    End If
    Dragging = False
End Sub
