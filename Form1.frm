VERSION 5.00
Begin VB.Form Form1 
   BorderStyle     =   3  'Fixed Dialog
   Caption         =   "Stack Sample"
   ClientHeight    =   2970
   ClientLeft      =   45
   ClientTop       =   330
   ClientWidth     =   3615
   LinkTopic       =   "Form1"
   MaxButton       =   0   'False
   MinButton       =   0   'False
   ScaleHeight     =   2970
   ScaleWidth      =   3615
   StartUpPosition =   3  'Windows Default
   Begin VB.CommandButton ShowDat 
      Caption         =   "Show Stacks >>"
      Height          =   330
      Left            =   1785
      TabIndex        =   10
      Top             =   105
      Width           =   1485
   End
   Begin VB.ListBox List2 
      Height          =   2400
      Left            =   5070
      TabIndex        =   9
      Top             =   420
      Width           =   1275
   End
   Begin VB.ListBox List1 
      Height          =   2400
      Left            =   3705
      TabIndex        =   7
      Top             =   420
      Width           =   1275
   End
   Begin VB.CommandButton LIFOOut 
      Caption         =   "Pop From LIFO"
      Height          =   645
      Left            =   2100
      TabIndex        =   4
      Top             =   1890
      Width           =   1275
   End
   Begin VB.CommandButton LIFOIn 
      Caption         =   "Push to LIFO"
      Height          =   645
      Left            =   2100
      TabIndex        =   3
      Top             =   945
      Width           =   1275
   End
   Begin VB.CommandButton FIFOOut 
      Caption         =   "Pop From FIFO"
      Height          =   645
      Left            =   315
      TabIndex        =   2
      Top             =   1890
      Width           =   1275
   End
   Begin VB.CommandButton FIFOIn 
      Caption         =   "Push to FIFO"
      Height          =   645
      Left            =   315
      TabIndex        =   1
      Top             =   945
      Width           =   1275
   End
   Begin VB.TextBox Text1 
      Height          =   285
      Left            =   420
      TabIndex        =   0
      Text            =   "Data to Push"
      Top             =   525
      Width           =   2955
   End
   Begin VB.Label Label3 
      Alignment       =   2  'Center
      Caption         =   "FIFO                        LIFO"
      Height          =   225
      Left            =   3915
      TabIndex        =   8
      Top             =   105
      Width           =   2220
   End
   Begin VB.Label Label2 
      Caption         =   "Label1"
      Height          =   225
      Left            =   2100
      TabIndex        =   6
      Top             =   2625
      Width           =   1275
   End
   Begin VB.Label Label1 
      Caption         =   "Label1"
      Height          =   225
      Left            =   315
      TabIndex        =   5
      Top             =   2625
      Width           =   1275
   End
End
Attribute VB_Name = "Form1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit
Dim LIFO As LIFOStack
Dim FIFO As FIFOStack

Private Sub DoLists()
  Dim I As Integer
  List1.Clear
  For I = 1 To FIFO.Count
    List1.AddItem FIFO.Item(I)
  Next
  List2.Clear
  For I = 1 To LIFO.Count
    List2.AddItem LIFO.Item(I)
  Next
End Sub

Private Sub Form_Load()
  Dim A As String
  A = 100
  Do
    If Not IsNumeric(A) Then MsgBox "Please enter a number"
    A = InputBox("Size of stacks?", , A)
    If A = "" Then Unload Me: Exit Sub
  Loop Until IsNumeric(A)
  Set LIFO = New LIFOStack
  Set FIFO = New FIFOStack
  LIFO.Size = CInt(A)
  FIFO.Size = CInt(A)
  Label1.Caption = FIFO.Count & " of " & FIFO.Size
  Label2.Caption = LIFO.Count & " of " & LIFO.Size
End Sub

Private Sub Form_Unload(Cancel As Integer)
  Set LIFO = Nothing
  Set FIFO = Nothing
End Sub

Private Sub LIFOIn_Click()
  On Error GoTo ErrHandler
  LIFO.Push Text1.Text
  On Error GoTo 0
  Label2.Caption = LIFO.Count & " of " & LIFO.Size
  If Width = 6255 Then DoLists
  Exit Sub
ErrHandler:
  If Err.Number = 6 Then
    MsgBox "Stack Full"
  Else
    Err.Raise Err.Number
  End If
End Sub

Private Sub LIFOOut_Click()
  Dim A As Variant
  A = LIFO.Pop
  If VarType(A) = vbNull Then MsgBox "Stack Empty" Else MsgBox A
  Label2.Caption = LIFO.Count & " of " & LIFO.Size
  If Width = 6255 Then DoLists
End Sub

Private Sub FIFOIn_Click()
  On Error GoTo ErrHandler
  FIFO.Push Text1.Text
  On Error GoTo 0
  Label1.Caption = FIFO.Count & " of " & FIFO.Size
  If Width = 6255 Then DoLists
  Exit Sub
ErrHandler:
  If Err.Number = 6 Then
    MsgBox "Stack Full"
  Else
    Err.Raise Err.Number
  End If
End Sub

Private Sub FIFOOut_Click()
  Dim A As Variant
  A = FIFO.Pop
  If VarType(A) = vbNull Then MsgBox "Stack Empty" Else MsgBox A
  Label1.Caption = FIFO.Count & " of " & FIFO.Size
  If Width = 6255 Then DoLists
End Sub

Private Sub ShowDat_Click()
  Width = 9780 - Width
  If Width = 6255 Then
    ShowDat.Caption = "<< Show Stacks"
    If Left + Width >= Screen.Width Then Left = Screen.Width - Width
    DoLists
  Else
    ShowDat.Caption = "Show Stacks >>"
  End If
End Sub
