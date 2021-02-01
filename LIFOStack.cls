VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "LIFOStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Code by Phlip Bradbury
'phlipping@yahoo.com
'Standard code for a Last In-First Out Stack
'ported to VB

'How it works:
'lCount simply is the number of items in the stack.
'is is incremented after a push and decremented after
'a pop.
'Key:
'C = lCount
' 1 2 3 X X
'     C
'Push 4:
' 1 2 3 4 X
'       C
' Pop:
' 1 2 3 X X
'     C
Option Explicit
Private lStack() As Variant 'Local for Stack
Private lSize As Integer 'Size of stack's allocated mem
Private lCount As Integer 'number of items in stack
Public Enum enmLIFOState 'for LIFO::State
  LIFOEmpty = 0
  LIFOPartial = 1
  LIFOFull = 2
End Enum

'add an item to stack
Public Sub Push(Data As Variant)
  If lCount >= lSize Then 'if stack is full,
    Err.Raise 6 'Overflow
  Else
    'add to end of stack
    lCount = lCount + 1
    lStack(lCount) = Data
  End If
End Sub

'remove last item from stack
Public Function Pop() As Variant
  If lCount <= 0 Then 'if stack is empty,
    Pop = Null 'return Null
  Else
    'get last item
    Pop = lStack(lCount)
    lCount = lCount - 1
  End If
End Function

'change size
Public Property Let Size(NewValue As Integer)
  If NewValue <= 0 Then Err.Raise 5 'Illegal Function Call
  lSize = NewValue
  'empty stack
  lCount = 0
  'allocate new memory
  ReDim lStack(1 To NewValue)
End Property

Public Property Get Size() As Integer
  Size = lSize
End Property

Public Property Get Count() As Integer
  Count = lCount
End Property

'find state of stack
Public Property Get State() As enmLIFOState
  If lCount <= 0 Then
    State = LIFOEmpty
  ElseIf lCount >= lSize Then
    State = LIFOFull
  Else
    State = LIFOPartial
  End If
End Property

Private Sub Class_Initialize()
  'default size=100 elements
  Size = 100
End Sub

Public Property Get Item(Index As Integer) As Variant
  'check index is in stack
  If Index <= 0 Or Index > lCount Then Err.Raise 9 'Subscript out of range
  Item = lStack(Index)
End Property
