VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "FIFOStack"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'Code by Phlip Bradbury
'phlipping@yahoo.com
'Standard code for a First In-First Out Stack
'ported to VB

'How it works:
'The lTop and lBottom vars point to the indexes
'to next Pop and Push from, respectively.
'They are incremented after a pop or a push, they
'are incremented, and wrap around the array to keep
'the whole array in use.
'Key:
'T = lTop
'B = lBottom
' 1 2 3 X X
' T     B
'Push 4:
' 1 2 3 4 X
' T       B
'Pop:
' X 2 3 4 X
'   T     B
'Pop:
' X X 3 4 X
'     T   B
'Push 5 and 6:
' 6 X 3 4 5
'   B T
'Push 7 (full):
' 6 7 3 4 5
'     B
'     T
'Pop 5 times (empty):
' X X X X X
'     B
'     T
Option Explicit
Private lStack() As Variant 'Local for Stack
Private lSize As Integer 'Size of stack's allocated mem
Private lTop As Integer 'Item next to be popped
Private lBottom As Integer 'Item next to be pushed
Public Enum enmFIFOState
  FIFOEmpty = 0
  FIFOPartial = 1
  FIFOFull = 2
End Enum
Private lState As enmFIFOState 'Full/Empty?

'add an item to stack
Public Sub Push(Data As Variant)
  If lState = FIFOFull Then
    Err.Raise 6 'Overflow
  Else
    'add to bottom
    lStack(lBottom) = Data
    'increment bottom
    lBottom = lBottom + 1
    'wrap
    If lBottom > lSize Then lBottom = 1
    'if bottom=top then full
    If lBottom = lTop Then lState = FIFOFull Else lState = FIFOPartial
  End If
End Sub

'remove first item from stack
Public Function Pop() As Variant
  If lState = FIFOEmpty Then
    Pop = Null 'Null means stack is empty
  Else
    'pop from top
    Pop = lStack(lTop)
    'increment top
    lTop = lTop + 1
    'wrap
    If lTop > lSize Then lTop = 1
    'if bottom=top then empty
    If lBottom = lTop Then lState = FIFOEmpty Else lState = FIFOPartial
  End If
End Function

'change size
Public Property Let Size(NewValue As Integer)
  If NewValue <= 0 Then Err.Raise 5 'Illegal Function Call
  lSize = NewValue
  'empty stack
  lTop = 1
  lBottom = 1
  lState = FIFOEmpty
  'allocate new memory
  ReDim lStack(1 To NewValue)
End Property

Public Property Get Size() As Integer
  Size = lSize
End Property

Public Property Get Count() As Integer
  Select Case lState
    'can we do it the quick way?
    Case FIFOEmpty
      Count = 0
    Case FIFOFull
      Count = lSize
    Case FIFOPartial
      'actually calculate the number of elements
      If lTop < lBottom Then
        'data doesn't wrap
        Count = lBottom - lTop
      Else
        'data wraps
        Count = lBottom - lTop + lSize
      End If
  End Select
End Property

Public Property Get State() As enmFIFOState
  State = lState
End Property

Private Sub Class_Initialize()
  'default size=100 elements
  Size = 100
End Sub

Public Property Get Item(Index As Integer) As Variant
  Dim A As Integer
  'check index is in stack
  If Index <= 0 Or Index > Count Then Err.Raise 9 'Subscript out of range
  'find item number
  A = Index + lTop - 1
  'wrap
  If A > lSize Then A = A - lSize
  Item = lStack(A)
End Property
