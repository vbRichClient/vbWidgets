VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "cUndoRedo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit 'A generic Undo/Redo History-List...
 
Private States As cCollection, Idx As Long

Private Sub Class_Initialize()
  Set States = New_c.Collection
End Sub

Public Sub Clear()
  Set States = New_c.Collection
  Idx = 0
End Sub

Public Sub SaveState(State)
Dim i As Long
  'delete all states up to the last UndoIdx, before adding the new state on Top
  For i = 1 To Idx - 1: States.Remove 1: Next i

  States.Add State, , 1 'add state on top of the States-Collection
  Idx = 1 'last added state is alwys on top
End Sub

Public Property Get PreviousState()
  If States.Count = 0 Or Idx >= States.Count Then Exit Property Else Idx = Idx + 1
  Assign PreviousState, States(Idx)
End Property

Public Property Get NextState()
  If States.Count = 0 Or Idx <= 1 Then Exit Property Else Idx = Idx - 1
  Assign NextState, States(Idx)
End Property

Private Sub Assign(LHS, RHS)
  If IsObject(RHS) Then Set LHS = RHS Else LHS = RHS
End Sub

Public Property Get UndoEnabled() As Boolean
  UndoEnabled = (States.Count > 1 And Idx < States.Count)
End Property

Public Property Get RedoEnabled() As Boolean
  RedoEnabled = (Idx > 1)
End Property

