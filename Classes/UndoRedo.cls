VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "UndoRedo"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
Private m_UndoStack As Collection
Private m_RedoStack As Collection

Private Sub Class_Initialize()
    ClearUndo
    ClearRedo
End Sub

Private Sub ClearUndo()
    Set m_UndoStack = New Collection
End Sub
Private Sub ClearRedo()
    Set m_RedoStack = New Collection
End Sub

Public Sub Clear()
    Class_Initialize
End Sub

Public Function CanUndo() As Boolean
    CanUndo = 1 < m_UndoStack.Count
End Function

Public Function CanRedo() As Boolean
    CanRedo = 0 < m_RedoStack.Count
End Function

Private Sub SaveRedo(Obj)
    m_RedoStack.Add Obj
End Sub

Public Sub SaveUndo(Obj)
    m_UndoStack.Add Obj
    If 0 < m_RedoStack.Count Then ClearRedo
End Sub

Public Function Undo() As Object
    If Not Me.CanUndo Then Exit Function
    m_RedoStack.Add m_UndoStack.Item(m_UndoStack.Count)
    m_UndoStack.Remove m_UndoStack.Count
    Set Undo = m_UndoStack.Item(m_UndoStack.Count)
End Function

Public Function Redo() As Object
    If Not Me.CanRedo Then Exit Function
    m_UndoStack.Add m_RedoStack.Item(m_RedoStack.Count)
    Set Redo = m_RedoStack.Item(m_RedoStack.Count)
    m_RedoStack.Remove m_RedoStack.Count
End Function
