VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "CGanttOutlineView"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit

Private WithEvents mCommandBars As Office.CommandBars
Private mUpdating As Boolean

Public Sub StartListening()
    Set mCommandBars = Application.CommandBars
End Sub

Public Sub StopListening()
    Set mCommandBars = Nothing
End Sub

Private Sub mCommandBars_OnUpdate()
    If mUpdating Then Exit Sub
    mUpdating = True
    On Error GoTo Finished
    ' Native outline buttons update Excel controls without a cell-change event.
    TaskOutlineViewTick
Finished:
    mUpdating = False
End Sub
