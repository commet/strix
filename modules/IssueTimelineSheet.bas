Attribute VB_Name = "IssueTimelineSheet"
' Issue Timeline Sheet Event Handlers
Option Explicit

' Worksheet_Change event for Issue Timeline sheet
' This code should be placed in the Issue Timeline sheet's code module
Private Sub Worksheet_Change(ByVal Target As Range)
    On Error Resume Next
    
    ' Check if change is in category filter (C5)
    If Not Intersect(Target, Range("C5")) Is Nothing Then
        Call FilterIssuesByCategory
    End If
    
    ' Check if change is in status filter (D5)
    If Not Intersect(Target, Range("D5")) Is Nothing Then
        Call FilterIssuesByStatus
    End If
    
    On Error GoTo 0
End Sub

' Instructions for setup:
' 1. In Excel VBA Editor, find the "Issue Timeline" sheet in the project explorer
' 2. Double-click on it to open its code module
' 3. Copy and paste the Worksheet_Change event above into that module
' 4. The dropdowns will now automatically trigger filtering when changed