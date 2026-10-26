Attribute VB_Name = "WhereAreMyDrafts"
Option Explicit

' Finds the exact Drafts folder the robot saved into,
' tells you its name and how many emails are inside,
' and opens the first one on screen so you can see it.

Sub WhereAreMyDrafts()
    Dim ol As Object, ns As Object, df As Object
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then MsgBox "Could not reach Outlook.": Exit Sub

    Set ns = ol.GetNamespace("MAPI")
    Set df = ns.GetDefaultFolder(16)   ' 16 = the Drafts folder the robot uses

    MsgBox "The robot saves drafts into:" & vbCrLf & vbCrLf & _
           "   Mailbox:  " & df.Parent.Name & vbCrLf & _
           "   Folder:   " & df.FolderPath & vbCrLf & vbCrLf & _
           "It currently holds " & df.Items.Count & " item(s).", _
           vbInformation, "Where are my drafts?"

    If df.Items.Count > 0 Then df.Items(1).Display   ' opens the first draft so you can see it
End Sub
