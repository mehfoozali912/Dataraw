Attribute VB_Name = "SendAllDrafts"
Option Explicit

' Sends every draft in your Drafts folder whose subject starts with
' "Nigeria Energy 2026" - and ONLY those. Your other personal drafts
' are left untouched. Run it from Excel with classic Outlook open.

Sub SendAllNigeriaDrafts()
    Dim ol As Object, ns As Object, df As Object
    Dim i As Long, n As Long, total As Long

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then MsgBox "Could not reach Outlook. Open classic Outlook first.": Exit Sub

    Set ns = ol.GetNamespace("MAPI")
    Set df = ns.GetDefaultFolder(16)   ' Drafts

    total = 0
    For i = 1 To df.Items.Count
        If InStr(1, df.Items(i).Subject & "", "Nigeria Energy 2026", vbTextCompare) = 1 Then total = total + 1
    Next i
    If total = 0 Then MsgBox "No Nigeria Energy 2026 drafts found in Drafts.": Exit Sub

    If MsgBox("Found " & total & " Nigeria Energy 2026 draft(s)." & vbCrLf & _
              "SEND them all now?", vbYesNo + vbExclamation, "Send all drafts?") <> vbYes Then Exit Sub

    ' Loop backwards: sending removes items from the folder as we go
    For i = df.Items.Count To 1 Step -1
        If InStr(1, df.Items(i).Subject & "", "Nigeria Energy 2026", vbTextCompare) = 1 Then
            df.Items(i).Send
            n = n + 1
            Application.Wait Now + TimeValue("0:00:02")
        End If
    Next i

    MsgBox n & " emails sent. Check Sent Items in Outlook.", vbInformation
End Sub
