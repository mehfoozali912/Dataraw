Attribute VB_Name = "SendMatchmakingEmails"
Option Explicit

' =====================================================================
'  MATCHMAKING BULK EMAIL SENDER
'  Works on any tab that has header columns: To... | Subject | ...Body
'  (both the Egypt and the Nigeria email workbooks).
'
'  HOW TO USE
'  1. Make sure desktop Outlook is open and signed in.
'  2. Click on the tab you want to send (e.g. "1-Exhibitor Emails").
'  3. Press Alt+F8, choose CreateDrafts, click Run.
'     -> Check your Outlook Drafts folder, spot-check a few emails.
'  4. When happy: Alt+F8, choose SendForReal, click Run.
'  Rows whose To address still contains "[" (unfilled placeholder)
'  are skipped automatically.
' =====================================================================

Sub CreateDrafts()
    ProcessEmails False
End Sub

Sub SendForReal()
    If MsgBox("This will SEND one email per row of the active tab." & vbCrLf & _
              "Have you checked the drafts and filled all placeholders?", _
              vbYesNo + vbExclamation, "Send for real?") = vbYes Then
        ProcessEmails True
    End If
End Sub

Private Sub ProcessEmails(ByVal doSend As Boolean)
    Dim ws As Worksheet, ol As Object, mail As Object
    Dim r As Long, c As Long, n As Long, skipped As Long
    Dim cTo As Long, cSub As Long, cBody As Long

    Set ws = ActiveSheet
    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then MsgBox "Could not reach Outlook. Open Outlook first.": Exit Sub

    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Select Case True
            Case ws.Cells(1, c).Value Like "To*":      cTo = c
            Case ws.Cells(1, c).Value Like "Subject*": cSub = c
            Case ws.Cells(1, c).Value Like "*Body*":   cBody = c
        End Select
    Next c
    If cTo = 0 Or cSub = 0 Or cBody = 0 Then
        MsgBox "This tab doesn't look like an email sheet (need To / Subject / Body headers)."
        Exit Sub
    End If

    For r = 2 To ws.Cells(ws.Rows.Count, cTo).End(xlUp).Row
        If Len(ws.Cells(r, cTo).Value) > 3 And InStr(ws.Cells(r, cTo).Value, "[") = 0 Then
            Set mail = ol.CreateItem(0)
            mail.To = ws.Cells(r, cTo).Value
            mail.Subject = ws.Cells(r, cSub).Value
            mail.Body = ws.Cells(r, cBody).Value
            If doSend Then mail.Send Else mail.Save
            n = n + 1
            Application.Wait Now + TimeValue("0:00:02")
        Else
            skipped = skipped + 1
        End If
    Next r

    MsgBox n & IIf(doSend, " emails SENT.", " drafts created - see Outlook's Drafts folder.") & _
           IIf(skipped > 0, vbCrLf & skipped & " row(s) skipped (empty or placeholder address).", "")
End Sub
