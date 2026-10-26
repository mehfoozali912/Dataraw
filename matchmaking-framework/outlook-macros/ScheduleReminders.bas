Attribute VB_Name = "ScheduleReminders"
Option Explicit

' Queues every email on the ACTIVE tab with a timed release
' (Outlook "deferred delivery"). Run it once per day-tab:
'   stand on "Day 1 Reminders" -> run -> type Day 1's date
'   stand on "Day 2 Reminders" -> run -> type Day 2's date, etc.
' Emails wait in classic Outlook's OUTBOX and leave at the set time.
' RULE: laptop ON + classic Outlook OPEN at release time, or they
' leave the moment you next open classic Outlook that day.

Sub ScheduleDayReminders()
    Dim ws As Worksheet, ol As Object, mail As Object
    Dim r As Long, c As Long, n As Long
    Dim cTo As Long, cSub As Long, cBody As Long
    Dim dStr As String, tStr As String, dt As Date

    Set ws = ActiveSheet
    dStr = InputBox("Date to DELIVER the '" & ws.Name & "' emails" & vbCrLf & _
                    "(example: 10/11/2026 for 10 Nov 2026):", "Delivery date")
    If Len(dStr) = 0 Then Exit Sub
    If Not IsDate(dStr) Then MsgBox "Could not read that date - try like 10/11/2026": Exit Sub

    tStr = InputBox("Release time on YOUR laptop clock (Dubai time)." & vbCrLf & _
                    "11:00 Dubai = 08:00 in Lagos, Nigeria.", "Release time", "11:00")
    If Len(tStr) = 0 Then Exit Sub
    If Not IsDate(tStr) Then MsgBox "Could not read that time - try like 11:00": Exit Sub

    dt = CDate(dStr) + CDate(tStr)
    If dt <= Now Then MsgBox "That time is in the past (" & Format(dt, "ddd dd mmm yyyy hh:nn") & ").": Exit Sub

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then MsgBox "Open classic Outlook first.": Exit Sub

    For c = 1 To ws.Cells(1, ws.Columns.Count).End(xlToLeft).Column
        Select Case True
            Case ws.Cells(1, c).Value Like "To*":      cTo = c
            Case ws.Cells(1, c).Value Like "Subject*": cSub = c
            Case ws.Cells(1, c).Value Like "*Body*":   cBody = c
        End Select
    Next c
    If cTo = 0 Or cSub = 0 Or cBody = 0 Then MsgBox "This tab has no To/Subject/Body headers.": Exit Sub

    If MsgBox("Queue the '" & ws.Name & "' emails to leave at:" & vbCrLf & vbCrLf & _
              Format(dt, "dddd dd mmmm yyyy, hh:nn") & "  (your laptop time)?", _
              vbYesNo + vbQuestion, "Schedule reminders") <> vbYes Then Exit Sub

    For r = 2 To ws.Cells(ws.Rows.Count, cTo).End(xlUp).Row
        If Len(ws.Cells(r, cTo).Value) > 3 And InStr(ws.Cells(r, cTo).Value, "[") = 0 Then
            Set mail = ol.CreateItem(0)
            mail.To = ws.Cells(r, cTo).Value
            mail.Subject = ws.Cells(r, cSub).Value
            mail.Body = ws.Cells(r, cBody).Value
            mail.DeferredDeliveryTime = dt
            mail.Send
            n = n + 1
            Application.Wait Now + TimeValue("0:00:01")
        End If
    Next r

    MsgBox n & " reminder(s) queued in the OUTBOX." & vbCrLf & _
           "They will leave at " & Format(dt, "ddd dd mmm, hh:nn") & " (laptop time)." & vbCrLf & vbCrLf & _
           "Keep your laptop on with classic Outlook OPEN at that time." & vbCrLf & _
           "To cancel or change: open the Outbox and delete them.", vbInformation
End Sub
