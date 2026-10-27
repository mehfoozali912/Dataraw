Attribute VB_Name = "FlushReminders"
Option Explicit

' Fixes stuck reminders in the Outbox:
'  - any Nigeria reminder whose release time HAS PASSED -> sends it NOW
'  - any whose release time is still in the future -> re-queues it properly
' Run it any morning the Outbox looks stuck. Classic Outlook must be open.

Sub FlushDueReminders()
    Dim ol As Object, ns As Object, ob As Object, it As Object
    Dim i As Long, sentNow As Long, requeued As Long
    Const NODATE As Date = #1/1/4501#

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then MsgBox "Open classic Outlook first.": Exit Sub

    Set ns = ol.GetNamespace("MAPI")
    Set ob = ns.GetDefaultFolder(4)   ' Outbox

    For i = ob.Items.Count To 1 Step -1
        Set it = ob.Items(i)
        If InStr(1, it.Subject & "", "Nigeria Energy 2026", vbTextCompare) > 0 Then
            If it.DeferredDeliveryTime <= Now Or it.DeferredDeliveryTime = NODATE Then
                it.DeferredDeliveryTime = NODATE   ' remove the timer
                it.Send                            ' goes immediately
                sentNow = sentNow + 1
            Else
                it.Send                            ' re-register, timer kept
                requeued = requeued + 1
            End If
            Application.Wait Now + TimeValue("0:00:01")
        End If
    Next i

    MsgBox "Sent NOW: " & sentNow & vbCrLf & _
           "Re-queued for later (timer kept): " & requeued & vbCrLf & vbCrLf & _
           "Check Sent Items in a minute.", vbInformation, "Flush reminders"
End Sub
