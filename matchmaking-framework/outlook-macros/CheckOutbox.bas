Attribute VB_Name = "CheckOutbox"
Option Explicit

' READ-ONLY check: counts the emails waiting in the Outbox
' per scheduled release time. Touches nothing, sends nothing.

Sub CheckScheduledEmails()
    Dim ol As Object, ns As Object, ob As Object, it As Object
    Dim i As Long, k As String, msg As String
    Dim dic As Object
    Set dic = CreateObject("Scripting.Dictionary")

    On Error Resume Next
    Set ol = GetObject(, "Outlook.Application")
    If ol Is Nothing Then Set ol = CreateObject("Outlook.Application")
    On Error GoTo 0
    If ol Is Nothing Then MsgBox "Open classic Outlook first.": Exit Sub

    Set ns = ol.GetNamespace("MAPI")
    Set ob = ns.GetDefaultFolder(4)   ' 4 = Outbox

    If ob.Items.Count = 0 Then MsgBox "The Outbox is empty - nothing is queued.": Exit Sub

    For i = 1 To ob.Items.Count
        Set it = ob.Items(i)
        If it.DeferredDeliveryTime > #1/1/2000# Then
            k = Format(it.DeferredDeliveryTime, "ddd dd mmm yyyy, hh:nn")
        Else
            k = "NO timer (will send immediately!)"
        End If
        If dic.Exists(k) Then dic(k) = dic(k) + 1 Else dic.Add k, 1
    Next i

    msg = ob.Items.Count & " email(s) waiting in the Outbox:" & vbCrLf & vbCrLf
    Dim ky
    For Each ky In dic.Keys
        msg = msg & "  " & dic(ky) & "  emails  ->  " & ky & vbCrLf
    Next ky
    MsgBox msg, vbInformation, "Scheduled email check"
End Sub
