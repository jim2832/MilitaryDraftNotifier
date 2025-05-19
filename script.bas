Attribute VB_Name = "Module1"
Sub CheckColumnPairsAndNotifyLINE()
    Dim ws As Worksheet
    Set ws = ThisWorkbook.Sheets(1)

    Dim row As Long
    Dim name As String
    Dim rocDateStr As String
    Dim targetDate As Date
    Dim daysLeft As Long
    Dim messageBody As String
    Dim hasAlert As Boolean
    hasAlert = False

    Dim msgLines As Collection
    Set msgLines = New Collection
    
    ' ��l�Ƽ��D
    messageBody = "�H�U�Шk���X�ҴN�ǲפ����֭n�L���A�гq���Шk������T�C"
    
    ' �����C���Шk���
    For row = 2 To ws.UsedRange.Rows.Count
        name = Trim(ws.Cells(row, 2).Value)
        rocDateStr = Trim(ws.Cells(row, 5).Value)
    
        If name <> "" And rocDateStr <> "" Then
            On Error GoTo SkipRow
            targetDate = ROCtoAD(rocDateStr)
            daysLeft = DateDiff("d", Date, targetDate)
    
            If daysLeft >= 0 And daysLeft <= 365 Then
                msgLines.Add name & ": �Ѿl" & daysLeft & "��"
                hasAlert = True
            End If
SkipRow:
            On Error GoTo 0
        End If
    Next row
    
    ' �զX�T�����e�A����W�٤�����̫�@�椣�h�[����
    If hasAlert Then
        messageBody = messageBody & "\n\n" & Join(CollectionToArray(msgLines), "\n")
        Debug.Print messageBody
        Call SendLineBroadcast(messageBody)
    Else
        Debug.Print "�S�����`�ƭȡC"
    End If
End Sub


Function ROCtoAD(rocDateStr As String) As Date
    Dim parts() As String
    Dim y As Integer, m As Integer, d As Integer
    parts = Split(rocDateStr, "/")

    If UBound(parts) = 2 Then
        y = CInt(parts(0)) + 1911
        m = CInt(parts(1))
        d = CInt(parts(2))
        ROCtoAD = DateSerial(y, m, d)
    Else
        Err.Raise vbObjectError + 513, , "����榡���~"
    End If
End Function

Sub SendLineBroadcast(ByVal msg As String)
    Dim url As String
    Dim token As String
    Dim json As String

    url = "https://api.line.me/v2/bot/message/broadcast"
    token = "VzXTTodOWnD9TP+sByY7V/2o0Z5B9TLqFrbizuzRzH/bmWe3dscDViVWmUrqEyHfSy44NybtNgxhnYsYHRK8fC+W6bc+L3i6+EHBaoCb3LVjjERIkgX9+wE+JpZXW5aqW20l+1HpOiiyWnwqzORZ8wdB04t89/1O/w1cDnyilFU="

    json = "{""messages"":[{""type"":""text"",""text"":""" & msg & """}]}"

    Dim http As Object
    Set http = CreateObject("MSXML2.XMLHTTP")

    With http
        .Open "POST", url, False
        .setRequestHeader "Content-Type", "application/json"
        .setRequestHeader "Authorization", "Bearer " & token
        .send json
    End With

    Debug.Print "LINE �s���T���w�o�e�G" & msg
End Sub

Function CollectionToArray(col As Collection) As Variant
    Dim arr() As String
    ReDim arr(0 To col.Count - 1)
    Dim i As Long
    For i = 1 To col.Count
        arr(i - 1) = col(i)
    Next i
    CollectionToArray = arr
End Function
