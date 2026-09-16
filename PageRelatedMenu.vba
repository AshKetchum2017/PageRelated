Option Explicit

' MacroRunner integration: no reference to the runner project is required.
Private pMRObserver As Object
Private pMRToken As String


Private Sub cmdCancel_Click()

    Unload Me
    
End Sub

Private Sub cmdProcess_Click()

    Dim doc As Document
    Dim pageValues() As String
    Dim pageValueCount As Long
    Dim processCount As Long
    Dim valueIndex As Long
    Dim pageIndex As Long
    Dim pageName As String
    Dim targetPage As Page
    Dim hasPageIndex As Boolean
    Dim requestedPageIndex As Long
    Dim nextRenamePageIndex As Long
    Dim unprocessedCount As Long
    Dim nextAddPageIndex As Long
    Dim maxAddPageIndex As Long
    Dim addPageIndices() As Long
    Dim addPageNames() As String
    Dim commandGroupOpen As Boolean

    On Error GoTo ProcessFailed

    Set doc = ActiveDocument
    If Len(Trim$(txtText.Text)) = 0 Then
        MsgBox "Masukkan value page terlebih dahulu.", vbExclamation, "Page Related"
        Exit Sub
    End If

    pageValues = ParsePageValues(txtText.Text, pageValueCount)
    If pageValueCount = 0 Then
        MsgBox "Masukkan value page terlebih dahulu.", vbExclamation, "Page Related"
        Exit Sub
    End If

    If Not optAddPage.Value And Not optRenPage.Value Then
        MsgBox "Pilih AddPage atau RenamePage terlebih dahulu.", vbExclamation, "Page Related"
        Exit Sub
    End If

    doc.BeginCommandGroup "Page Related"
    commandGroupOpen = True

    If optAddPage.Value Then
        nextAddPageIndex = doc.Pages.Count + 1
        ReDim addPageIndices(0 To pageValueCount - 1)
        ReDim addPageNames(0 To pageValueCount - 1)

        For valueIndex = 0 To pageValueCount - 1
            ParsePageValue pageValues(valueIndex), requestedPageIndex, hasPageIndex, pageName
            If Not hasPageIndex Then requestedPageIndex = nextAddPageIndex
            If requestedPageIndex < 1 Then Err.Raise 5, , "Nomor page harus lebih besar dari 0."

            addPageIndices(valueIndex) = requestedPageIndex
            addPageNames(valueIndex) = pageName
            If requestedPageIndex > maxAddPageIndex Then maxAddPageIndex = requestedPageIndex
            nextAddPageIndex = requestedPageIndex + 1
        Next valueIndex

        Do While doc.Pages.Count < maxAddPageIndex
            Set targetPage = doc.InsertPagesEx(1, False, doc.Pages.Count, _
                doc.Pages(1).SizeWidth, doc.Pages(1).SizeHeight)
        Loop

        For valueIndex = 0 To pageValueCount - 1
            If Len(addPageNames(valueIndex)) > 0 And addPageNames(valueIndex) <> "~" Then
                doc.Pages(addPageIndices(valueIndex)).Name = addPageNames(valueIndex)
            End If
        Next valueIndex
    Else
        nextRenamePageIndex = 1

        For valueIndex = 0 To pageValueCount - 1
            ParsePageValue pageValues(valueIndex), requestedPageIndex, hasPageIndex, pageName
            If Not hasPageIndex Then requestedPageIndex = nextRenamePageIndex

            If requestedPageIndex >= 1 And requestedPageIndex <= doc.Pages.Count Then
                If Len(pageName) > 0 And pageName <> "~" Then
                    doc.Pages(requestedPageIndex).Name = pageName
                End If
            Else
                unprocessedCount = unprocessedCount + 1
            End If

            If requestedPageIndex >= nextRenamePageIndex Then
                nextRenamePageIndex = requestedPageIndex + 1
            End If
        Next valueIndex

        If unprocessedCount > 0 Then
            On Error Resume Next
            If commandGroupOpen Then doc.EndCommandGroup
            commandGroupOpen = False
            On Error GoTo 0

            MsgBox unprocessedCount & _
                " value tidak diproses karena jumlah page kurang.", _
                vbExclamation, "Page Related"
            Exit Sub
        End If
    End If

CleanExit:
    On Error Resume Next
    If commandGroupOpen Then doc.EndCommandGroup
    On Error GoTo 0
    Unload Me
    Exit Sub

ProcessFailed:
    Dim errorNumber As Long
    Dim errorDescription As String

    errorNumber = Err.Number
    errorDescription = Err.Description

    On Error Resume Next
    If commandGroupOpen Then doc.EndCommandGroup
    On Error GoTo 0

    MsgBox "Error " & errorNumber & vbCrLf & _
        "Description: [" & errorDescription & "]", _
        vbCritical, "Page Related"

End Sub

Private Function ParsePageValues(ByVal inputText As String, _
                                 ByRef valueCount As Long) As String()
    Dim rawValues() As String
    Dim values() As String
    Dim valueIndex As Long

    inputText = Trim$(inputText)
    If Left$(inputText, 1) = "(" And Right$(inputText, 1) = ")" Then
        inputText = Trim$(Mid$(inputText, 2, Len(inputText) - 2))
    End If

    rawValues = Split(inputText, ",")
    ReDim values(0 To UBound(rawValues))

    For valueIndex = 0 To UBound(rawValues)
        values(valueIndex) = Trim$(rawValues(valueIndex))
    Next valueIndex

    valueCount = UBound(values) + 1
    ParsePageValues = values
End Function

Private Sub ParsePageValue(ByVal rawValue As String, _
                           ByRef pageIndex As Long, _
                           ByRef hasPageIndex As Boolean, _
                           ByRef pageName As String)
    Dim separatorIndex As Long
    Dim pageIndexText As String

    rawValue = Trim$(rawValue)
    separatorIndex = InStr(1, rawValue, ":", vbBinaryCompare)
    If separatorIndex = 0 Then
        hasPageIndex = False
        pageIndex = 0
        pageName = rawValue
        Exit Sub
    End If

    pageIndexText = Trim$(Left$(rawValue, separatorIndex - 1))
    If Not IsNumeric(pageIndexText) Then
        Err.Raise 5, , "Format nomor page tidak valid: " & rawValue
    End If

    hasPageIndex = True
    pageIndex = CLng(pageIndexText)
    pageName = Trim$(Mid$(rawValue, separatorIndex + 1))
End Sub

Private Sub optAddPage_Click()

End Sub

Private Sub optRenPage_Click()

End Sub

Private Sub txtText_Change()

End Sub

' Called only by MRTargetBridge; normal menu entry points remain unchanged.
Public Sub MRBindRunner(ByVal observer As Object, ByVal token As String)
    Set pMRObserver = observer
    pMRToken = token
End Sub

Public Sub MRDetachRunner()
    Set pMRObserver = Nothing
    pMRToken = vbNullString
End Sub

Private Sub UserForm_Terminate()
    Dim observer As Object, token As String
    On Error GoTo NotifyFailed
    Set observer = pMRObserver
    token = pMRToken
    MRDetachRunner
    If Not observer Is Nothing Then CallByName observer, "MacroUnloaded", VbMethod, token
    Exit Sub
NotifyFailed:
    MsgBox "Gagal memberitahu Macro Runner bahwa form sudah ditutup (" & CStr(Err.Number) & "): " & _
        Err.Description, vbExclamation, "Macro Runner"
End Sub
