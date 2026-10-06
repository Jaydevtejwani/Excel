=COUNTA(INDEX(Sheet1!$B$2:$XFD$1048576,0,ROW(A1)))
=LET(x,TEXTSPLIT(A2,"|"),IF(ROWS(UNIQUE(TRIM(x)))<COUNTA(x),"DUPLICATE","OK"))



=LET(x,TRIM(TEXTSPLIT(A2,"|")),u,UNIQUE(x),FILTER(u,COUNTIF(x,u)>1,"No Duplicate"))



=LET(x,TRIM(TEXTSPLIT(A2,"|")),IF(COUNTA(UNIQUE(x))<COUNTA(x),"DUPLICATE","OK"))



=LET(x,TOCOL(TRIM(TEXTSPLIT(A2,"|")),1),u,UNIQUE(x),IFERROR(TEXTJOIN(" | ",TRUE,FILTER(u,COUNTIF(x,u)>1)),""))






Sub FindDuplicateDocumentIDs()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim arr As Variant
    Dim dict As Object
    Dim i As Long
    Dim docID As String
    Dim result As String
    Dim key As Variant

    Set ws = ActiveSheet
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'Header
    ws.Range("B1").Value = "Duplicate Document ID"

    For r = 2 To lastRow

        Set dict = CreateObject("Scripting.Dictionary")
        result = ""

        'Split Document IDs by |
        arr = Split(ws.Cells(r, "A").Value, "|")

        'Count each Document ID
        For i = LBound(arr) To UBound(arr)

            docID = Trim(arr(i))

            If docID <> "" Then
                If dict.Exists(docID) Then
                    dict(docID) = dict(docID) + 1
                Else
                    dict.Add docID, 1
                End If
            End If

        Next i

        'Find IDs appearing more than once
        For Each key In dict.Keys

            If dict(key) > 1 Then

                If result <> "" Then
                    result = result & " | "
                End If

                result = result & key

            End If

        Next key

        'Put result in Column B
        ws.Cells(r, "B").Value = result

    Next r

    MsgBox "Duplicate Document IDs found successfully.", vbInformation

End Sub




Option Explicit

Sub FindDuplicateDocumentIDs()

    Dim ws As Worksheet
    Dim lastRow As Long
    Dim r As Long
    Dim arr As Variant
    Dim dict As Object
    Dim i As Long
    Dim docID As String
    Dim result As String
    Dim key As Variant

    'Use active worksheet
    Set ws = ActiveSheet

    'Find last used row in Column A
    lastRow = ws.Cells(ws.Rows.Count, "A").End(xlUp).Row

    'Header
    ws.Range("B1").Value = "Duplicate Document ID"

    'Process each row
    For r = 2 To lastRow

        Set dict = CreateObject("Scripting.Dictionary")
        result = ""

        'Split Document IDs using |
        arr = Split(CStr(ws.Cells(r, "A").Value), "|")

        'Count each Document ID
        For i = LBound(arr) To UBound(arr)

            docID = Trim(CStr(arr(i)))

            If docID <> "" Then

                If dict.Exists(docID) Then
                    dict(docID) = dict(docID) + 1
                Else
                    dict.Add docID, 1
                End If

            End If

        Next i

        'Find IDs appearing more than once
        For Each key In dict.Keys

            If dict(key) > 1 Then

                If result <> "" Then
                    result = result & " | "
                End If

                result = result & CStr(key)

            End If

        Next key

        'Write result in Column B
        ws.Cells(r, "B").Value = result

    Next r

    'Auto-fit Column B
    ws.Columns("B").AutoFit

    MsgBox "Duplicate Document IDs found successfully.", _
           vbInformation, "Completed"

End Sub
