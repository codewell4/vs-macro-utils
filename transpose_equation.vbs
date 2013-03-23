Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE90
Imports EnvDTE90a
Imports EnvDTE100
Imports System.Diagnostics
Public Module TransposeEquation
    Sub TransposeSelection()

        Dim objDocument As EnvDTE.Document
        Dim objTextDocument As EnvDTE.TextDocument
        Dim objTextSelection As EnvDTE.TextSelection
        Dim TransposedText As String
        Dim TransposedTextArray() As String
        Dim TransposedText_Start As String
        Dim TransposedText_End As String
        Dim pos As Integer

        Try
            ' Get the active document
            objDocument = DTE.ActiveDocument

            ' Get the text document
            objTextDocument = CType(objDocument.Object, EnvDTE.TextDocument)

            ' Get the text selection object
            objTextSelection = objTextDocument.Selection
            TransposedText = Trim(objTextSelection.Text)

            TransposedTextArray = TransposedText.Replace(vbCrLf, vbLf).Split(vbLf)

            ' Show some properties of the text selection
            Dim LastNonEmpty As Integer = -1
            For i As Integer = 0 To TransposedTextArray.Length - 1
                If Trim(TransposedTextArray(i)) <> "" Then
                    ' Show some properties of the text selection
                    TransposedText = Trim(TransposedTextArray(i))
                    pos = InStr(1, TransposedText, "=")
                    LastNonEmpty += 1
                    If pos > 0 Then
                        TransposedText_Start = Trim(TransposedText.Substring(0, pos - 2))
                        TransposedText_End = Trim(TransposedText.Substring(pos, Len(TransposedText) - pos))
                        If (InStr(1, TransposedText_End, ";") > 0) Then
                            TransposedText_End = TransposedText_End.Replace(";", "")
                            TransposedText_Start = TransposedText_Start + ";"
                        End If
                        TransposedTextArray(LastNonEmpty) = TransposedText_End + " = " + TransposedText_Start
                    Else
                        TransposedTextArray(LastNonEmpty) = TransposedTextArray(i)
                    End If
                End If
            Next
            ReDim Preserve TransposedTextArray(LastNonEmpty)
            TransposedText = TransposedText.Join(vbCrLf, TransposedTextArray)

            DTE.ActiveDocument.Selection.Insert(TransposedText, vsInsertFlags.vsInsertFlagsContainNewText)
        Catch objException As System.Exception
            MsgBox(objException.ToString)
        End Try
    End Sub
End Module
