Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE90
Imports EnvDTE90a
Imports EnvDTE100
Imports System.Diagnostics

Public Module CreateGUID
    Sub CreateGUIDL()
        Dim NewGUID As String = Guid.NewGuid().ToString().ToLower()
        DTE.ActiveDocument.Selection.Insert(NewGUID, vsInsertFlags.vsInsertFlagsContainNewText)
    End Sub
End Module
