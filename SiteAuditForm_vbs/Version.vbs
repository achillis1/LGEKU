Attribute VB_Name = "Version"
Sub SaveCodeModules()
' http://stackoverflow.com/questions/131605/best-way-to-do-version-control-for-ms-excel
'This code Exports all VBA modules
Dim i%, sName$

With ThisWorkbook.VBProject
    For i% = 1 To .VBComponents.Count
        If .VBComponents(i%).CodeModule.CountOfLines > 0 Then
            sName$ = .VBComponents(i%).CodeModule.Name
            .VBComponents(i%).Export "C:\share\DriveZ\LGE\documents\repository\SiteAuditForm_vbs\" & sName$ & ".vbs"
        End If
    Next i
End With

End Sub

'Sub ImportCodeModules()
'
'With ThisWorkbook.VBProject
'    For i% = 1 To .VBComponents.Count
'
'        ModuleName = .VBComponents(i%).CodeModule.Name
'
'        If ModuleName <> "VersionControl" Then
'            If Right(ModuleName, 6) = "Macros" Then
'                .VBComponents.Remove .VBComponents(ModuleName)
'                .VBComponents.Import "X:\Data\MySheet\" & ModuleName & ".vba"
'           End If
'        End If
'    Next i
'End With
'
'End Sub
