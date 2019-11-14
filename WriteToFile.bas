Attribute VB_Name = "WriteToFile"
Option Explicit
Dim ws As Worksheet
'Create a file containing all class stat values for the Progression asset
Sub WriteStatsToFile()
    Dim arrClass As Variant
    
    Set ws = Worksheets("Key Stats")
    
    arrClass = Worksheets("Enumerations").ListObjects("tblCharacterClasses").DataBodyRange
    
    Call PrintArraysToFile(arrClass)
    
    Erase arrClass
    
End Sub
'Get filepath based on FileName
Private Function GetSaveName(ByVal sFile As String, ByVal sFilter As String, ByRef sNamespace As String, ByRef sEnum As String)
    Dim varSaveName As Variant              'Used for filename dialog
    Dim tblFile As ListObject
    Dim rngTemp As Range
    
    Set tblFile = Worksheets("Filepaths").ListObjects("tblFilePath")
    
    If tblFile.Parent.AutoFilterMode Then
        tblFile.AutoFilter.ShowAllData
    End If
    
    tblFile.Range.AutoFilter field:=1, Criteria1:=sFile
    Set rngTemp = tblFile.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible)
        
    varSaveName = rngTemp(1, 2).Value
    sNamespace = rngTemp(1, 3).Value2
    sEnum = rngTemp(1, 4).Value2
    
    If varSaveName = Empty Then
        varSaveName = Application.GetSaveAsFilename(FileFilter:=sFilter)
        If varSaveName = False Then Exit Function
        rngTemp(1, 2).Value = varSaveName
    End If
    
    tblFile.AutoFilter.ShowAllData
    
    GetSaveName = varSaveName
    
    Set varSaveName = Nothing
    Set tblFile = Nothing
    Set rngTemp = Nothing
    
End Function
'Print arrays to file
Private Sub PrintArraysToFile(ByVal arrClass As Variant)
    Dim sFile As String                     'File to write into
    Dim sLineSpacing As String              'Used for indenting at different levels of progression
    Dim iStatEnumPosition As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim arrStats As Variant                 'Array to hold the stats range for each class
    
    'Check if file path has been set, otherwise open file dialog
    sFile = "Progression Stats"         'Identifier for lookup
    sFile = GetSaveName(sFile, "Text File (*.txt), *.txt", sLineSpacing, sLineSpacing)
    
    If Len(sFile) = 0 Then Exit Sub
    
    'Create and open file for use
    Open sFile For Output As #1
    sLineSpacing = Space(2)
    
    Print #1, "Paste this into Progression.asset"
    
    For i = LBound(arrClass, 1) To UBound(arrClass, 1)
        Print #1, sLineSpacing & "- characterClass: " & arrClass(i, 2)
        
        sLineSpacing = sLineSpacing & Space(2)
        Print #1, sLineSpacing & "stats:"
        
        Call GetStatArray(arrClass(i, 1), arrClass(i, 3), arrStats)
        
        For j = LBound(arrStats, 1) To UBound(arrStats, 1)
            
            If Not arrStats(j, 1) = Empty Then
                Print #1, sLineSpacing & "- stat: " & arrStats(j, 0)
                sLineSpacing = sLineSpacing & Space(2)
                Print #1, sLineSpacing & "levels:"
                'Starting on second element of array. First element identifies Enum position
                For k = LBound(arrStats, 2) + 1 To UBound(arrStats, 2)
                    Print #1, sLineSpacing & "- " & arrStats(j, k)
                Next k
            sLineSpacing = Space(4)
            End If
        Next j
        sLineSpacing = Space(2)
        Erase arrStats
    Next i
    
    Close #1

End Sub
'Overwrite reference enumeration file
Private Sub OverwriteEnumerationsFile(ByVal sFile As String, ByVal rngEnum As Range, ByVal strNameSpace As String, ByVal strEnumName As String)
    Dim c As Range
    Dim sEnumEntry As String
    Dim sLineSpacing As String
    
    Open sFile For Output As #1
    
    'Namespace and enum
    Print #1, "namespace " & strNameSpace & vbNewLine & _
        "{" & vbNewLine & Space(4) & _
        "public enum " & strEnumName & vbNewLine & Space(4) & _
        "{"
        
    sLineSpacing = Space(8)
    
    'Each entry in table gets added to enum
    For Each c In rngEnum.Cells
        sEnumEntry = Replace(c.Value, " ", "")
        If c.Row <= rngEnum.Rows.Count Then
            Print #1, sLineSpacing & sEnumEntry & ","
        Else
            Print #1, sLineSpacing & sEnumEntry
        End If
    Next c
    
    'Closing brackets
    Print #1, Space(4) & "}" & vbNewLine & "}"
    
    Close #1
    
    Set c = Nothing
    
End Sub
'Overwrite Class enumeration file for newly added classes
Public Sub OverwriteClassEnumerationsFile()
    Dim rngClass As Range
    Dim sFile As String
    Dim strNameSpace As String
    Dim strEnumName As String
    
    sFile = GetSaveName("Character Class", "C# Script (*.cs), *.cs", strNameSpace, strEnumName)
    Set rngClass = Worksheets("Enumerations").ListObjects("tblCharacterClasses").DataBodyRange.Columns(1)
    
    Call OverwriteEnumerationsFile(sFile:=sFile, rngEnum:=rngClass, strNameSpace:=strNameSpace, strEnumName:=strEnumName)
    
    Set rngClass = Nothing
    Set cClass = Nothing
    
End Sub
'Overwrite Stat enumeration file for newly added stats
Public Sub OverwriteStatEnumerationsFile()
    Dim rngStat As Range
    Dim sFile As String
    Dim strNameSpace As String
    Dim strEnumName As String
    
    sFile = GetSaveName("Stats", "C# Script (*.cs), *.cs", strNameSpace, strEnumName)
    Set rngStat = Worksheets("Enumerations").ListObjects("tblStats").DataBodyRange.Columns(1)
    
    Call OverwriteEnumerationsFile(sFile:=sFile, rngEnum:=rngStat, strNameSpace:=strNameSpace, strEnumName:=strEnumName)
    
    Set rngStat = Nothing
    
End Sub
'By Character Class, find stat range and throw into an array
Private Sub GetStatArray(ByVal sClass As String, ByVal iStats As Long, ByRef arrStats As Variant)
    Dim arrTemp As Variant          'Temp array to transfer stat rows to main array
    Dim iLevels As Integer          'Width of stat table, should be worksheet value at some point
    Dim i As Integer
    Dim j As Integer
    Dim rngStats As Range           'Stat table to look up stat names
    Dim rngEnemyDetails As Range    'Find enemy stat range by name
    Dim cStat As Range
    Dim rngFindStatInClass As Range 'Find stat name in player or enemy ranges
    
    iLevels = ws.Cells(3, ws.Columns.Count).End(xlToLeft).Value
    Set rngStats = Worksheets("Enumerations").ListObjects("tblStats").DataBodyRange.Columns(1)
    Set rngEnemyDetails = ws.Range("Enemies")
    
    'arrStat height is number of stats for the class
    'arrStat width is number of configured levels + 1. First element is stat position in the enum
    ReDim arrStats(iStats - 1, iLevels)
    
        
    'Set for first row of data to array
    i = 0
    
    For Each cStat In rngStats.Cells
        If LCase(sClass) = "player" Then        'Lookup stat in player detail range
            Set rngFindStatInClass = ws.Range("Player_Details").Find(cStat.Value, , xlValues, xlWhole)
        Else                                    'Lookup stat in Enemy detail range
            Set rngEnemyDetails = rngEnemyDetails.Find(sClass).Resize(1 + iStats, 1)
            Set rngFindStatInClass = rngEnemyDetails.Find(cStat.Value, , xlValues, xlWhole)
        End If
        
        If Not rngFindStatInClass Is Nothing Then       'Set stat row in temp array and loop it into arrStat
            arrTemp = rngFindStatInClass.Offset(0, 3).Resize(1, 20).Value
            arrStats(i, 0) = cStat.Offset(0, 1).Value
            For j = LBound(arrStats, 2) + 1 To UBound(arrStats, 2)
                arrStats(i, j) = Round(arrTemp(1, j), 1)
            Next j
            
            'Advance arrStat row
            i = i + 1
        End If
    Next cStat
    
    Erase arrTemp
    Set rngStats = Nothing
    Set cStat = Nothing
    Set rngFindStatInClass = Nothing
                    
End Sub
