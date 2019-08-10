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
'Print arrays to file
Private Sub PrintArraysToFile(ByVal arrClass As Variant)
    Dim varSaveName As Variant             'Used for filename dialog
    Dim sFile As String                     'File to write into
    Dim sLineSpacing As String              'Used for indenting at different levels of progression
    Dim iStatEnumPosition As Long
    Dim i As Long
    Dim j As Long
    Dim k As Long
    Dim arrStats As Variant                 'Array to hold the stats range for each class
    
    'Create and open file for use
    varSaveName = Application.GetSaveAsFilename(FileFilter:="Text File (*.txt), *.txt")
    If varSaveName = False Then Exit Sub
    sFile = varSaveName
    
    Open sFile For Output As #1
    sLineSpacing = "  "
    
    'New line for easier copying
    Print #1, vbLf
    
    For i = LBound(arrClass, 1) To UBound(arrClass, 1)
        Print #1, sLineSpacing & "- characterClass: " & arrClass(i, 2)
        
        sLineSpacing = sLineSpacing & "  "
        Print #1, sLineSpacing & "stats:"
        
        Call GetStatArray(arrClass(i, 1), arrClass(i, 3), arrStats)
        
        For j = LBound(arrStats, 1) To UBound(arrStats, 1)
            
            If Not arrStats(j, 1) = Empty Then
                Print #1, sLineSpacing & "- stat: " & arrStats(j, 0)
                sLineSpacing = sLineSpacing & "  "
                Print #1, sLineSpacing & "levels:"
                'Starting on second element of array. First element identifies Enum position
                For k = LBound(arrStats, 2) + 1 To UBound(arrStats, 2)
                    Print #1, sLineSpacing & "- " & arrStats(j, k)
                Next k
            sLineSpacing = "    "
            End If
        Next j
        sLineSpacing = "  "
        Erase arrStats
    Next i
    
    Close #1

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
    Dim rngFindStatInClass          'Find stat name in player or enemy ranges
    
    iLevels = ws.Cells(1, ws.Columns.Count).End(xlToLeft).Value
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
                arrStats(i, j) = arrTemp(1, j)
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
