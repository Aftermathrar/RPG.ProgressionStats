Attribute VB_Name = "ResizeNamedRange"
Option Explicit

'Update the size of the Enemies named range
Public Sub ExtendEnemiesNamedRange()
    Dim iFirstRow As Long
    Dim iLastRow As Long
    Dim rngEnemy As Range
    
    Set rngEnemy = Worksheets("Key Stats").Range("Enemies")
    
    iFirstRow = rngEnemy(1, 1).Row
    iLastRow = Worksheets("Key Stats").Cells(Worksheets("Key Stats").Rows.Count, 1).End(xlUp).Row
    
    rngEnemy.Resize(iLastRow - iFirstRow + 1, 1).Name = "Enemies"
    
End Sub
'Update the size of the Player_Details named range
Public Sub ExtendPlayerNamedRange()
    Dim iFirstRow As Long
    Dim iLastRow As Long
    Dim rngPlayer As Range
    
    Set rngPlayer = Worksheets("Key Stats").Range("Player_Details")
    
    iFirstRow = rngPlayer(1, 1).Row
    iLastRow = rngPlayer(1, 1).End(xlDown).Row
    
    rngPlayer.Resize(iLastRow - iFirstRow + 1, 1).Name = "Player_Details"
    
End Sub
'Update the size of the Base_Enemy_Details named range
Public Sub ExtendBaseEnemiesNamedRange()
    Dim iFirstRow As Long
    Dim iLastRow As Long
    Dim rngBaseDetails As Range
    
    Set rngBaseDetails = Worksheets("Key Stats").Range("Base_Enemy_Details")
    
    iFirstRow = rngBaseDetails(1, 1).Row
    iLastRow = rngBaseDetails(1, 1).End(xlDown).Row
    
    rngBaseDetails.Resize(iLastRow - iFirstRow + 1, 1).Name = "Base_Enemy_Details"
    
End Sub
