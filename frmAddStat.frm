VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddStat 
   Caption         =   "Add Base Stat Type"
   ClientHeight    =   6630
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   4710
   OleObjectBlob   =   "frmAddStat.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddStat"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Initialize form and Class list
Private Sub UserForm_Initialize()
    
    Call PopulateClassList
    Call PopulateOperatorComboBox
    
End Sub
'Submit new stat
Private Sub btnSubmit_Click()
    
    If Not IsFormFilledOut Then Exit Sub
    
    Application.ScreenUpdating = False
    
    Select Case mpgStatTypes.Value
        Case 0          'Generic stat
            'Cycle through class list
            If lboClass.Selected(0) Then
                Call AddToPlayerTable(False)
                If chkAddToBase Then
                    'Set base multiplier as main multiplier and reference to player stat
                    txtMultiplier = txtBaseMultiplier
                    refScaleSelect.Value = Range("Player_Details").Find(txtStatName, , xlValues, xlWhole).Address
                    If Not IsReferenceValid(refScaleSelect) Then Exit Sub
                    If cboRefOperator.Value = "No scaling" Then
                        cboRefOperator.Value = "*"
                        Call BuildFormula
                    End If
                End If
            End If
            
            If chkAddToBase Then
                Call AddToBaseTable
            End If
            
            'Add to stat table
            Call AddToStatTable(False)
            
            'Add to enemy, using base stat as reference if no scaling
            Call AddToEnemyTable(chkAddToBase)
            
        Case 1          'Player Only
            Call AddToPlayerTable(True)
            
            If chkAddToStatTable Then Call AddToStatTable(True)
            
        Case 2          'Base Enemy stat
            Call AddToBaseTable
    End Select
    
    Me.Hide
    
    Call OverwriteStatEnumerationsFile
    Call WriteStatsToFile
    
    Application.ScreenUpdating = True
    
End Sub
'Add stat to reference Details section
Private Sub AddToDetailsTable(ByVal rngDetails As Range, ByVal rngNewStat As Range, ByVal isPlayerOnly As Boolean, ByVal isEnemy As Boolean, ByVal isEnemyBase As Boolean)
    Dim iLevels As Integer          'Number of levels in table. Width of range to fill
    Dim strFormulaTemp As String    'Used to hold temp formula replacements
    
    iLevels = Range("D3").End(xlToRight).Value2
    
    If rngNewStat Is Nothing Then
        'Add a new stat row at the bottom of the section
        Set rngNewStat = rngDetails(rngDetails.Rows.Count, 1).Offset(1, 0)
        rngNewStat.Offset(1, 0).EntireRow.Insert (xlShiftDown)
    End If
    
    'Update formula
    strFormulaTemp = LCase(txtFormula.Value)
    If Left(strFormulaTemp, 1) <> "=" Then
        strFormulaTemp = "=" & strFormulaTemp
    End If
    strFormulaTemp = Replace(strFormulaTemp, "scale", rngNewStat.Offset(0, 1).Address)
    strFormulaTemp = Replace(strFormulaTemp, "ref", refScaleSelect.Value)
    
    'Start setting values in each column
    If isEnemyBase Then
        rngNewStat.Value = "Base " & txtStatName
    Else
        rngNewStat.Value = txtStatName
    End If
    
    rngNewStat.Offset(0, 1).Value = CDbl(txtMultiplier.Value)
    
    Set rngNewStat = rngNewStat.Offset(0, 3)
    
    If isPlayerOnly Then
        If chkFixedLevelIncrease Then
            rngNewStat.Offset(0, -1).Value = CDbl(txtFixedLevelIncrease)
            rngNewStat.Offset(0, 1) = "=" & rngNewStat.Address(True, False) & "+" & rngNewStat.Offset(0, -1).Address(True, True)
            rngNewStat.Offset(0, 1).Resize(1, iLevels - 1).FillRight
        Else
            If cboRefOperator.Value <> "No scaling" Then
                rngNewStat.Value = strFormulaTemp
            Else
                rngNewStat.Value = CDbl(txtMultiplier.Value)
            End If
            rngNewStat.Resize(1, iLevels).FillRight
        End If
        
        If chkHasStartValue Then
            rngNewStat.Value = CDbl(txtStartValue.Value)
        End If
    Else
        If cboRefOperator.Value <> "No scaling" Or isEnemy Then
            rngNewStat.Value = strFormulaTemp
        Else
            rngNewStat.Value = CDbl(txtMultiplier.Value)
        End If
        rngNewStat.Resize(1, iLevels).FillRight
    End If
    
    Set rngDetails = Nothing
    Set rngNewStat = Nothing

End Sub
'Add stat to Player Details section
Private Sub AddToPlayerTable(ByVal isPlayerOnly As Boolean)
    Dim rngNewStat As Range
    
    'Check if stat exists in table
    Set rngNewStat = Range("Player_Details").Find(txtStatName, , xlValues, xlWhole)
    If rngNewStat Is Nothing Then
        Worksheets("Enumerations").ListObjects("tblCharacterClasses").Range(2, 3).Value = Worksheets("Enumerations").ListObjects("tblCharacterClasses").Range(2, 3).Value + 1
    End If
    
    Call AddToDetailsTable(Range("Player_Details"), rngNewStat, isPlayerOnly, False, False)
    
    Call ExtendPlayerNamedRange
    
    Set rngNewStat = Nothing
    
End Sub
'Add stat to Base Enemy Details section
Private Sub AddToBaseTable()
    Dim rngNewStat As Range
    
    Set rngNewStat = Range("Base_Enemy_Details").Find(txtStatName, , xlValues, xlWhole)
    If rngNewStat Is Nothing Then
        Call AddToDetailsTable(Range("Base_Enemy_Details"), rngNewStat, False, False, True)
    End If

    Call ExtendBaseEnemiesNamedRange
    
    Set rngNewStat = Nothing
    
End Sub
'Add stat to Enemy Details section
Private Sub AddToEnemyTable(ByVal hasBaseReference As Boolean)
    Dim i As Long
    Dim rngTemp As Range
    Dim rngEnemy As Range
    Dim rngNewStat As Range
    Dim tblClasses As ListObject
    
    Set tblClasses = Worksheets("Enumerations").ListObjects("tblCharacterClasses")
    Set rngEnemy = Range("Enemies")
    
    'Update formula
    If hasBaseReference Then
        refScaleSelect.Value = Range("Base_Enemy_Details")(Range("Base_Enemy_Details").Rows.Count, 1).Address
        If Not IsReferenceValid(refScaleSelect) Then Exit Sub
        txtMultiplier.Value = 1
        txtFormula.Value = "=scale * ref"
    End If
    
    'Loop through enemy class selection and send the individual ranges for stat addition
    For i = 1 To tblClasses.ListRows.Count - 1
        If lboClass.Selected(i) Then
            'Find enemy name and resize to the existing details range
            Set rngTemp = rngEnemy.Find(lboClass.List(i), , xlValues, xlWhole). _
                Resize(tblClasses.ListRows(i + 1).Range(1, 3).Value + 1, 1)
                
            Set rngNewStat = rngTemp.Find(txtStatName, , xlValues, xlWhole)
            If rngNewStat Is Nothing Then
                'If enemy is selected, update stat table with new stat number
                tblClasses.ListRows(i + 1).Range(1, 3).Value = tblClasses.ListRows(i + 1).Range(1, 3).Value + 1
    
                Call AddToDetailsTable(rngTemp, rngNewStat, False, True, False)
            End If
        End If
    Next i
    
    Call ExtendEnemiesNamedRange
    
    Set rngNewStat = Nothing

End Sub
'Add new stat to enumerations table
Private Sub AddToStatTable(ByVal isPlayerOnly As Boolean)
    Dim tblStats As ListObject
    Dim rngNewStat As Range
    Dim arrTemp As Variant
    
    'Check if stat has an entry on the table
    Set tblStats = Worksheets("Enumerations").ListObjects("tblStats")
    Set rngNewStat = tblStats.ListColumns(1).Range.Find(txtStatName.Value, , xlValues, xlWhole)
    
    'If no previous entry, set up row for new entry
    'Old values will be overwritten
    If rngNewStat Is Nothing Then
        With tblStats
            .ListRows.Add
            Set rngNewStat = .ListRows(.ListRows.Count).Range
        End With
    Else
        Set rngNewStat = rngNewStat.Resize(1, tblStats.ListColumns.Count)
    End If
    
    'Store values in temp array to place in table
    arrTemp = rngNewStat
    arrTemp(1, 1) = txtStatName.Value
    arrTemp(1, 2) = tblStats.ListRows.Count - 1
    arrTemp(1, 3) = Replace(txtStatName, " ", "")
    arrTemp(1, 4) = "=IFNA(ADDRESS(MATCH([@[Multiplied Stat Name]],'Key Stats'!$A$1:$A$50,0), 4, 2, 1),"""")"
    
    'Skip if from isPlayerOnly submission
    If isPlayerOnly Then
        'Do nothing
    'If scaling for enemy section, add string reference to table
    ElseIf chkAddToBase Then
        'Enemy scaler must draw from newly added base stat
        arrTemp(1, 5) = "Base " & txtStatName.Value
    ElseIf cboRefOperator.Value <> "No scaling" Then
        'refScaleSelect should always point to column D, so -3 offset gives up stat name
        arrTemp(1, 5) = Range(refScaleSelect.Value).Offset(0, -3).Value2
    End If
    
    rngNewStat.Value = arrTemp
    
    Erase arrTemp
    Set tblStats = Nothing
    Set rngNewStat = Nothing
    
End Sub
'Check if required parts of form are filled out
Private Function IsFormFilledOut() As Boolean
    Dim answer As Integer
    
    If txtStatName.Value = Empty Then
        MsgBox "Please name the stat you'd like to add."
        txtStatName.SetFocus
        Exit Function
    End If
    
    'Set proper case
    txtStatName.Value = StrConv(txtStatName.Value, vbProperCase)
    
    'Check if stat exists in Enumeration table. Base Enemy stats aren't in table, checked for later.
    If Not Range("tblStats[Stats]").Find(txtStatName, , xlValues, xlWhole) Is Nothing Then
        answer = MsgBox("Stat name already exists, some data may be overwritten." _
            & "Is this okay?", vbCritical + vbYesNo, "Possible data loss")
        If answer = vbNo Then
            txtStatName.SetFocus
            Exit Function
        End If
    End If
    
    'If using PlayerOnly, can skip scaling checks if using fixed increases
    If mpgStatTypes.Value = 1 And chkFixedLevelIncrease Then
        GoTo ByPassOperatorCheck
    End If
    
    If cboRefOperator.Value = "Formula" Then        'Check if custom formula is used
        If Not IsFormulaValid Then
            MsgBox ("Check your formula, does not return a result.")
            txtFormula.SetFocus
            Exit Function
        Else
            GoTo ByPassOperatorCheck                'If formula checks out, no need for next check
        End If
    End If
    
    'If using an operator, build the formula and check if it's valid
    If cboRefOperator.Value <> "No scaling" Then
        If Not IsReferenceValid(refScaleSelect) Then
            MsgBox "Cell Reference is not valid, please select a cell."
            Exit Function
        End If
        
        Call BuildFormula
        
        If Not IsFormulaValid Then
            MsgBox "Check that your Multiplier and Cell Reference are set up correctly."
            Exit Function
        End If
    End If
    
ByPassOperatorCheck:
    Select Case mpgStatTypes.Value
        Case 0            'Generic Stat
            If Not IsClassSelected Then
                MsgBox "Please select at least one class for the new stat."
                Exit Function
            End If
            If chkAddToBase And Not IsNumeric(txtBaseMultiplier.Value) Then
                MsgBox "Base multiplier valie is not a number."
                txtBaseMultiplier.SetFocus
                Exit Function
            End If
        Case 1            'Player Only Stat
            If chkHasStartValue And Not IsNumeric(txtStartValue.Value) Then
                MsgBox "Starting value is not a number."
                txtStartValue.SetFocus
                Exit Function
            End If
            If chkFixedLevelIncrease And Not IsNumeric(txtFixedLevelIncrease.Value) Then
                MsgBox "Fixed Level Increase value is not a number."
                txtFixedLevelIncrease.SetFocus
                Exit Function
            End If
        Case 2            'Base Enemy
            'Append Base to the start of the stat name
            If Left(txtStatName.Value, 4) <> "Base" Then
                txtStatName.Value = "Base " & txtStatName
            End If
            'Check for duplicate stat name
            If Not Range("Base_Enemy_Details").Find(txtStatName.Value, , xlValues, xlWhole) Is Nothing Then
                MsgBox "Stat name already exists, please choose another."
                txtStatName.SetFocus
                Exit Function
            End If
    End Select

    IsFormFilledOut = True
    
End Function
'Build formula from entered data
Private Sub BuildFormula()

        If tglOrder Then        'Scaler + operator + reference
            txtFormula.Value = "=" & txtMultiplier.Value & cboRefOperator.Value & refScaleSelect.Value
        Else                    'Reference + operator + scaler
            txtFormula.Value = "=" & refScaleSelect.Value & cboRefOperator.Value & txtMultiplier.Value
        End If
        
End Sub
'Check if any classes are selected in lboClass
Private Function IsClassSelected() As Boolean
    Dim i As Integer
    
    With lboClass
        For i = 0 To .ListCount - 1
            If .Selected(i) Then Exit For
        Next i
        If i = .ListCount Then Exit Function
    End With
    
    IsClassSelected = True
    
End Function
'Check if formula returns a number
Private Function IsFormulaValid() As Boolean
    Dim sFormula As String
    Dim dblMultiplier As Double
    Dim dblResult As Double
    
    sFormula = LCase(txtFormula.Text)
    'Append equals sign if absent
    If Left(sFormula, 1) <> "=" Then
        sFormula = "=" & sFormula
    End If
    
    'Check if scale reference filled in and valid
    If Len(refScaleSelect.Value) > 0 Then
        If IsReferenceValid(refScaleSelect) Then
            sFormula = Replace(sFormula, "ref", Range(refScaleSelect.Value).Cells(1, 1).Address(True, False))
        Else
            Exit Function
        End If
    End If
    
    'Check if multiplier box is filled in
    On Error GoTo InvalidMultiplier
    
    dblMultiplier = CDbl(txtMultiplier.Value)
    
    sFormula = Replace(sFormula, "scale", CStr(dblMultiplier))
    
    dblResult = Evaluate(sFormula)
    'MsgBox Prompt:="Formula result is " & Evaluate(sFormula), Title:="Debug custom formula"
    On Error GoTo 0
    
    IsFormulaValid = True
    
    Exit Function
    
CannotEvaluate:

    'Messaging handled in main form check sub
    
    Exit Function
    
InvalidMultiplier:
    
    MsgBox "Multiplier value in selected stat cannot be read as a number"
    
    Exit Function
    
End Function
'Check Scale reference cell
Private Function IsReferenceValid(ByRef ref As Control) As Boolean
    Dim rngTest As Range

    On Error Resume Next
    Set rngTest = Range(ref.Value)
    Set rngTest = Nothing
    If Err > 0 Then
        GoTo InvalidReference
    End If
    On Error GoTo 0
    
    'Assign address to start of stat value column
    ref.Value = Cells(Range(ref.Value).Row, 4).Address(True, False)
    
    IsReferenceValid = True
    
    Exit Function
    
InvalidReference:
    
    MsgBox ("Check your Scaler Cell Reference, not a valid reference")
    Exit Function
    
End Function
'Toggle Formula textbox
Private Sub cboRefOperator_Change()
    Dim bool As Boolean
    
    Select Case cboRefOperator.Value
        Case "Formula"      'Show formula box and text
            bool = True
            tglOrder.Caption = "Decided by custom formula"
            lblGenericMult.Caption = "Multiplier"
        Case "No scaling"   'Hide formula and change operator toggle caption
            bool = False
            tglOrder.Caption = "No order needed"
            lblGenericMult.Caption = "Start Value"
        Case Else           'Operator to be used
            bool = False
            lblGenericMult.Caption = "Multiplier"
            Call tglOrder_Click
    End Select
    
    lblFormula.Visible = bool
    txtFormula.Visible = bool
    txtFormula.Enabled = bool
    lblFormulaCheck.Visible = bool
    
End Sub
'After typing formula, check if valid
Private Sub txtFormula_Exit(ByVal Cancel As MSForms.ReturnBoolean)

    With lblFormulaCheck
        If Len(txtFormula) > 0 Then
            .Font.Bold = True
            If IsFormulaValid Then
                .Caption = "Valid result"
            Else
                .Caption = "Not valid"
            End If
        Else
            .Font.Bold = False
            .Caption = "Please type in your calculation formula"
        End If
    End With
    
End Sub
'Populate Operator options
Private Sub PopulateOperatorComboBox()
    
    With cboRefOperator
        .AddItem ("*")
        .AddItem ("/")
        .AddItem ("+")
        .AddItem ("Formula")
        .AddItem ("No scaling")
    End With
    
End Sub
'Populate Class listbox
Private Sub PopulateClassList()
    Dim c As Range
    
    For Each c In Worksheets("Enumerations").ListObjects("tblCharacterClasses").DataBodyRange.Columns(1).Cells
        With lboClass
            .AddItem (c.Value)
            .Selected(.ListCount - 1) = True
        End With
    Next c
    
    Set c = Nothing

End Sub
'Prevent common invalid characters in stat name
Private Sub txtStatName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr("/\?*[],.{}!@#$%^&-=+<>()~`;:'""", Chr(KeyAscii)) Then KeyAscii = 0
End Sub
'Toggle any textbox enable based on associated checkbox
Private Sub ToggleTextBox(ByVal bool As Boolean, ByRef txt As Control)
    
    With txt
        If bool Then
            .Enabled = True
            .BackColor = &H80000005
        Else
            .Enabled = False
            .BackColor = &H80000016
        End If
    End With
    
End Sub
'Toggle Base Multiplier textbox on Generic Stat page
Private Sub chkAddToBase_Click()

    Call ToggleTextBox(chkAddToBase, txtBaseMultiplier)

End Sub
'Toggle Starting value textbox
Private Sub chkHasStartValue_Click()
    
    Call ToggleTextBox(chkHasStartValue, txtStartValue)
            
End Sub
'Toggle Fixed level up increase textbox
Private Sub chkFixedLevelIncrease_Click()
    
    Call ToggleTextBox(chkFixedLevelIncrease, txtFixedLevelIncrease)

End Sub
'Toggle math order
Private Sub tglOrder_Click()
    If cboRefOperator.Value <> "No scaling" And cboRefOperator.Value <> "Formula" Then
        If tglOrder Then
            tglOrder.Caption = "Reference ? Scaler"
        Else
            tglOrder.Caption = "Scaler ? Reference"
        End If
    End If
End Sub
'Toggle selecting all items
Private Sub chkSelectAll_Click()
    Dim i As Integer
    Dim bool As Boolean
    
    bool = chkSelectAll
    
    With lboClass
        For i = 0 To .ListCount - 1
            .Selected(i) = bool
        Next i
    End With
End Sub
'Cancel form
Private Sub btnCancel_Click()
    Me.Hide
End Sub
