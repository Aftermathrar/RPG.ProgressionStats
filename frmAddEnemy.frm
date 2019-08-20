VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} frmAddEnemy 
   Caption         =   "Add Enemy Type"
   ClientHeight    =   2385
   ClientLeft      =   45
   ClientTop       =   390
   ClientWidth     =   3945
   OleObjectBlob   =   "frmAddEnemy.frx":0000
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "frmAddEnemy"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
Option Explicit

'Initialize UserForm with existing enemy types and stat options
Private Sub UserForm_Initialize()

    Call PopulateComboBox           'Adds all enemy types to combobox menu
    Call AddStatCheckBoxes          'Adds all enemy stat types
    
End Sub
'Create new enemy type
Private Sub btnSubmit_Click()
    Dim iStats As Integer
    
    'Calc number of stats selected
    iStats = NumberOfStats
    
    If Not IsFormFilledOut(iStats) Then Exit Sub
    
    Call AddToEnumerations(iStats)
    Call AddToEnemyTable
    Call ExtendEnemiesNamedRange
    
    'Scary
    Call OverwriteClassEnumerationsFile
    Call WriteStatsToFile
    
    Me.Hide
End Sub
'Check to make sure required fields are completed
Private Function IsFormFilledOut(ByVal iStats As Integer) As Boolean
    
    'Enemy name empty
    If txtEnemyName = Empty Then
        MsgBox Prompt:="Please type in an enemy name.", Buttons:=vbOKOnly, Title:="No enemy name entered"
        txtEnemyName.SetFocus
        Exit Function
    End If
    
    'Enemy name matches existing name
    If Not Worksheets("Enumerations").ListObjects("tblCharacterClasses").DataBodyRange.Columns(1).Find(txtEnemyName.Text) Is Nothing Then
        MsgBox Prompt:="Your enemy name matches an already created Class." & vbNewLine _
            & "Please set a unique name or use the Edit Enemy form instead.", _
            Title:="Duplicate enemy class"
        With txtEnemyName
            .SetFocus
            .SelStart = 0
            .SelLength = Len(txtEnemyName)
        End With
        Exit Function
    End If
    If iStats = 0 Then
        MsgBox Prompt:="No stats are selected, please choose some stats!", Title:="No stats selected"
        Exit Function
    End If
    
    IsFormFilledOut = True
End Function
'Add enemy type to character class table
Private Sub AddToEnumerations(ByVal iStats As Integer)
    Dim iRows As Integer
    Dim tblClass As ListObject
    Dim rngClass As Range
    
    Set tblClass = Worksheets("Enumerations").ListObjects("tblCharacterClasses")
    tblClass.ListRows.Add
    
    Set rngClass = tblClass.DataBodyRange
    iRows = rngClass.Rows.Count
    
    rngClass(iRows, 1).Value = txtEnemyName
    rngClass(iRows, 2).Value = rngClass(iRows - 1, 2).Value + 1
    rngClass(iRows, 3).Value = iStats
    
    Set tblClass = Nothing
    Set rngClass = Nothing
    
End Sub
'Add enemy name, multiplier, and stats to enemy table
Private Sub AddToEnemyTable()
    Dim ctl As Control
    Dim rngEnemy As Range
    Dim dblMultValue As Single
    Dim sTextBoxName As String
    
    'Set start of enemy stat data 3 rows from last set
    Set rngEnemy = Worksheets("Key Stats").Cells(Rows.Count, 1).End(xlUp).Offset(3, 0)
    'Set Enemy name format
    With rngEnemy.Offset(-1, 0)
        .Value = txtEnemyName
        .Font.Size = 12
        .Font.Bold = True
    End With
    If Len(txtComment) > 0 Then     'Add enemy comment if there is one
        rngEnemy.Offset(-1, 0).AddComment (txtComment)
    End If
    
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "CheckBox"
                If ctl Then
                    'Get associated textbox name
                    sTextBoxName = BuildControlName("txt", ctl.Name)
                    
                    'Stat Name
                    rngEnemy.Value = ctl.Caption
                    
                    'Multiplier value
                    dblMultValue = CDbl(Me.Controls(sTextBoxName).Text)
                    rngEnemy.Offset(0, 1).Value2 = Application.WorksheetFunction.Round(dblMultValue, 2)
                    
                    'Formula for calculating stats
                    rngEnemy.Offset(0, 3).Formula = "=" & rngEnemy.Offset(0, 1).Address & "*" & GetStatMultiplierAddress(ctl.Caption)
                    rngEnemy.Offset(0, 3).Resize(1, GetLevelRange).FillRight   'Filler formula across
                    Set rngEnemy = rngEnemy.Offset(1, 0)
                End If
        End Select
    Next ctl
    
    Set ctl = Nothing
    Set rngEnemy = Nothing
                    
End Sub
'Get number of levels configured
Private Function GetLevelRange() As Integer
    GetLevelRange = Worksheets("Key Stats").Cells(3, Worksheets("Key Stats").Columns.Count).End(xlToLeft).Value
End Function
'Get Stat multiplier address
Private Function GetStatMultiplierAddress(ByVal sName As String) As String
    Dim tblStats As ListObject
    Dim rngTemp As Range
    
    Set tblStats = Worksheets("Enumerations").ListObjects("tblStats")
    'Check if filter is already applied
    If tblStats.Parent.AutoFilterMode Then
        tblStats.AutoFilter.ShowAllData
    End If
    
    tblStats.Range.AutoFilter field:=1, Criteria1:=sName
    Set rngTemp = tblStats.AutoFilter.Range.Offset(1, 0).SpecialCells(xlCellTypeVisible)
    GetStatMultiplierAddress = rngTemp(1, 4).Value
    
    tblStats.AutoFilter.ShowAllData
    
    Set tblStats = Nothing
    Set rngTemp = Nothing
    
End Function
'Get number of stats to be created for enemy
Private Function NumberOfStats() As Integer
    Dim ctl As Control
    Dim i As Integer
    
    i = 0
    
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "CheckBox"
                If ctl Then
                    i = i + 1
                End If
        End Select
    Next ctl
    
    NumberOfStats = i
    
    Set ctl = Nothing
    
End Function
'Fill ComboBox with existing enemy types
Private Sub PopulateComboBox()
    Dim rng As Range
    Dim cEnemy As Range
    
    Set rng = Worksheets("Enumerations").ListObjects("tblCharacterClasses").DataBodyRange.Columns(1)
    
    cboEnemyVariant.AddItem (cboEnemyVariant.Text)
    
    For Each cEnemy In rng.Cells
        If LCase(cEnemy.Value) <> "player" Then
            cboEnemyVariant.AddItem (cEnemy.Value)
        End If
    Next cEnemy
    
    Set rng = Nothing
    Set cEnemy = Nothing
    
End Sub
'Loop through stats and generate checkboxes and textboxes for each
Private Sub AddStatCheckBoxes()
    Dim newChk As clsCheckBox
    Dim cStat As Range
    Dim ctl As Control
    Dim iCountOffset As Integer
    
    iCountOffset = 0
    
    For Each cStat In Worksheets("Enumerations").ListObjects("tblStats").DataBodyRange.Columns(3).Cells
        If cStat <> Empty Then
            Set ctl = Me.Controls.Add("Forms.CheckBox.1", "chk" & CStr(cStat.Value))
            
            'Set as checkbox class for clicky stuff later
            Set newChk = New clsCheckBox
            Set newChk.CheckBox = ctl
            
            With ctl
                .Caption = cStat.Offset(0, -2).Value
                .Top = 84 + 15 * iCountOffset
                .Left = 6
                .Width = 90
                If Len(.Caption) > 17 Then
                    .Height = .Height * 2
                    iCountOffset = iCountOffset + 1
                End If
            End With
            
            'Add textbox
            Set ctl = Me.Controls.Add("Forms.TextBox.1", "txt" & CStr(ctl.Name))
            
            With ctl
                .Top = 84 + 15 * iCountOffset
                .Left = 106
                .Width = 40
                .Enabled = False
                .BackColor = &H80000016
            End With
            
            'Resize form
            Me.Height = Me.Height + 15
            
            iCountOffset = iCountOffset + 1
        End If
        
    Next cStat
    
    'Move buttons to bottom
    btnSubmit.Top = Me.Height - 51
    btnCancel.Top = Me.Height - 51
    
    Set cStat = Nothing
    Set ctl = Nothing
    
End Sub
'Prevent common invalid characters in enemy name
Private Sub txtEnemyName_KeyPress(ByVal KeyAscii As MSForms.ReturnInteger)
    If InStr("/\?*[],.{}!@#$%^&-=+<>()~`;:'""", Chr(KeyAscii)) Then KeyAscii = 0
End Sub
'Checkbox control
Public Sub frmAddEnemyCheckBox_Click()
    Dim sCtlName As String
    
    sCtlName = BuildControlName("txt", frmAddEnemy.ActiveControl.Name)
    
    Call EnableDisableText(sCtlName, frmAddEnemy.ActiveControl, 1)
    
    'MsgBox frmAddEnemy.ActiveControl.Name & " Top: " & frmAddEnemy.ActiveControl.Top
End Sub
'Enable or disable text and set values
'Takes name of text control and default value if setting enabled
Private Sub EnableDisableText(ByVal sCtlName As String, ByVal bool As Boolean, ByVal multValue As Single)

    If bool Then
        With Me.Controls(sCtlName)
            .Enabled = True
            .BackColor = &H80000005
            .Text = multValue
        End With
    Else
        With Me.Controls(sCtlName)
            .Enabled = False
            .BackColor = &H80000016
            .Text = vbNullString
        End With
    End If
    
End Sub
'Build control name
Private Function BuildControlName(ByVal sType As String, ByVal ctlName As String)
    BuildControlName = sType & ctlName
End Function
'Update stat checkboxes
Private Sub cboEnemyVariant_Change()
    Dim rngEnemy As Range
    Dim rngStatName As Range
    Dim sCtlName As String
    Dim ctl As Control
    
    'Reset fields
    For Each ctl In Me.Controls
        Select Case TypeName(ctl)
            Case "CheckBox"
                If ctl Then
                    ctl.Value = False
                    Call EnableDisableText(BuildControlName("txt", ctl.Name), ctl, 1)
                End If
        End Select
    Next ctl
    
    Set rngEnemy = Range("Enemies").Find(What:=cboEnemyVariant.Text, LookIn:=xlValues, LookAt:=xlWhole)
    If Not rngEnemy Is Nothing Then
        'Get stat name range to lookup control names
        Set rngStatName = Worksheets("Enumerations").ListObjects("tblStats").DataBodyRange.Columns(1)
        'Offset rngEnemy to the start of the stat body
        Set rngEnemy = rngEnemy.Offset(1, 0)
        
        'Go through each stat, match up the form control name, enable it and set multiplier value
        Do While rngEnemy <> Empty
            sCtlName = rngStatName.Find(rngEnemy.Value).Offset(0, 2).Value
            Set ctl = Me.Controls(BuildControlName("chk", sCtlName))
            ctl = True
            Call EnableDisableText(BuildControlName("txt", ctl.Name), ctl, rngEnemy.Offset(0, 1).Value)
            Set rngEnemy = rngEnemy.Offset(1, 0)
        Loop
    End If
    
    Set rngEnemy = Nothing
    Set rngStatName = Nothing
    Set ctl = Nothing
    
End Sub

Private Sub btnCancel_Click()
    Me.Hide
End Sub
