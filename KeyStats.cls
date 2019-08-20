VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet2"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub btnAddEnemy_Click()
    Load frmAddEnemy
    frmAddEnemy.StartUpPosition = 0
    frmAddEnemy.Left = Application.Left + (0.5 * Application.Width) - (0.5 * frmAddEnemy.Width)
    frmAddEnemy.Top = Application.Top + (0.5 * Application.Height) - (0.5 * frmAddEnemy.Height)
    frmAddEnemy.Show
    Unload frmAddEnemy
End Sub

Private Sub btnAddStat_Click()
    Load frmAddStat
    frmAddStat.StartUpPosition = 0
    frmAddStat.Left = Application.Left + (0.5 * Application.Width) - (0.5 * frmAddStat.Width)
    frmAddStat.Top = Application.Top + (0.5 * Application.Height) - (0.5 * frmAddStat.Height)
    frmAddStat.Show
    Unload frmAddStat
End Sub

Private Sub btnWriteToFile_Click()
    Call WriteStatsToFile
End Sub
