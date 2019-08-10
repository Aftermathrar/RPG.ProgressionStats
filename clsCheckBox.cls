VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "clsCheckBox"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
Option Explicit
'Adapted from http://www.vbaexpress.com/forum/showthread.php?56930
 
Public WithEvents CheckBox As MSForms.CheckBox
Attribute CheckBox.VB_VarHelpID = -1
Dim myColl As Collection
 
Private Sub CheckBox_Click()
    On Error Resume Next
    CallByName chkParent, CheckBox.Parent.Name & "CheckBox_Click", VbMethod
    On Error GoTo 0
End Sub
 
Private Function chkParent() As Object
   
    Set chkParent = CheckBox
    On Error Resume Next
    Do
        Set chkParent = chkParent.Parent
    Loop Until Err
    On Error GoTo 0
   
End Function
'Find out what this is for
Private Sub Class_Initialize()
    Set myColl = New Collection
    myColl.Add Item:=Me
End Sub
