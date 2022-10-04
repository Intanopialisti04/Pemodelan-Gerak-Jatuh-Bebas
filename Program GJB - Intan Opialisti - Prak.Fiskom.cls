VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "Sheet1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = True
Private Sub Start_Click()
 Range("B16").Value = 0 '0
 delta_t = Range("D12").Value '0,1
  While Range("B16").Value < 9
  Range("B16").Value = Range("B16").Value + delta_t
  DoEvents
 Wend
End Sub

