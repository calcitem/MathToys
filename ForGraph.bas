Attribute VB_Name = "ForGraph"
Public Sub DrawGrp()
Call Pic.Command1_Click
  If Pic.fclb.Checked = True Then
    Pic.cshsh.Checked = False
    Pic.ExplicitFun.Checked = True
    prmtfct.Hide
    Call FctList.DrawList_Click
  Else
    If Pic.cshsh.Checked = True Then
      prmtfct.Show
      Call prmtfct.Draw_Click
    Else
      Call Pic.Command3_Click
    End If
  End If
End Sub
'Call drawgrp
