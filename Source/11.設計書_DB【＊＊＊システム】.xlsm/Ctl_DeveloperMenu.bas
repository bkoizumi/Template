Attribute VB_Name = "Ctl_DeveloperMenu"
Option Explicit

Function Reset()

  Call Library.startScript
  Sheets("<CopyTable>").Rows("14:70").copy
  ActiveSheet.Range("A14:A15").Select
  ActiveSheet.Paste

  Call Library.endScript

End Function

