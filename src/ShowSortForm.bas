Attribute VB_Name = "ShowSortForm"
' # ------------------------------------------------------------------------------
' # Name:        ShowSortForm.bas
' # Purpose:     Worker functions for "Budget Backup Manager" Excel VBA Add-In
' #
' # Author:      Brian Skinn
' #                bskinn@alum.mit.edu
' #
' # Created:     13 Jan 2015
' # Copyright:   (c) Brian Skinn 2015-2018
' # License:     The MIT License; see "LICENSE.txt" for full license terms
' #                   and contributor agreement.
' #
' #       http://www.github.com/bskinn/excel-budgetbackup
' #
' # ------------------------------------------------------------------------------

Option Explicit

Public Sub showBackupForm()
Attribute showBackupForm.VB_ProcData.VB_Invoke_Func = "J\n14"
    FrmBackupSort.Show
End Sub

Public Sub clearAddinCustDocProps()
    ' Helper to clear the add-in custom doc props
    FrmBackupSort.clearAddinCustDocProps
    ThisWorkbook.Save
    Unload FrmBackupSort
    
End Sub

Public Sub setAddinNameProp()
    ' Helper to update the 'Name' built-in docprop
    ThisWorkbook.BuiltinDocumentProperties(1) = "Budget Backup Manager v2.0"
End Sub
