Attribute VB_Name = "ShowSortForm"
' # ------------------------------------------------------------------------------
' # Name:        ShowSortForm.bas
' # Purpose:     Worker functions for "Budget Backup Manager" Excel VBA Add-In
' #
' # Author:      Brian Skinn
' #                bskinn@alum.mit.edu
' #
' # Created:     13 Jan 2015
' # Copyright:   (c) Brian Skinn 2017
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

Public Sub doClearReaderLoc()
    ' Helper function to clear the stored Reader location before distribution
    '  of an .xlam addin.
    ' Ultimately obsolete; will be removed once the Reader search functionality
    '  is culled.
    FrmBackupSort.clearReaderLocation
End Sub
