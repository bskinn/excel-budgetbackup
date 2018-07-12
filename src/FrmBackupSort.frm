VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmBackupSort 
   Caption         =   "Sort Budget Items"
   ClientHeight    =   8400
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7125
   OleObjectBlob   =   "FrmBackupSort.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmBackupSort"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False

' # ------------------------------------------------------------------------------
' # Name:        FrmBackupSort.bas
' # Purpose:     Core form for "Budget Backup Manager" Excel VBA Add-In
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

' General objects
Dim fs As Scripting.FileSystemObject, fld As Scripting.Folder, fl As Scripting.File
Dim wsf As WorksheetFunction

' Regexes
' General filename, checks that starts with valid parenned index, incl or excl
Dim rxFNIdxValid As New VBScript_RegExp_55.RegExp

' Detailed formating match of 'included' file
Dim rxInclFnameDetail As New VBScript_RegExp_55.RegExp

' Detailed formatting match of 'excluded' file
Dim rxExclFnameDetail As New VBScript_RegExp_55.RegExp

' Detailed formatting match of 'included' or 'excluded' file
Dim rxFnameDetail As New VBScript_RegExp_55.RegExp

' Global hash tracker variable
Dim hash As Long

' Global collisions-or-no tracker variable
Dim anyCollisions As Boolean

' Global 'matching hash?' tracker
Dim anyHashMismatch As Boolean

' Global 'last attempt found locked file' flag
Dim lockedFileFound As Boolean

' Global 'critical locked file' flag
Dim criticalLockedFile As Boolean


Const NONE_FOUND As String = "<none found>"
Const EMPTY_LIST As String = "<empty>"
Const NUM_FORMAT As String = "00"
Const CANCEL_RETURN As String = "!!CANCELED!!"
Const UNINIT_RETURN As String = "!!FOLDER NOT INITIALIZED!!"""
Const DEF_PADWIDTH As Long = 2
Const CDP_PADWIDTH As String = "PadWidth"

Dim NL As String    ' To contain Newline




Private Sub setCtrls()
    ' Helper function to call all individual control-setter functions
    setFldCtrls
    setInclCtrls
    setExclCtrls
    
End Sub

Private Sub setFldCtrls()
    ' Helper function for setting the folder-related buttons
    
    Dim fldIsBound As Boolean
    
    fldIsBound = (Not fld Is Nothing)
    If fldIsBound Then fldIsBound = fs.FolderExists(fld.Path)
    
    BtnOpen.Enabled = True  ' always enabled here
    
    BtnReload.Enabled = fldIsBound
    BtnShowFolder.Enabled = fldIsBound
    
End Sub

Private Sub setInclCtrls()
    ' Helper for setting the 'included' list buttons, and related
    ' (includes the sheet generation button)
    
    Dim anyIncls As Boolean
    
    anyIncls = (LBxIncl.ListCount >= 1) And (LBxIncl.List(0, 0) <> NONE_FOUND) _
                And (LBxIncl.List(0, 0) <> EMPTY_LIST)
    
    LBxIncl.Enabled = anyIncls
    
    BtnOpenIncl.Enabled = anyIncls And (Not anyHashMismatch) And (Not criticalLockedFile)
    
    BtnMoveUp.Enabled = anyIncls And (Not anyCollisions) And (Not anyHashMismatch) And (Not criticalLockedFile)
    BtnMoveDown.Enabled = anyIncls And (Not anyCollisions) And (Not anyHashMismatch) And (Not criticalLockedFile)
    BtnMoveAfter.Enabled = anyIncls And (Not anyCollisions) And (Not anyHashMismatch) And (Not criticalLockedFile)
    BtnRemove.Enabled = anyIncls And (Not anyCollisions) And (Not anyHashMismatch) And (Not criticalLockedFile)
    BtnRemoveAll.Enabled = anyIncls And (Not anyCollisions) And (Not anyHashMismatch) And (Not criticalLockedFile)
    
    BtnGenSheet.Enabled = anyIncls And (Not anyHashMismatch) And (Not criticalLockedFile)
    
End Sub

Private Sub setExclCtrls()
    ' Helper for setting the 'excluded' list buttons
    
    Dim anyExcls As Boolean
    
    anyExcls = (LBxExcl.ListCount >= 1) And (LBxExcl.List(0, 0) <> NONE_FOUND) _
                And (LBxExcl.List(0, 0) <> EMPTY_LIST)

    LBxExcl.Enabled = anyExcls
    BtnOpenExcl.Enabled = anyExcls And (Not anyHashMismatch) And (Not criticalLockedFile)
    
    BtnAppend.Enabled = anyExcls And (Not anyCollisions) And (Not anyHashMismatch) And (Not criticalLockedFile)
    BtnAppendAll.Enabled = anyExcls And (Not anyCollisions) And (Not anyHashMismatch) And (Not criticalLockedFile)
    BtnInsert.Enabled = anyExcls And (Not anyCollisions) And (Not anyHashMismatch) And (Not criticalLockedFile)
    BtnInsertAll.Enabled = anyExcls And (Not anyCollisions) And (Not anyHashMismatch) And (Not criticalLockedFile)
    
End Sub




Private Sub popLists(Optional internalCall As Boolean = False)
    ' Clear and repopulate the included/excluded items LBxes
    '
    ' internalCall should be False for all calls to popLists
    ' outside of the popLists function itself.
    ' popLists uses internalCall = True for repeat internal
    ' calls, in cases such as where packNums actually results
    ' in a change to the folder contents.
    
    Dim mch As VBScript_RegExp_55.Match
    Dim inclFileCount As Long
    Dim calcMinPadWidth As Long
    Dim didPack As Boolean
    
    Static inclIdx As Long, exclIdx As Long
    Static inclView As Long, exclView As Long

    ' Store current selection/view indices if this is an external call,
    ' for restore after list repopulation.
    If Not internalCall Then
        exclIdx = LBxExcl.ListIndex
        inclIdx = LBxIncl.ListIndex
        
        exclView = LBxExcl.TopIndex
        inclView = LBxIncl.TopIndex
    End If
    
    ' Clear list contents in prep for repopulating
    LBxExcl.Clear
    LBxIncl.Clear
    
    If fld Is Nothing Then  ' there's nothing to populate
        ' So, set Empty
        LBxExcl.AddItem EMPTY_LIST
        LBxIncl.AddItem EMPTY_LIST
        
        ' ... and by definition this can't be an internal call,
        ' so *do* go to the final exit code.
        GoTo Final_Exit
    End If
    
    ' Count all the included files and update the padding setting
    ' if it's too small
    inclFileCount = 0
    For Each fl In fld.Files
        If rxInclFnameDetail.Test(fl.Name) Then
            inclFileCount = inclFileCount + 1
        End If
    Next fl
    
    calcMinPadWidth = CLng(wsf.Floor(wsf.Log10(wsf.Max(1, inclFileCount)), 1)) + 1
    If CLng(LblPadWidth.Caption) < calcMinPadWidth Then
        setPadWidth calcMinPadWidth
    End If
    
    ' Pad all sequence numbers in filenames
    padNums
    
    ' Cancel progress if critical file was locked
    If criticalLockedFile Then
        setCtrls
        Exit Sub
    End If
    
    ' Iterate through the files in the folder and sort to
    ' include/exclude lists as relevant.
    ' Ignores any files with names not matching the rigorous format
    For Each fl In fld.Files
        If rxInclFnameDetail.Test(fl.Name) Then
            LBxIncl.AddItem fl.Name
        End If
        If rxExclFnameDetail.Test(fl.Name) Then
            LBxExcl.AddItem fl.Name
        End If
    Next fl
    
    ' Attempt number packing; if it was done, then re-invoke
    ' popLists as internal. BUT, reset from controls and
    ' drop from Sub if critical file was locked
    didPack = packNums
    If criticalLockedFile Then
        setCtrls
        Exit Sub
    End If
    If didPack Then Call popLists(internalCall:=True)
    
    ' Only run the finalizing code for the outermost, external
    ' call of the function
    If Not internalCall Then GoTo Final_Exit
    
    ' Exit if this is an internal call
    Exit Sub

Final_Exit:
    ' Indicate empty include/excluded lists if detected
    If LBxExcl.ListCount < 1 Then
        LBxExcl.AddItem NONE_FOUND
        LBxExcl.ListIndex = 0
    End If
    
    If LBxIncl.ListCount < 1 Then
        LBxIncl.AddItem NONE_FOUND
        LBxIncl.ListIndex = 0
    End If
    
    ' Restore selections and views
    ' The .Min calls avoid index overflows when the size
    ' of a given list shrinks
    LBxExcl.ListIndex = wsf.Min(exclIdx, LBxExcl.ListCount - 1)
    LBxIncl.ListIndex = wsf.Min(inclIdx, LBxIncl.ListCount - 1)
    LBxExcl.TopIndex = exclView
    LBxIncl.TopIndex = inclView

End Sub






Private Sub padNums()
    ' Scan the files in the working folder and reformat any
    ' numerical values with zero-padding
    '
    ' Pads based on the current setting of LblPadWidth
    
    Dim mch As VBScript_RegExp_55.Match
    Dim padWidth As Long
    Dim idxWidth As Long
    Dim oldID As String
    Dim newID As String
    
    padWidth = CLng(getCustDocProp(CDP_PADWIDTH))
    
    For Each fl In fld.Files
        ' Only work with 'included' files
        If rxInclFnameDetail.Test(fl.Name) Then
            ' But store the match from the generic template matcher
            Set mch = rxFNIdxValid.Execute(fl.Name)(0)
            ' I think .SubMatches(0) is the inner item that has the '+' applied to it,
            '  while .SM(1)is the full numerical match.  .SM(2) is the remainder of the
            '  filename.
            
            ' Store old and calc new IDs
            oldID = mch.SubMatches(1)
            newID = Format(mch.SubMatches(1), String(padWidth, "0"))
            
            ' Only rename if they're different
            If oldID <> newID Then
                ' But confirm editability first
                If Not isFileEditable(fl) Then
                    ' This is a critical problem
                    notifyCriticalLockedFile fl.Name
                    criticalLockedFile = True
                    Exit Sub
                End If
                
                ' Go ahead and rename
                fl.Name = "(" & newID & ")" & mch.SubMatches(2)
            End If
        End If
    Next fl
End Sub

Private Function packNums() As Boolean
    ' Scan the 'included' listbox and repack all of the
    ' filename indexing so that there are no gaps, and no
    ' repeated indices.
    
    Dim workStr As String, iter As Long, fl As File
    Dim mch As VBScript_RegExp_55.Match
    Dim padFmt As String
    
    ' Helper format string for the padding
    ' Assumes the custom doc prop is tightly matched to LbxPadWidth
    padFmt = String(CLng(getCustDocProp(CDP_PADWIDTH).Value), "0")
    
    ' (Redundant) init of the return value
    packNums = False
    
    ' Only bother packing if there's something in LBxIncl
    If LBxIncl.ListCount > 0 Then
        ' ... and skip if no included items were found on last repop
        If LBxIncl.List(0, 0) <> NONE_FOUND Then
            ' Explicit iteration used here to track the index value
            ' that each file *should* hold.
            For iter = 0 To LBxIncl.ListCount - 1
                ' This 'included file' regex *should* always match
                Set mch = rxFNIdxValid.Execute(LBxIncl.List(iter, 0))(0)
                
                ' If the index doesn't match the iteration variable, fix it.
                If Not CLng(mch.SubMatches(1)) - 1 = iter Then
                    ' Store the file for reference
                    Set fl = fs.GetFile(fs.BuildPath(fld.Path, LBxIncl.List(iter, 0)))
                    
                    ' Check for editability
                    If Not isFileEditable(fl) Then
                        ' Critical problem, notify and exit
                        notifyCriticalLockedFile fl.Name
                        criticalLockedFile = True
                        Exit Function
                    End If
                    
                    ' Go ahead and rename, indicating that packing happened
                    packNums = True
                    fl.Name = "(" & Format(iter + 1, padFmt) & ")" & mch.SubMatches(2)
                End If
            Next iter
        End If
    End If
    
End Function

Private Function checkParenNames() As String
    ' Check if all filenames starting with a paren are valid,
    ' whether included or excluded
    '
    ' Returns special error return string if folder not set
    '
    ' Returns newline-separated list of invalid files, if any found
    '
    ' Returns empty string if all is ok.
    
    Dim fl As File
    
    checkParenNames = ""
    
    ' fld must be defined
    If fld Is Nothing Then
        checkParenNames = UNINIT_RETURN
        Exit Function
    End If
    
    ' Check all the files
    For Each fl In fld.Files
        ' If it starts with a paren...
        If Left(fl.Name, 1) = "(" Then
            ' ... then it needs to match either the 'include' or 'exclude' regex
            If Not rxFnameDetail.Test(fl.Name) Then
                checkParenNames = checkParenNames & fl.Name & NL
            End If
        End If
    Next fl
    
End Function

Private Function checkNameCollisions(Optional onlyValid As Boolean = True) As String
    ' Check if any *VALID INCLUDED/EXCLUDED* filenames are
    ' identical other than the "(...)" key
    '
    ' If onlyValid is True (default), then name collisions will only be checked
    ' for files whose names are properly formatted for sheet generation
    ' If False, then *ALL* files starting with '(##)' or '(x)' will be checked.
    '
    ' Returns error return string if folder not set
    '
    ' Returns newline-separated list of colliding files, if any found
    '
    ' Returns empty string if all is ok.
    
    Dim iter As Long, iter2 As Long
    Dim fl As File, fl2 As File
    Dim rxWork As RegExp
    Dim nStr As String, nStr2 As String
    Dim collStr As String, seenStr As String
    Dim outStr As String
    
    Const sep As String = "|"
    
    outStr = ""
    seenStr = sep
    
    ' fld must be defined
    If fld Is Nothing Then
        checkNameCollisions = UNINIT_RETURN
        Exit Function
    End If
    
    ' Choose the relevant regex
    If onlyValid Then
        Set rxWork = rxFnameDetail
    Else
        Set rxWork = rxFNIdxValid
    End If
    
    ' Crosscheck all the files
    For Each fl In fld.Files
        ' Only care here about collisions between valid-formatted files
        If rxWork.Test(fl.Name) Then
            ' Valid file; store everything past the key
            nStr = rxFNIdxValid.Execute(fl.Name)(0).SubMatches(2)
            
            For Each fl2 In fld.Files
                If rxWork.Test(fl2.Name) Then
                    ' This one's good also; store its name for checking
                    nStr2 = rxFNIdxValid.Execute(fl2.Name)(0).SubMatches(2)
                    
                    ' Examine only if they're not the same file, or if the file
                    ' hasn't already been seen as nStr/fl (vs nStr2/fl2)
                    If fl.Name <> fl2.Name And _
                            InStr(seenStr, sep & fl2.Name & sep) < 1 Then
                            
                        ' If the names match, store for complaint!
                        If nStr = nStr2 Then
                            ' Store the first filename as colliding for output, if new
                            If InStr(outStr, fl.Name & NL) < 1 Then
                                outStr = outStr & fl.Name & NL
                            End If
                            
                            ' Store the second filename as colliding for output, if new
                            If InStr(outStr, fl2.Name & NL) < 1 Then
                                outStr = outStr & fl2.Name & NL
                            End If
                            
                        End If
                    End If
                End If
            Next fl2
            
            ' Store as seen, to speed up the looping
            seenStr = seenStr & fl.Name & sep
        End If
    Next fl
    
    ' Transfer collection variable to function return value
    checkNameCollisions = outStr
    
End Function

Private Function doHashCheck() As Boolean
    ' Perform hash check and confirm whether it matches
    ' the global stored value. Set global flag accordingly
    
    ' If hash doesn't match, alert to need to reload
    anyHashMismatch = (hash <> hashFilenames)
    doHashCheck = Not anyHashMismatch
    
    If anyHashMismatch Then
        MsgBox "Folder contents have changed. Reload form to continue.", _
                vbOKOnly + vbExclamation, "Folder Contents Changed"
    End If
    
    setCtrls
    
End Function

Private Function hashFilenames() As Long
    ' Hashing function for aggregated file names and sizes
    ' Returns -1 if fld is not set
    
    Dim fl As File, iter As Long
    Const NAMEMULT As Long = 17
    Const SIZEMULT As Long = 37
    
    ' This is chosen based on the largest MULT, to avoid overflow
    Const modVal As Long = 54054000#
    
    ' Dummy exit if folder not set
    If fld Is Nothing Then
        hashFilenames = -1
        Exit Function
    End If
    
    ' For each file...
    For Each fl In fld.Files
        ' Only hash if it's a valid included or excluded file
        If rxFnameDetail.Test(fl.Name) Then
            ' Hash the name
            hashFilenames = (hashFilenames * NAMEMULT + hashName(fl.Name)) Mod modVal
    
            ' Hash the size
            hashFilenames = (hashFilenames * SIZEMULT + fl.Size) Mod modVal
        End If
    Next fl

End Function

Private Function hashName(nm As String) As Long
    ' Internal helper for hashing a filename
    
    Dim iter As Long
    
    For iter = 1 To Len(nm)
        hashName = hashName + iter * Asc(Mid(nm, iter, 1))
    Next iter
    
End Function



Private Sub proofParens()
    ' Wrapper function - Check for any suspect starts-with-paren files
    '
    ' Pops a messagebox if suspect things are found.
    '
    ' No specific action needs to be taken by the caller if any
    ' suspect files ARE found, as the suspect files will not be
    ' populated into the listboxes.
    
    Dim workStr As String
    
    ' Proof the files in the folder and report any problems
    workStr = checkParenNames
    
    If workStr <> "" Then
        MsgBox "The following files in the selected folder " _
                & "appear to be improperly formatted budget items:" _
                & NL & NL & workStr, vbOKOnly + vbExclamation, _
                "Possible Malformed Filenames"
    End If
    
End Sub

Private Sub proofCollisions()
    ' Wrapper function - Check for any filename collisions in the selected folder.
    '
    ' Pop a msgbox if any found, and update the global status flag accordingly, either way
    
    Dim workStr As String
    
    ' Proof for collisions
    workStr = checkNameCollisions
    
    If workStr <> "" Then
        MsgBox "The following files in the selected folder " & _
                "have name collisions; file manipulation will be disabled:" & _
                NL & NL & workStr, _
                vbOKOnly + vbCritical, _
                "Name Collisions Detected"
        anyCollisions = True
    Else
        anyCollisions = False
    End If
    
End Sub

Private Function minPadWidth() As Long
    ' Helper function to evaluate the minimum allowed
    ' pad width based on the number of items in the
    ' 'included' list
    
    minPadWidth = CLng(wsf.Floor(wsf.Log10(wsf.Max(1, LBxIncl.ListCount)), 1)) + 1
    
End Function

Private Sub setPadWidth(width As Long)
    ' Helper to set the pad width, set the docprop,
    ' save the addin workbook
    With LblPadWidth
        .Caption = width
        setCustDocProp CDP_PADWIDTH, CStr(.Caption)
    End With
    
End Sub

Private Function isFileEditable(fl As File) As Boolean
    ' Helper function to check editable status of a file.
    '
    ' Basically checks to see if it can be opened for appending.
    ' This should be nondestructive, while also giving information
    ' on editability permissions
    
    Dim errNum As Long
    
    On Error Resume Next
        fl.OpenAsTextStream(ForAppending).Close
    errNum = Err.Number: Err.Clear: On Error GoTo 0
    
    isFileEditable = (errNum = 0)
    
End Function

Private Sub notifyLockedFile(fname As String)
    ' Helper method to raise an alert box when
    ' a file is locked.
    
    MsgBox "The following file is locked by another application:" & NL & NL & _
            fname & NL & NL & "Close the file and retry.", _
            vbOKOnly + vbCritical, "File Locked"
    
End Sub

Private Sub notifyCriticalLockedFile(fname As String)
    ' Helper method for locked file preventing further form
    ' interaction
    
    MsgBox "The following file is locked by another application:" & NL & NL & _
            fname & NL & NL & "The file must be closed to continue.", _
            vbOKOnly + vbCritical, "File Locked"
    
End Sub


Private Sub BtnAppend_Click()
    ' Append the selected item from the Exclude list
    ' to the end of the Include list
    
    Dim mch As VBScript_RegExp_55.Match
    
    ' No folder selected, so exit with no action
    ' TO BE OBSOLETED BY CONTROL ACTIVATION/INACTIVATION
    If fld Is Nothing Then Exit Sub
    
    ' Excluded list is empty, do nothing
    ' TO BE OBSOLETED BY CONTROL ACTIVATION/INACTIVATION
    If LBxExcl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    ' No Excluded list item is selected; do nothing
    If LBxExcl.ListIndex < 0 Then Exit Sub
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Reset the locked file found flag
    lockedFileFound = False
    
    ' Retrieve the filename and assign to File object;
    ' ASSUMES ALREADY VETTED AGAINST rxFNIdxValid
    Set mch = rxFNIdxValid.Execute(LBxExcl.List(LBxExcl.ListIndex, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    
    ' Check editable status
    If Not isFileEditable(fl) Then
        notifyLockedFile fl.Name
        lockedFileFound = True
        Exit Sub
    End If
    
    ' Rename the file to an 'included' form
    If LBxIncl.List(0, 0) = NONE_FOUND Then
        ' Start list from nothing; index is one
        fl.Name = "(" & Format(1, NUM_FORMAT) & ")" & mch.SubMatches(2)
    Else
        ' Assign the index for the end of the 'included' list
        fl.Name = "(" & Format(LBxIncl.ListCount + 1, NUM_FORMAT) & ")" & mch.SubMatches(2)
    End If
    
    
    ' Refresh listboxes
    popLists
    
    ' Can only get here if there was no hash problem beforehand,
    ' so just update the hash
    hash = hashFilenames
    
    ' Set the control states
    setCtrls
    
End Sub

Private Sub BtnAppendAll_Click()
    ' Iteratively append all 'excluded' items at the current position
    
    ' Folder must be defined
    If fld Is Nothing Then Exit Sub
    
    ' Excluded list has to have items
    If LBxExcl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Must ensure something is selected in Excl
    LBxExcl.ListIndex = CLng(wsf.Max(0, LBxExcl.ListIndex))
    
    ' Reset locked file found flag
    lockedFileFound = False
    
    ' Append everything, allowing events to occur after each loop
    ' Adding in FORWARD sequence because the Append dynamic causes
    ' each new item to be included BELOW the last item inserted.
    Do Until LBxExcl.List(0, 0) = NONE_FOUND Or lockedFileFound Or criticalLockedFile
        BtnAppend_Click
        DoEvents
    Loop
    
    ' Form refresh and hash updates are handled by BtnRemove,
    ' so no need to do either here.
End Sub

Private Sub BtnClose_Click()
    ' Jettison the form entirely
    
    ' Save the addin workbook, to store the custom doc props
    ThisWorkbook.Save
    
    ' Then unload!
    Unload FrmBackupSort
    
End Sub

Private Sub BtnGenSheet_Click()
    ' Create and populate a budget sheet based on
    ' the current contents of the 'include' list

    Dim genBk As Workbook, genSht As Worksheet
    Dim sht As Worksheet
    Dim workCel As Range, tblCel As Range
    Dim celS As Range, celE As Range, celM As Range, celC As Range, celT As Range
    Dim mchs As MatchCollection, mch As Match
    Dim fl As File
    Dim counts As Variant, anyFlsFound As Boolean
    Dim inlaids As Variant
    Dim iter As Long
    Const idxS As Long = 0
    Const idxE As Long = 1
    Const idxM As Long = 2
    Const idxC As Long = 3
    Const idxT As Long = 4
    Const smchNum As Long = 0
    Const smchType As Long = 1
    Const smchVend As Long = 2
    Const smchDesc As Long = 3
    Const smchCost As Long = 4
    Const smchQty As Long = 5
    Const costFmt As String = "$#,##0.00"
    
    ' Drop if folder is not selected
    If fld Is Nothing Then Exit Sub
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Scan the work folder for properly configured filenames
    counts = Array(0, 0, 0, 0, 0)
    inlaids = Array(0, 0, 0, 0, 0)
    anyFlsFound = False
    For Each fl In fld.Files
        ' Ignore any files not starting with a paren
        If Left(fl.Name, 1) = "(" Then
            ' If a file is found matching the 'included' filter, then count it
            If rxInclFnameDetail.Test(fl.Name) Then
                anyFlsFound = True
                Set mch = rxInclFnameDetail.Execute(fl.Name)(0)
                
                ' Increment the relevant category count
                Select Case UCase(mch.SubMatches(smchType))
                Case "S"
                    counts(idxS) = counts(idxS) + 1
                Case "E"
                    counts(idxE) = counts(idxE) + 1
                Case "M"
                    counts(idxM) = counts(idxM) + 1
                Case "C"
                    counts(idxC) = counts(idxC) + 1
                Case "T"
                    counts(idxT) = counts(idxT) + 1
                End Select
            End If
        End If
    Next fl
    
    ' If nothing found, warn and exit
    If Not anyFlsFound Then
        Call MsgBox("No properly formatted files were found. Exiting...", _
                vbOKOnly, "No formatted files")
        Exit Sub
    End If
    
    ' Create new workbook
    Set genBk = Workbooks.Add
    
    ' Strip down to a single worksheet if needed
    Application.DisplayAlerts = False
    Do Until genBk.Worksheets.Count < 2
        genBk.Worksheets(genBk.Worksheets.Count).Delete
    Loop
    Application.DisplayAlerts = True
    
    ' Bind the sheet
    Set genSht = genBk.Worksheets(1)
    
    ' Initialize the sheet structure
    ' Define the reference cells
    Set tblCel = genSht.Cells(3, 1)
    Set celS = tblCel.Offset(1, 0)
    Set celE = celS.Offset(counts(idxS) + 3, 0)
    Set celM = celE.Offset(counts(idxE) + 3, 0)
    Set celC = celM.Offset(counts(idxM) + 3, 0)
    Set celT = celC.Offset(counts(idxC) + 3, 0)
    
    ' Headers
    'Set workCel = tblCel
    tblCel.Formula = "Item No"
    tblCel.Offset(0, 1).Formula = "Description"
    tblCel.Offset(0, 2).Formula = "Vendor"
    tblCel.Offset(0, 3).Formula = "Unit Cost"
    tblCel.Offset(0, 4).Formula = "Qty"
    tblCel.Offset(0, 5).Formula = "Extended Cost"
    tblCel.Resize(1, 6).Font.Bold = True
    
    celS.Offset(0, 1) = "Services"
    celS.Offset(0, 1).Font.Bold = True
    
    celE.Offset(0, 1) = "Equipment"
    celE.Offset(0, 1).Font.Bold = True
    
    celM.Offset(0, 1) = "Materials"
    celM.Offset(0, 1).Font.Bold = True
    
    celC.Offset(0, 1) = "Chemicals"
    celC.Offset(0, 1).Font.Bold = True
    
    celT.Offset(0, 1) = "Travel"
    celT.Offset(0, 1).Font.Bold = True
    
    ' Loop over the files and, if rx.Test, insert
    For Each fl In fld.Files
        If rxInclFnameDetail.Test(fl.Name) Then
            Set mch = rxInclFnameDetail.Execute(fl.Name)(0)
            Select Case UCase(mch.SubMatches(smchType))
            Case "S"
                Set workCel = celS.Offset(1 + inlaids(idxS), 0)
                inlaids(idxS) = inlaids(idxS) + 1
            Case "E"
                Set workCel = celE.Offset(1 + inlaids(idxE), 0)
                inlaids(idxE) = inlaids(idxE) + 1
            Case "M"
                Set workCel = celM.Offset(1 + inlaids(idxM), 0)
                inlaids(idxM) = inlaids(idxM) + 1
            Case "C"
                Set workCel = celC.Offset(1 + inlaids(idxC), 0)
                inlaids(idxC) = inlaids(idxC) + 1
            Case "T"
                Set workCel = celT.Offset(1 + inlaids(idxT), 0)
                inlaids(idxT) = inlaids(idxT) + 1
            End Select
            
            workCel.Value = CLng(mch.SubMatches(smchNum))
            With workCel.Offset(0, 1)
                .NumberFormat = "@"
                .Formula = mch.SubMatches(smchDesc)
            End With
            With workCel.Offset(0, 2)
                .NumberFormat = "@"
                .Formula = mch.SubMatches(smchVend)
            End With
            With workCel.Offset(0, 3)
                .NumberFormat = costFmt
                .Value = CDbl(mch.SubMatches(smchCost))
            End With
            With workCel.Offset(0, 4)
                .NumberFormat = "@"
                .Value = CDbl(mch.SubMatches(smchQty))
            End With
            With workCel.Offset(0, 5)
                .NumberFormat = costFmt
                .Formula = "=" & .Offset(0, -1).Address(False, False) & _
                        "*" & .Offset(0, -2).Address(False, False)
            End With
            
            ' Alignment
            workCel.Offset(0, 3).Resize(1, 3).HorizontalAlignment = xlLeft
        End If
    Next fl
    
    ' Apply borders
    With genSht.UsedRange
        .Borders(xlEdgeBottom).Weight = xlThin
        .Borders(xlEdgeTop).Weight = xlThin
        .Borders(xlEdgeLeft).Weight = xlThin
        .Borders(xlEdgeRight).Weight = xlThin
        .Borders(xlInsideHorizontal).Weight = xlThin
        .Borders(xlInsideVertical).Weight = xlThin
    End With
    
    ' Summary fields
    celS.Offset(0, 6).Formula = "Services"
    If counts(idxS) > 0 Then
        celS.Offset(1, 6).Formula = "=SUM(" & _
                celS.Offset(1, 5).Resize(counts(idxS), 1).Address & ")"
    Else
        celS.Offset(1, 6).Formula = "0"
    End If
    
    
    celE.Offset(0, 6).Formula = "Equipment"
    If counts(idxE) > 0 Then
        celE.Offset(1, 6).Formula = "=SUM(" & _
                celE.Offset(1, 5).Resize(counts(idxE), 1).Address & ")"
    Else
        celE.Offset(1, 6).Formula = "0"
    End If
    
    
    celM.Offset(0, 6).Formula = "Materials"
    If counts(idxM) > 0 Then
        celM.Offset(1, 6).Formula = "=SUM(" & _
                celM.Offset(1, 5).Resize(counts(idxM), 1).Address & ")"
    Else
        celM.Offset(1, 6).Formula = "0"
    End If
    
    
    celC.Offset(0, 6).Formula = "Chemicals"
    If counts(idxC) > 0 Then
        celC.Offset(1, 6).Formula = "=SUM(" & _
                celC.Offset(1, 5).Resize(counts(idxC), 1).Address & ")"
    Else
        celC.Offset(1, 6).Formula = "0"
    End If
    
    
    celT.Offset(0, 6).Formula = "Travel"
    If counts(idxT) > 0 Then
        celT.Offset(1, 6).Formula = "=SUM(" & _
                celT.Offset(1, 5).Resize(counts(idxT), 1).Address & ")"
    Else
        celT.Offset(1, 6).Formula = "0"
    End If
    
    
    ' Change entire UsedRange to left-aligned
    genSht.UsedRange.Columns(genSht.UsedRange.Columns.Count) _
            .HorizontalAlignment = xlLeft
    
    ' 'Grand Total' field
    With tblCel.Offset(-2, 5)
        .Formula = "Total"
        .Font.Bold = True
        .Font.Size = 13
    End With
    With tblCel.Offset(-1, 5)
        .Formula = "=SUM(" & _
                celS.Offset(1, 6).Address & "," & _
                celE.Offset(1, 6).Address & "," & _
                celM.Offset(1, 6).Address & "," & _
                celC.Offset(1, 6).Address & "," & _
                celT.Offset(1, 6).Address & ")"
        .Font.Bold = True
        .Font.Size = 13
    End With
    
    ' Autofit
    genSht.UsedRange.EntireColumn.AutoFit
    
End Sub

Private Sub BtnInsert_Click()
    ' Insert selected 'excluded' item at cursor of 'included' list
    Dim val As Long, iter As Long, workStr As String
    Dim mch As VBScript_RegExp_55.Match
    
    ' Proofing; exit if invalid state
    If fld Is Nothing Then Exit Sub
    
    If LBxExcl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    If LBxExcl.ListIndex < 0 Then Exit Sub
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Just append if nothing selected, or if <none found> is selected
    If LBxIncl.ListIndex < 0 Or LBxIncl.List(0, 0) = NONE_FOUND Then
        BtnAppend_Click
        Exit Sub
    End If
    
    ' Reset the locked-file-found flag
    lockedFileFound = False
    
    ' Loop from the end of the 'included' list to the selection point, incrementing filenames
    val = LBxIncl.ListIndex
    For iter = LBxIncl.ListCount - 1 To LBxIncl.ListIndex Step -1
        Set mch = rxFNIdxValid.Execute(LBxIncl.List(iter, 0))(0)
        Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
        
        ' If file is locked, drop out; should trigger renumbering back to
        ' previous list state
        If Not isFileEditable(fl) Then
            notifyLockedFile fl.Name
            lockedFileFound = True
            Exit For
        End If
        
        ' Rename the file to shift it up, out of the way
        fl.Name = "(" & Format(iter + 2, NUM_FORMAT) & ")" & mch.SubMatches(2)
    Next iter
    
    ' Only bother with attempting to renumber the file to insert if no
    ' locked files were found
    If Not lockedFileFound Then
        ' Number the item to be added appropriately.
        ' Start by binding the file
        Set mch = rxFNIdxValid.Execute(LBxExcl.List(LBxExcl.ListIndex, 0))(0)
        Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
        
        ' Check if the file to be added is locked
        If Not isFileEditable(fl) Then
            notifyLockedFile fl.Name
            lockedFileFound = True
        Else
            ' Do the rename
            fl.Name = "(" & Format(val + 1, NUM_FORMAT) & ")" & mch.SubMatches(2)
        End If
    End If
    
    ' ALWAYS refresh listboxes, even if locked file found. This will restore
    ' the numbering to what it was before the Insert was attempted.
    popLists
    
    ' Can only get here if there was no hash problem beforehand,
    ' so just update the hash
    hash = hashFilenames
    
    ' Set the control states
    setCtrls
    
End Sub

Private Sub BtnInsertAll_Click()
    ' Iteratively insert all 'excluded' items at the current position
    
    ' Folder must be defined
    If fld Is Nothing Then Exit Sub
    
    ' Excluded list has to have items
    If LBxExcl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Must ensure something is selected in Excl
    LBxExcl.ListIndex = CLng(wsf.Max(0, LBxExcl.ListIndex))
    
    ' Reset locked file found flag
    lockedFileFound = False
    
    ' If nothing selected in LbxIncl, treat this as an append-all
    ' command, since a single Insert is treated as an Append.
    ' Otherwise, the items are appended in reverse order.
    If LBxIncl.ListIndex < 0 Then
        BtnAppendAll_Click
        Exit Sub
    Else
        ' Insert everything, allowing events to occur after each loop
        ' Adding in REVERSE sequence because the Insert dynamic causes
        ' each new item to be included ABOVE the last item inserted.
        ' Therefore, the reverse sequence results in the inserted objects
        ' retaining the same ordering as in LBxExcl before the mass insert.
        Do Until LBxExcl.List(0, 0) = NONE_FOUND Or lockedFileFound Or criticalLockedFile
            LBxExcl.ListIndex = LBxExcl.ListCount - 1
            BtnInsert_Click
            DoEvents
        Loop
    End If
    
    ' Form refresh and hash updates are handled by BtnRemove,
    ' so no need to do either here.
    
End Sub

Private Sub BtnMoveDown_Click()
    ' Move selected item down in the 'included' list
    
    Dim val As Long, iter As Long, fl As File
    Dim mch As VBScript_RegExp_55.Match
    
    ' Must be something in the 'included' list
    If LBxIncl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    ' Something must be selected
    If LBxIncl.ListIndex < 0 Then Exit Sub
    
    ' Can't move the last item down
    If LBxIncl.ListIndex > LBxIncl.ListCount - 2 Then Exit Sub
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Do the switch
    ' Store the index for later reference
    val = LBxIncl.ListIndex
    
    ' Check that both files are editable; warn and cancel if not
    For iter = val To val + 1
        Set fl = fs.GetFile(fs.BuildPath(fld.Path, LBxIncl.List(iter, 0)))
        If Not isFileEditable(fl) Then
            notifyLockedFile fl.Name
            lockedFileFound = True
            Exit Sub
        End If
    Next iter
    
    ' No locked file found
    lockedFileFound = False
    
    ' Move the selected file down
    Set mch = rxFNIdxValid.Execute(LBxIncl.List(val, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    fl.Name = "(" & Format(val + 2, NUM_FORMAT) & ")" & mch.SubMatches(2)
    
    ' Move the 'down' file into the vacated spot
    Set mch = rxFNIdxValid.Execute(LBxIncl.List(val + 1, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    fl.Name = "(" & Format(val + 1, NUM_FORMAT) & ")" & mch.SubMatches(2)
    
    ' Select the 'moved down' item
    LBxIncl.ListIndex = val + 1
    
    
    ' Refresh listboxes
    popLists
    
    ' Can only get here if there was no hash problem beforehand,
    ' so just update the hash
    hash = hashFilenames
    
    ' Set the control states
    setCtrls
    
End Sub

Private Sub BtnMoveAfter_Click()
    ' Move selected 'included' item to after a given index
    
    Dim srcIdx As Long, tgtIdx As Long, workStr As String
    
    ' Must be items in the list
    If LBxIncl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    ' Must have at least two items
    If LBxIncl.ListCount < 2 Then Exit Sub
    
    ' Something must be selected
    If LBxIncl.ListIndex < 0 Then Exit Sub
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Query for the desired destination
    workStr = ""
    Do
        If Not workStr = "" Then
            Call MsgBox("Please enter a number.", vbOKOnly + vbExclamation, "Warning")
        End If
        workStr = InputBox("Move selected item to a position" & vbLf & "just after item number:" & vbLf & vbLf & _
                    "(Zero moves to top of list)", "Move After...")
        If workStr = "" Then Exit Sub  ' because user cancelled
    Loop Until IsNumeric(workStr)
    
    ' Identify relevant indices
    ' If a too-big or too-small index is provided,
    ' just move to end or start of list.
    srcIdx = LBxIncl.ListIndex
    tgtIdx = wsf.Max(wsf.Min(CLng(workStr) - 1, LBxIncl.ListCount - 1), -1)
    
    ' Reset the 'locked file found' flag since this is a new attempt
    lockedFileFound = False
    
    ' Perform the move. Relies on BtnMoveDown_Click and BtnMoveUp_Click
    ' keeping the item being moved as the selected item after the move
    ' is done!
    If srcIdx < tgtIdx Then
        Do Until (LBxIncl.ListIndex = tgtIdx) Or lockedFileFound
            BtnMoveDown_Click
            DoEvents
        Loop
    ElseIf srcIdx > tgtIdx Then
        Do Until (LBxIncl.ListIndex = tgtIdx + 1) Or lockedFileFound
            BtnMoveUp_Click
            DoEvents
        Loop
    ' Do nothing if source and target indices are equal
    End If
    
    ' No repop should be needed, since this is implemented using
    ' other methods that do the repop themselves.
    
    ' Do reset the locked file found flag, though
    lockedFileFound = False
    
End Sub

Private Sub BtnMoveUp_Click()
    
    Dim val As Long, iter As Long, fl As File
    Dim mch As VBScript_RegExp_55.Match
    
    ' Must be items in the list
    If LBxIncl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    ' Can't move the top item up...
    If LBxIncl.ListIndex < 1 Then Exit Sub
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Do the switch
    ' Store the index for later reference
    val = LBxIncl.ListIndex
    
    ' Check that both files are editable; warn and cancel if not
    For iter = val To val - 1 Step -1
        Set fl = fs.GetFile(fs.BuildPath(fld.Path, LBxIncl.List(iter, 0)))
        If Not isFileEditable(fl) Then
            notifyLockedFile fl.Name
            lockedFileFound = True
            Exit Sub
        End If
    Next iter
    
    ' No locked file found
    lockedFileFound = False
    
    ' Move the selected file up
    Set mch = rxFNIdxValid.Execute(LBxIncl.List(val, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    fl.Name = "(" & Format(val, NUM_FORMAT) & ")" & mch.SubMatches(2)
    
    ' Move the 'up' file into the vacated spot
    Set mch = rxFNIdxValid.Execute(LBxIncl.List(val - 1, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    fl.Name = "(" & Format(val + 1, NUM_FORMAT) & ")" & mch.SubMatches(2)
    
    ' Select the 'moved up' item
    LBxIncl.ListIndex = val - 1
    
    
    ' Refresh listboxes
    popLists
    
    ' Can only get here if there was no hash problem beforehand,
    ' so just update the hash
    hash = hashFilenames
    
    ' Set the control states
    setCtrls
    
End Sub

Private Sub BtnOpen_Click()
    ' Prompt for user selection of folder to use

    Dim fd As FileDialog
    Dim workStr As String
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .InitialView = msoFileDialogViewList
        .Title = "Select folder for sorting"
        If .Show = 0 Then Exit Sub
        
        Set fld = fs.GetFolder(.SelectedItems(1))
    End With
    
    ' Populate the folder path textbox with the full path
    TBxFld = fld.Path
    
    ' Reset the critical locked file flag, presuming user
    ' fixed the problem
    criticalLockedFile = False
    
    ' Proof for parens and collisions
    proofParens
    proofCollisions
    
    
    ' Update the listboxes
    popLists
    
    ' Update the hash and reset the hash-match flag
    hash = hashFilenames
    anyHashMismatch = False
    
    ' Set the control states
    setCtrls
    
End Sub

Private Sub BtnOpenExcl_Click()
    ' Open the selected file in the exclude list with the default viewer.
    
    Dim shl As New Shell, filePath As String
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Open the file
    If Not fld Is Nothing Then
        If LBxExcl.ListIndex > -1 And LBxExcl.Value <> NONE_FOUND Then
            filePath = fs.BuildPath(fld.Path, LBxExcl.Value)
            shl.ShellExecute filePath
        End If
    End If

End Sub

Private Sub BtnOpenIncl_Click()
    ' Open the selected file in the include list with the default viewer.
    
    Dim shl As New Shell, filePath As String
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Open the file
    If Not fld Is Nothing Then
        If LBxIncl.ListIndex > -1 And LBxIncl.Value <> NONE_FOUND Then
            filePath = fs.BuildPath(fld.Path, LBxIncl.Value)
            shl.ShellExecute filePath
        End If
    End If
    
End Sub

Private Sub BtnReload_Click()
    ' Reset the critical locked file flag, presuming user
    ' fixed the problem
    criticalLockedFile = False
    
    ' Proof parens and collisions
    proofParens
    proofCollisions
    
    
    ' Update the listboxes
    popLists
    
    ' Update the hash and reset the hash-match flag
    hash = hashFilenames
    anyHashMismatch = False
    
    ' Set the control states
    setCtrls
    
End Sub

Private Sub BtnRemove_Click()
    ' Remove the selected 'included' item to the 'excluded' list
    
    Dim mch As VBScript_RegExp_55.Match
    
    ' Folder has to be selected
    If fld Is Nothing Then Exit Sub
    
    ' Included list has to have items in it
    If LBxIncl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    ' Something has to be selected in the 'included' list
    If LBxIncl.ListIndex < 0 Then Exit Sub
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Should be fine to remove now, as long as it's editable
    Set mch = rxFNIdxValid.Execute(LBxIncl.List(LBxIncl.ListIndex, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    If Not isFileEditable(fl) Then
        notifyLockedFile fl.Name
        lockedFileFound = True
        Exit Sub
    End If
    
    ' Is editable; remove
    lockedFileFound = False
    fl.Name = "(x)" & mch.SubMatches(2)
    
    ' Refresh listboxes
    popLists
    
    ' Can only get here if there was no hash problem beforehand,
    ' so just update the hash
    hash = hashFilenames
    
    ' Set the control states
    setCtrls
    
End Sub

Private Sub BtnRemoveAll_Click()
    ' Iteratively remove all 'included' items
    
    ' Folder must be defined
    If fld Is Nothing Then Exit Sub
    
    ' Included list has to have items
    If LBxIncl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    ' Hash check; will notify of need to refresh the form if fails
    ' Need to exit sub if it fails
    If Not doHashCheck Then Exit Sub
    
    ' Must ensure something is selected in Incl
    LBxIncl.ListIndex = CLng(wsf.Max(0, LBxIncl.ListIndex))
    
    ' Reset locked file found flag
    lockedFileFound = False
    
    ' Remove everything, allowing events to occur after each loop
    Do Until LBxIncl.List(0, 0) = NONE_FOUND Or lockedFileFound Or criticalLockedFile
        LBxIncl.ListIndex = LBxIncl.ListCount - 1
        BtnRemove_Click
        DoEvents
    Loop
    
    ' Reselect the first index, so that LBxIncl.Value is updated to match
    ' the NONE_FOUND value added to .List by popLists
    LBxIncl.ListIndex = 0
    
    ' Form refresh and hash updates are handled by BtnRemove,
    ' so no need to do either here.
    
End Sub

Private Sub BtnShowFolder_Click()
    ' Open the selected folder in Explorer
    
    Dim shl As New Shell
    
    If Not fld Is Nothing Then
        shl.ShellExecute "explorer.exe", fld.Path, , "Open", 1
    End If
    
End Sub

Private Sub SpBtnPadWidth_SpinDown()
    ' Adjust the pad width label down; min 1
    
    Dim fl As File
    Dim newPadWidth As Long
    
    ' Only worry about folder contents if folder selected
    If Not fld Is Nothing Then
        ' Hash check; will notify of need to refresh the form if fails
        ' Need to exit sub if it fails
        If Not doHashCheck Then Exit Sub
    End If
    
    ' Decrement the label whether the folder is bound or not,
    ' and store to the doc prop
    newPadWidth = wsf.Max(minPadWidth, LblPadWidth.Caption - 1)
    
    If newPadWidth <> CLng(LblPadWidth.Caption) Then
        setPadWidth newPadWidth
    End If
    
    ' Only rename files if a folder is selected
    If Not fld Is Nothing Then
        ' The list repopulation method should take care of updating
        ' all the filenames with the new pad width
        popLists
        
        ' Can only get here if there was no hash problem beforehand,
        ' so just update the hash
        hash = hashFilenames
        
        ' Set the control states
        setCtrls
    End If
    
End Sub

Private Sub SpBtnPadWidth_SpinUp()
    
    Dim fl As File
    Dim newPadWidth As Long
    
    ' Only worry about folder contents if folder selected
    If Not fld Is Nothing Then
        ' Hash check; will notify of need to refresh the form if fails
        ' Need to exit sub if it fails
        If Not doHashCheck Then Exit Sub
    End If
    
    newPadWidth = wsf.Min(6, LblPadWidth.Caption + 1)
    
    ' Adjust the pad width label up
    If newPadWidth <> CLng(LblPadWidth.Caption) Then
        setPadWidth newPadWidth
    End If
    
    ' Only rename files if a folder is selected
    If Not fld Is Nothing Then
        ' The list repopulation method should take care of updating
        ' all the filenames with the new pad width
        popLists
        
        ' Can only get here if there was no hash problem beforehand,
        ' so just update the hash
        hash = hashFilenames
        
        ' Set the control states
        setCtrls
    End If
    
End Sub













Private Sub UserForm_Initialize()
    ' Initialize userform globals &c.
    
    Dim workStr As String
    Dim dp As DocumentProperty
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set wsf = Application.WorksheetFunction
    NL = Chr(10)
    
    ' Init Regexes
    compileRegexes
    
    ' Calculate the initial hash
    hash = hashFilenames
    
    ' Set the pad width, either from the default or stored custom docprop
    Set dp = getCustDocProp(CDP_PADWIDTH)
    If dp Is Nothing Then
        ' Set the doc prop and populate the form label
        setCustDocProp CDP_PADWIDTH, CStr(DEF_PADWIDTH)
        LblPadWidth.Caption = CStr(DEF_PADWIDTH)
    Else
        ' Populate the form label frmo the doc prop
        LblPadWidth.Caption = dp.Value
    End If
    
    ' For now, the form will not be loading with a folder bound.
    ' So, just set the controls' states without attempting
    ' to populate the listboxes.
    setCtrls
    
End Sub

Private Sub compileRegexes()
    ' Helper to recompile regexes
    ' Anticipates implementation of customizable item categories
    
    ' Valid starting '(...)' index format, generalized to allow
    ' multiple "x"s for 'excluded' files.
    With rxFNIdxValid
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^\((([0-9]+|x)+)\)(.+)$"
    End With
    
    ' Detailed matching of an included file, with submatches
    With rxInclFnameDetail
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^\(([0-9]+)\)\s+\[([SEMCT])\](.+?) - (.+) -- ([0-9.]+)\(([0-9.]+)\)\.[_0-9a-z]+$"
    End With
    
    ' Detailed matching of an excluded file, with submatches
    ' Submatch catching the 'x' is retained for index parity with the included files
    With rxExclFnameDetail
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^\((x)\)\s+\[([SEMCT])\](.+?) - (.+) -- ([0-9.]+)\(([0-9.]+)\)\.[_0-9a-z]+$"
    End With
    
    ' Detailed matching of an included or excluded file, with submatches
    With rxFnameDetail
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^\(([0-9]+|x)\)\s+\[([SEMCT])\](.+?) - (.+) -- ([0-9.]+)\(([0-9.]+)\)\.[_0-9a-z]+$"
    End With
    
End Sub






Private Function getCustDocProp(dpName As String) As DocumentProperty
    ' Helper for retrieving custom document property.
    ' Returns the DocumentProperty if it's present; otherwise, Nothing
    
    Dim errNum As Long, dp As DocumentProperty
    
    ' Attempt retrieve
    On Error Resume Next
        Set dp = ThisWorkbook.CustomDocumentProperties(dpName)
    errNum = Err.Number: Err.Clear: On Error GoTo 0
    
    ' Return based on whether error occurred
    If errNum > 0 Then
        Set getCustDocProp = Nothing
    Else
        Set getCustDocProp = dp
    End If
    
End Function

Private Function setCustDocProp(dpName As String, strVal As String) As DocumentProperty
    ' Set the indicated custom doc property.
    ' Any existing property is deleted.
    '
    ' Returns the created docprop object
    '
    ' DOES NOT SAVE THE ADDIN WORKBOOK, so the change will not be retained
    ' on Excel exit/reopen UNLESS ThisWorkbook.Save is called.
    ' Here, this .Save is called on click of BtnClose.
    
    ' Delete prop if present
    On Error Resume Next
        ThisWorkbook.CustomDocumentProperties(dpName).Delete
    Err.Clear: On Error GoTo 0
    
    With ThisWorkbook
        ' Add a new property
        .CustomDocumentProperties.Add _
                Name:=dpName, _
                LinkToContent:=False, _
                Type:=msoPropertyTypeString, _
                Value:=strVal
        
        ' Set the return value
        Set setCustDocProp = .CustomDocumentProperties(dpName)
    End With
    
End Function


Sub clearAddinCustDocProps()
    ' Helper to clear the custom document properties defined on
    ' the add-in .xlam itself, to facilitate preparation of
    ' binaries for release.
    
    On Error Resume Next
        With ThisWorkbook.CustomDocumentProperties
            .Item(CDP_PADWIDTH).Delete
        End With
    Err.Clear: On Error GoTo 0

End Sub
