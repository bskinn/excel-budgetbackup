VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmBackupSort 
   Caption         =   "Sort Budget Items"
   ClientHeight    =   8070
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
' # Copyright:   (c) Brian Skinn 2017
' # License:     The MIT License; see "LICENSE.txt" for full license terms
' #                   and contributor agreement.
' #
' #       http://www.github.com/bskinn/excel-budgetbackup
' #
' # ------------------------------------------------------------------------------

Option Explicit

Dim fs As Scripting.FileSystemObject, fld As Scripting.Folder, fl As Scripting.File
Dim mch As VBScript_RegExp_55.Match, wsf As WorksheetFunction
Dim rx As New VBScript_RegExp_55.RegExp
Dim populated As Boolean, wasPacked As Boolean
Dim inclIdx As Long, exclIdx As Long
Dim inclView As Long, exclView As Long
Dim cancelLoad As Boolean, openBtnState As Boolean
Const NONE_FOUND As String = "<none found>"
Const EMPTY_LIST As String = "<empty>"
Const NUM_FORMAT As String = "00"
Const READER_PATH As String = "C:\Program Files (x86)\Adobe\Reader 10.0\Reader\AcroRd32.exe"
Const READER_PROP_NAME As String = "ReaderEXE"
Const READER_EXE As String = "AcroRd32.exe"
Const CANCEL_RETURN As String = "!!CANCELED!!"

Public Sub clearReaderLocation()
    Dim prp As DocumentProperty, iter As Long
    
    For iter = ThisWorkbook.CustomDocumentProperties.Count To 1 Step -1
        Set prp = ThisWorkbook.CustomDocumentProperties(iter)
        If prp.Name = READER_PROP_NAME Then prp.Delete
    Next iter
    
End Sub

Private Sub popLists(Optional firstCall As Boolean = True)
    Dim ctrl As Control

    ' If the first call, disable everything
    If firstCall Then
        For Each ctrl In FrmBackupSort.Controls
            If TypeOf ctrl Is CommandButton Or _
                    TypeOf ctrl Is ListBox Then
                ctrl.Enabled = False
            End If
        Next ctrl
    End If

    ' Store current selections if not a re-call from the list packing
    If Not wasPacked Then
        exclIdx = LBxExcl.ListIndex
        inclIdx = LBxIncl.ListIndex
        
        exclView = LBxExcl.TopIndex
        inclView = LBxIncl.TopIndex
    End If
    
    ' Clear list contents
    LBxExcl.Clear
    LBxIncl.Clear
    
    If fld Is Nothing Then
        ' Empty population indication
        LBxExcl.AddItem "<empty>"
        LBxIncl.AddItem "<empty>"
        populated = False
        GoTo Final_Exit
    End If
    
    padNums
    
    For Each fl In fld.Files
        If rx.Test(fl.Name) Then
            Set mch = rx.Execute(fl.Name)(0)
            If LCase(mch.SubMatches(1)) = "x" Then
                LBxExcl.AddItem fl.Name
            Else
                LBxIncl.AddItem fl.Name
            End If
        End If
    Next fl
    
    wasPacked = False
    packNums
    If wasPacked Then Call popLists(False)
    
    ' Only run the finishing stuff if the initial call
    If firstCall Then GoTo Final_Exit
    
    Exit Sub

Final_Exit:
    ' Indicate empty include/excluded lists if detected
    If LBxExcl.ListCount < 1 Then LBxExcl.AddItem NONE_FOUND
    If LBxIncl.ListCount < 1 Then LBxIncl.AddItem NONE_FOUND
    
    ' Restore selections and views
    LBxExcl.ListIndex = wsf.Min(exclIdx, LBxExcl.ListCount - 1)
    LBxIncl.ListIndex = wsf.Min(inclIdx, LBxIncl.ListCount - 1)
    LBxExcl.TopIndex = exclView
    LBxIncl.TopIndex = inclView
    
    ' If the first call, re-enable everything
    For Each ctrl In FrmBackupSort.Controls
        If TypeOf ctrl Is CommandButton Or _
                TypeOf ctrl Is ListBox Then
            ctrl.Enabled = True
        End If
    Next ctrl
    
    ' Readjust the Enabled state of the Open buttons as needed
    BtnOpenExcl.Enabled = openBtnState
    BtnOpenIncl.Enabled = openBtnState

End Sub

Private Sub padNums()
    For Each fl In fld.Files
        If rx.Test(fl.Name) Then
            Set mch = rx.Execute(fl.Name)(0)
            ' I think .SubMatches(0) is the inner item that has the '+' applied to it,
            '  while .SM(1)is the full numerical match.  .SM(2) is the remainder of the
            '  filename.
            If LCase(mch.SubMatches(1)) <> "x" And Len(mch.SubMatches(1)) = 1 Then
                fl.Name = "(0" & mch.SubMatches(1) & ")" & mch.SubMatches(2)
            End If
        End If
    Next fl
End Sub

Private Sub packNums()
    Dim workStr As String, iter As Long
    
    If LBxIncl.ListCount > 0 Then
        If LBxIncl.List(0, 0) <> NONE_FOUND Then
            For iter = 0 To LBxIncl.ListCount - 1
                Set mch = rx.Execute(LBxIncl.List(iter, 0))(0)
                If Not CLng(mch.SubMatches(1)) - 1 = iter Then
                    wasPacked = True
                    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
                    fl.Name = "(" & Format(iter + 1, "00") & ")" & mch.SubMatches(2)
                End If
            Next iter
        End If
    End If
    
End Sub

Private Function LocateReader() As String
    ' Finds full path to AcroRd32.exe and returns it
    ' Also stores it in a docProp
    
    Dim dp As DocumentProperty, readerProp As DocumentProperty
    'Dim fs As FileSystemObject
    Dim workStr As String
    
    ' Check if docprop exists; bind if so; create if not
    With ThisWorkbook
        If .CustomDocumentProperties.Count > 0 Then
            For Each dp In .CustomDocumentProperties
                If dp.Name = READER_PROP_NAME Then
                    Set readerProp = dp
                    Exit For
                End If
            Next dp
        End If
        
        If readerProp Is Nothing Then
            .CustomDocumentProperties.Add READER_PROP_NAME, False, msoPropertyTypeString, ""
            Set readerProp = .CustomDocumentProperties(READER_PROP_NAME)
        End If
    End With
    
    ' Bind filesystem object (nah, should already be bound as a global)
    'Set fs = CreateObject("Scripting.FileSystemObject")
    
    ' Check if file in property exists. If so, assume valid and return as such.
    If fs.FileExists(readerProp.Value) Then
        LocateReader = readerProp.Value
        Exit Function
    End If
    
    ' File doesn't exist; have to track down Reader
    ' Assume only need to search C: drive, 'Program Files' subfolders
    workStr = RecursiveFileSearch(READER_EXE, fs.GetFolder("C:\Program Files"))
    
    ' Only check for the (x86) folder if not found in 'base' version
    If workStr = "" Then
        If fs.FolderExists("C:\Program Files (x86)") Then
            workStr = RecursiveFileSearch(READER_EXE, fs.GetFolder("C:\Program Files (x86)"))
        End If
    End If
    
    ' Apply string into property, save the addin file, and return the path
    readerProp.Value = workStr
    ThisWorkbook.Save
    LocateReader = workStr
    
End Function

Private Function RecursiveFileSearch(fName As String, baseFld As Folder) As String
    ' Empty string return means not found, keep looking
    ' Non-empty string should contain desired path to file
    Dim fl As File, fld As Folder
    Dim workStr As String
    
    ' Initialize unsuccessful return
    RecursiveFileSearch = ""
    
    ' Update the notification label
    FrmWait.LblCurrFld.Caption = baseFld.Path
    
    ' First search the base folder
    For Each fl In baseFld.Files
        If fl.Name = fName Then
            RecursiveFileSearch = fl.Path
            Exit Function
        End If
        
        ' Ensure able to 'hear' Cancel button press
        DoEvents
        
        ' If cancel pressed, dump out
        If FrmWait.stopFlag = True Then
            RecursiveFileSearch = CANCEL_RETURN
            Exit Function
        End If
    Next fl
    
    ' File not found; recurse through the subfolders
    For Each fld In baseFld.SubFolders
        workStr = RecursiveFileSearch(fName, fld)
        If workStr <> "" Then
            RecursiveFileSearch = workStr
            Exit Function
        End If
    Next fld
    
End Function

Private Sub BtnAppend_Click()
    If fld Is Nothing Then Exit Sub
    
    If LBxExcl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    If LBxExcl.ListIndex < 0 Then Exit Sub
    
    Set mch = rx.Execute(LBxExcl.List(LBxExcl.ListIndex, 0))(0)
    
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    If LBxIncl.List(0, 0) = NONE_FOUND Then
        ' Start list from nothing
        fl.Name = "(" & Format(1, NUM_FORMAT) & ")" & mch.SubMatches(2)
    Else
        fl.Name = "(" & Format(LBxIncl.ListCount + 1, NUM_FORMAT) & ")" & mch.SubMatches(2)
    End If
    
    popLists
    
End Sub

Private Sub BtnClose_Click()
    Unload FrmBackupSort
End Sub

Private Sub BtnGenSheet_Click()
    Dim genBk As Workbook, genSht As Worksheet
    Dim sht As Worksheet
    Dim workCel As Range, tblCel As Range
    Dim celS As Range, celE As Range, celM As Range, celC As Range, celT As Range
    Dim rx As New RegExp, mchs As MatchCollection, mch As Match
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
    
    ' Drop if not init
    If fld Is Nothing Then Exit Sub
    
    ' Set up the regex
    With rx
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^\(([0-9]+)\)\s+\[([SEMCT])\](.+?) - (.+) -- ([0-9.]+)\(([0-9.]+)\)\.[0-9a-z]+$"
    End With
    
    ' Scan the work folder for properly configured filenames
    counts = Array(0, 0, 0, 0, 0)
    inlaids = Array(0, 0, 0, 0, 0)
    anyFlsFound = False
    For Each fl In fld.Files
        If rx.Test(fl.Name) Then
            anyFlsFound = True
            Set mch = rx.Execute(fl.Name)(0)
            
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
        Else
            ' Notify of non-matching item, if not an 'excluded' item
            If Not UCase(Left(fl.Name, 3)) = "(X)" Then
                MsgBox "The following item is named in an unrecognized format " & _
                        "and will be skipped:" & vbCrLf & vbCrLf & fl.Name, _
                        vbOKOnly + vbExclamation, "Skipping item"
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
        If rx.Test(fl.Name) Then
            Set mch = rx.Execute(fl.Name)(0)
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
    
    
    genSht.UsedRange.Columns(genSht.UsedRange.Columns.Count) _
            .HorizontalAlignment = xlLeft
    
    ' Total field
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
    
    Dim val As Long, iter As Long, workStr As String
    
    If fld Is Nothing Then Exit Sub
    
    If LBxExcl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    If LBxExcl.ListIndex < 0 Then Exit Sub
    
    If LBxIncl.ListIndex < 0 Or LBxIncl.Value = NONE_FOUND Then
        ' Just append if nothing selected, or if <none found> is selected
        BtnAppend_Click
        Exit Sub
    End If
    
    ' Loop from the end of the 'included' list to the selection point, incrementing filenames
    val = LBxIncl.ListIndex
    For iter = LBxIncl.ListCount - 1 To LBxIncl.ListIndex Step -1
        Set mch = rx.Execute(LBxIncl.List(iter, 0))(0)
        Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
        ' Need to add trap for if file is locked, EVERY TIME a file is renamed.
        '  Probably will want a utility function for this
        fl.Name = "(" & Format(iter + 2, NUM_FORMAT) & ")" & mch.SubMatches(2)
    Next iter
    
    ' Number the item to be added appropriately
    Set mch = rx.Execute(LBxExcl.List(LBxExcl.ListIndex, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    fl.Name = "(" & Format(val + 1, NUM_FORMAT) & ")" & mch.SubMatches(2)
    
    ' Repopulate the lists
    popLists
    
End Sub

Private Sub BtnMoveDown_Click()
    
    Dim val As Long
    
    If LBxIncl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    ' Something must be selected
    If LBxIncl.ListIndex < 0 Then Exit Sub
    
    ' Can't move the last item down
    If LBxIncl.ListIndex > LBxIncl.ListCount - 2 Then Exit Sub
    
    ' Do the switch
    ' Store the index for later reference
    val = LBxIncl.ListIndex
    
    ' Fragile to identical filenames except for the number, but this should
    '  only happen in stupid cases, not most real-life scenarios
    ' Move the selected file down
    Set mch = rx.Execute(LBxIncl.List(val, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    fl.Name = "(" & Format(val + 2, NUM_FORMAT) & ")" & mch.SubMatches(2)
    
    ' Move the 'down' file into the vacated spot
    Set mch = rx.Execute(LBxIncl.List(val + 1, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    fl.Name = "(" & Format(val + 1, NUM_FORMAT) & ")" & mch.SubMatches(2)
    
    ' Select the 'moved down' item
    LBxIncl.ListIndex = val + 1
    
    ' Repop the lists
    popLists
    
End Sub

Private Sub BtnMoveAfter_Click()
    
    Dim srcIdx As Long, tgtIdx As Long, workStr As String
    
    If LBxIncl.List(0, 0) = NONE_FOUND Then Exit Sub
    If LBxIncl.ListCount < 2 Then Exit Sub
    If LBxIncl.ListIndex < 0 Then Exit Sub
       
    workStr = ""
       
    Do
        If Not workStr = "" Then
            Call MsgBox("Please enter a number.", vbOKOnly + vbExclamation, "Warning")
        End If
        workStr = InputBox("Move selected item to a position" & vbLf & "just after item number:" & vbLf & vbLf & _
                    "(Zero moves to top of list)", "Move After...")
        If workStr = "" Then Exit Sub
    Loop Until IsNumeric(workStr)
    
    srcIdx = LBxIncl.ListIndex
    tgtIdx = wsf.Max(wsf.Min(CLng(workStr) - 1, LBxIncl.ListCount - 1), -1)
    
    If srcIdx < tgtIdx Then
        Do Until LBxIncl.ListIndex = tgtIdx
            BtnMoveDown_Click
        Loop
    ElseIf srcIdx > tgtIdx Then
        Do Until LBxIncl.ListIndex = tgtIdx + 1
            BtnMoveUp_Click
        Loop
    End If
    
'    Do Until LBxIncl.ListIndex = tgtIdx
'        If srcIdx < tgtIdx Then
'            BtnMoveDown_Click
'        ElseIf srcIdx > tgtIdx Then
'            BtnMoveUp_Click
'        End If
'    Loop
    
    ' No repop appears to be needed
    
End Sub

Private Sub BtnMoveUp_Click()
    
    Dim val As Long
    
    If LBxIncl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    ' Can't move the top item up...
    If LBxIncl.ListIndex < 1 Then Exit Sub
    
    ' Do the switch
    ' Store the index for later reference
    val = LBxIncl.ListIndex
    
    ' Fragile to identical filenames except for the number, but this should
    '  only happen in stupid cases, not most real-life scenarios
    ' Move the selected file up
    Set mch = rx.Execute(LBxIncl.List(val, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    fl.Name = "(" & Format(val, NUM_FORMAT) & ")" & mch.SubMatches(2)
    
    ' Move the 'up' file into the vacated spot
    Set mch = rx.Execute(LBxIncl.List(val - 1, 0))(0)
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    fl.Name = "(" & Format(val + 1, NUM_FORMAT) & ")" & mch.SubMatches(2)
    
    ' Select the 'moved up' item
    LBxIncl.ListIndex = val - 1
    
    ' Repop the lists
    popLists
    
End Sub

Private Sub BtnOpen_Click()
    Dim fd As FileDialog
    
    Set fd = Application.FileDialog(msoFileDialogFolderPicker)
    
    With fd
        .AllowMultiSelect = False
        .ButtonName = "Select"
        .InitialView = msoFileDialogViewList
        .Title = "Select folder for sorting"
        If .Show = 0 Then Exit Sub
        
        Set fld = fs.GetFolder(.SelectedItems(1))
    End With
    
    popLists
    
    TBxFld = fld.Path
    
End Sub

Private Sub BtnOpenExcl_Click()
    Dim shl As New Shell, filePath As String
    
    If Not fld Is Nothing Then
        If LBxExcl.ListIndex > -1 And LBxExcl.Value <> NONE_FOUND Then
            filePath = fs.BuildPath(fld.Path, LBxExcl.Value)
            'shl.ShellExecute READER_PATH, filePath, , "Open", 1
            shl.ShellExecute filePath
        End If
    End If

End Sub

Private Sub BtnOpenIncl_Click()
    Dim shl As New Shell, filePath As String
    
    If Not fld Is Nothing Then
        If LBxIncl.ListIndex > -1 And LBxIncl.Value <> NONE_FOUND Then
            filePath = fs.BuildPath(fld.Path, LBxIncl.Value)
            'shl.ShellExecute READER_PATH, filePath, , "Open", 1
            shl.ShellExecute filePath
        End If
    End If
    
End Sub

Private Sub BtnReload_Click()
    popLists
End Sub

Private Sub BtnRemove_Click()
    
    If fld Is Nothing Then Exit Sub
    
    If LBxIncl.List(0, 0) = NONE_FOUND Then Exit Sub
    
    If LBxIncl.ListIndex < 0 Then Exit Sub
    
    ' Should be fine to remove now
    Set mch = rx.Execute(LBxIncl.List(LBxIncl.ListIndex, 0))(0)
    
    Set fl = fs.GetFile(fs.BuildPath(fld.Path, mch.Value))
    fl.Name = "(x)" & mch.SubMatches(2)
    
    popLists
    
End Sub

Private Sub BtnShowFolder_Click()
    Dim shl As New Shell
    
    If Not fld Is Nothing Then
        shl.ShellExecute "explorer.exe", fld.Path, , "Open", 1
    End If
    
End Sub

Private Sub UserForm_Activate()
    If cancelLoad Then Unload FrmBackupSort
End Sub

Private Sub UserForm_Initialize()
    Dim workStr As String
    Dim dp As DocumentProperty, resp As VbMsgBoxResult
    
    Set fs = CreateObject("Scripting.FileSystemObject")
    Set wsf = Application.WorksheetFunction
    'Set shAp = CreateObject("Shell.Application")
    
    cancelLoad = False
    populated = False
    
    With rx
        .Global = False
        .MultiLine = False
        .IgnoreCase = True
        .Pattern = "^\((([0-9]+|x)+)\)(.+)$"
    End With
    
    popLists
    
    ' Ensure Reader location is known
    workStr = ""
    For Each dp In ThisWorkbook.CustomDocumentProperties
        If dp.Name = READER_PROP_NAME Then
            workStr = dp.Value
        End If
    Next dp
    
    If Not fs.FileExists(workStr) Then
        resp = MsgBox("Adobe Reader must be located in order to enable opening of PDFs." & _
                        vbLf & vbLf & "The search process should take less than a minute. " & _
                        vbLf & vbLf & "Locate Reader now?", vbYesNoCancel + vbQuestion, _
                        "Locate Adobe Reader?")
        Select Case resp
        Case vbYes
            ' Show the status form and locate the file
            FrmWait.Show
            workStr = LocateReader
            Unload FrmWait
            
            ' Inform outcome
            If fs.FileExists(workStr) Then
                ' Presume found the right file
                Call MsgBox("Adobe Reader successfully located.", vbOKOnly + vbInformation, _
                        "Success")
            Else
                Call MsgBox("Adobe Reader was not found in any of the usual places.", _
                        vbOKOnly + vbExclamation, "Failure")
            End If
        Case vbNo
            ' Do nothing, just pass through
        Case vbCancel
            ' Hard exit
            cancelLoad = True
        End Select
    End If
    
    ' Enable/disable 'Open' buttons depending on valid file found
    openBtnState = fs.FileExists(workStr)
    BtnOpenExcl.Enabled = openBtnState
    BtnOpenIncl.Enabled = openBtnState
    
End Sub

