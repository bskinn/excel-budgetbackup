VERSION 5.00
Begin {C62A69F0-16DC-11CE-9E98-00AA00574A4F} FrmWait 
   Caption         =   "Locating Adobe Reader...."
   ClientHeight    =   1995
   ClientLeft      =   45
   ClientTop       =   375
   ClientWidth     =   7170
   OleObjectBlob   =   "FrmWait.frx":0000
   ShowModal       =   0   'False
   StartUpPosition =   1  'CenterOwner
End
Attribute VB_Name = "FrmWait"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = True
Attribute VB_Exposed = False
' # ------------------------------------------------------------------------------
' # Name:        FrmWait.frm
' # Purpose:     Form to show current search folder in Reader location code
' #                (obsolete)
' #
' # Author:      Brian Skinn
' #                bskinn@alum.mit.edu
' #
' # Created:     15 Jan 2015
' # Copyright:   (c) Brian Skinn 2017
' # License:     The MIT License; see "LICENSE.txt" for full license terms
' #                   and contributor agreement.
' #
' #       http://www.github.com/bskinn/excel-budgetbackup
' #
' # ------------------------------------------------------------------------------

Option Explicit

Public stopFlag As Boolean

Private Sub BtnCancel_Click()
    stopFlag = True
End Sub

Private Sub UserForm_Initialize()
    stopFlag = False
End Sub
