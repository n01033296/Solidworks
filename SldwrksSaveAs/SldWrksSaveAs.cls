VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
  Persistable = 0  'NotPersistable
  DataBindingBehavior = 0  'vbNone
  DataSourceBehavior  = 0  'vbNone
  MTSTransactionMode  = 0  'NotAnMTSObject
END
Attribute VB_Name = "Class1"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = True
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = True
Option Explicit
    Dim swApp               As SldWorks.SldWorks
    Dim swModel             As SldWorks.ModelDoc2
    Dim swModelDocExt       As SldWorks.ModelDocExtension
    Dim swExportData        As SldWorks.ExportPdfData
    Dim boolstatus          As Boolean
    Dim filename            As String
    Dim filename1           As String
    Dim filename2           As String
    Dim filenamepdf         As String
    Dim filenamedxf         As String
    Dim filenamedwg         As String
    Dim filepath            As String
    Dim lErrors             As Long
    Dim lWarnings           As Long


Sub main()
    Set swApp = Application.SldWorks
    Set swModel = swApp.ActiveDoc
    swApp.GetFirstDocument
        'Here is where we need to determine whether the filename is 2 alpha or 3 alph
        
        'Take the first 3 characters which determines sub-directory you will go to (LC1)
        filename2 = Left(swModel.GetTitle, 3)
        'Add 3 0s to make a complete folder name (LC1000, LC4000, LC5000, LC6000)
        filename2 = filename2 + "000"
        'Provide the common filepath to be used by all pdf's and ADD 1st letter directory (A, B, C, D, etc.)
        filepath = "C:\Users\amirtaheri\Desktop\test\" & Left(swModel.GetTitle, 1)
        'Adding to the above filepath the customer code (LC/BB/EL etc.)
        
        'this code is going to become gettitle, 2 or gettitle, 3 depending on # of characters prefixing part #
        filepath = filepath & "\" & Left(swModel.GetTitle, 2) & "\" & filename2 & "\"
       'That gives the full filepath to be used
       
    If swModel Is Nothing Then
        MsgBox "No current document", vbCritical
        End
    End If
    If swModel.GetType <> swDocDRAWING Then
        MsgBox "This Macro only works on Drawings", vbCritical
        End
    End If
    Set swModelDocExt = swModel.Extension
    Set swExportData = swApp.GetExportFileData(swExportPdfData)

    filename = swModel.GetTitle
   
   'Save drawing
'    boolstatus = swExportData.SetSheets(swExportData_ExportAllSheets, 1)
'    boolstatus = swModelDocExt.SaveAs(filename1, 0, 0, swExportData, lErrors, lWarnings)
    
'    If boolstatus Then
'        MsgBox "Save as Drawing successful" & vbNewLine & filename
'    Else
'        MsgBox "Save as Drawing failed, Error code:" & lErrors
'    End If
   
    'Save as PDF
    filenamepdf = filepath & filename & ".PDF"
    boolstatus = swExportData.SetSheets(swExportData_ExportAllSheets, 1)
    boolstatus = swModelDocExt.SaveAs(filenamepdf, 0, 0, swExportData, lErrors, lWarnings)
    If boolstatus Then
        MsgBox "Save as PDF successful" & vbNewLine & filepath
    Else
        MsgBox "Save as PDF failed, Error code:" & lErrors
    End If

    'Save as DXF
    filenamedxf = filepath & filename & ".DXF"
    boolstatus = swExportData.SetSheets(swExportData_ExportAllSheets, 1)
    boolstatus = swModelDocExt.SaveAs(filenamedxf, 0, 0, swExportData, lErrors, lWarnings)
    If boolstatus Then
        MsgBox "Save as DXF successful" & vbNewLine & filepath
    Else
        MsgBox "Save as DXF failed, Error code:" & lErrors
    End If
       
    'Save as AutoCad drawing
    filenamedwg = filepath & filename & ".DWG"
    boolstatus = swExportData.SetSheets(swExportData_ExportAllSheets, 1)
    boolstatus = swModelDocExt.SaveAs(filenamedwg, 0, 0, swExportData, lErrors, lWarnings)
    If boolstatus Then
        MsgBox "Save as DWG successful" & vbNewLine & filepath
    Else
        MsgBox "Save as DWG failed, Error code:" & lErrors
    End If
   
End Sub

