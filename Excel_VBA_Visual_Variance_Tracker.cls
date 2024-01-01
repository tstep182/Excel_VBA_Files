VERSION 1.0 CLASS
BEGIN
  MultiUse = -1  'True
END
Attribute VB_Name = "VarianceTracker"
Attribute VB_GlobalNameSpace = False
Attribute VB_Creatable = False
Attribute VB_PredeclaredId = False
Attribute VB_Exposed = False
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
'NOTE: The code is this module works only with the outstanding Dodeca add-in for Essbase
'By design it does NOT work with the so-called "Smart View" Essbase add-in
'''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

Option Explicit
Option Base 1

'The values for these will be set at the time the object is created
Public EssServer_Before As String
Public EssServer_After As String
Public EssServer_DR As String
Public CubeList As Dictionary

'Dodeca constants
Const APS_URL As String = "embedded"
Const DODECA_ESSBASE_URL As String = "http://dodeca.swacorp.com:8080/dodeca-essbase"

'Sheet-based constants
Const first_data_row As Integer = 8
Const first_data_column As Integer = 8
Const first_data_column_capital As Integer = 9

'Constants for the EssVGetMemberInfo function
Const EssChildLevel As Variant = 1
Const EssDescendentLevel As Variant = 2
Const EssBottomLevel As Variant = 3
Const EssSiblingLevel As Variant = 4
Const EssSameLevel As Variant = 5
Const EssSameGenerationLevel As Variant = 6
Const EssCalculationLevel As Variant = 7
Const EssParentLevel As Variant = 8
Const EssDimensionLevel As Variant = 9
Public sheetName As String

'For use with EssVSetSheetOptions and EssVGetSheetOptions
Public Enum SheetOptions
    DrillLevel = 1
    IncludeSelection = 2
    WithinSelectedGroup = 3
    RemoveUnselectedGroups = 4
    IndentStyle = 5
    SuppressMissing = 6
    SuppressZeros = 7
    SuppressUnderscores = 8
    MissingTextLabel = 9
    UpdateMode = 10
    RetainOnRetrieve = 11
    AdjustColumns = 12
    UseAliases = 13
    AliasTable = 14
    TemplateRetrieveMode = 15
    FreeFormMode = 16
    AutoSortRows = 17
    UseStyles = 18
    NoAccessLabel = 19
    RetainOnKeepOnlyAndRemoveOnly = 21
    RetainOnZoom = 22
    FormulaFill = 23
    UseBothMemberNameAndAliases = 24
    RepeatMemberLabels = 25
    SheetOptionForQueryDesigner = 26
    UseLatestTimePeriod = 27
    SpecifyLatestTimePeriod = 28
    EnableHybridAnalysis = 29
    EnableMetadataSampling = 30
    MetadataSamplingPercentage = 32
    QualifiedNameAsExcelComment = 33
    QualifiedNameInSpreadsheet = 34
End Enum

'For use with EssVSetGlobalOptions and EssVGetGlobalOptions
Public Enum GlobalOptions
    EnableSecondaryButton = 1
    EnableDoubleClicking = 2
    EnableFlashback = 3
    SpecifyMessageLevel = 5
    EnableDisplayUnknownMembers = 6
    EnableRouteMessagesToLog = 7
    EnablePurgeLog = 8
    EnableDblClickForLROs = 9
    DisplayMemberSelectSaveDialog = 10
    EnableNavigateWithoutData = 11
End Enum

'Drill levels for use with EssVSetSheetOptions and EssVGetSheetOptions
Public Enum DrillLevels
    NextLevel = 1
    AllLevels = 2
    BottomLevel = 3
    SiblingLevel = 4
    SameLevel = 5
    SameGeneration = 6
    CalcLevel = 7
End Enum

'Indent levels for use with EssVSetSheetOptions and EssVGetSheetOptions
Public Enum IndentSettings
    NoIndentation = 1
    SubItems = 2
    Totals = 3
End Enum


''''''''''''''''''''''''''''''''
' DO NOT CHANGE THESE FUNCTIONS
''''''''''''''''''''''''''''''''
Private EssbaseVBAFunctionWrapperInstance As Object


'Get the Dodeca Add-in for Essbase Spreadsheet Toolkit reference.
Private Function EssbaseVBAFunctionWrapper() As Object
    'Set an instance of the wrapper if necessary
    If EssbaseVBAFunctionWrapperInstance Is Nothing Then
        Dim addIn As COMAddIn
        
        'Get the COM addin object
        Set addIn = application.COMAddIns("AppliedOLAP.Dodeca.ExcelAddInForEssbase")
        
        'Set the default url locations
        addIn.Object.SetDefaultUrls APS_URL, DODECA_ESSBASE_URL
        
        'Set the value
        Set EssbaseVBAFunctionWrapperInstance = addIn.Object
    End If
    
    'Return the instance
    Set EssbaseVBAFunctionWrapper = EssbaseVBAFunctionWrapperInstance
End Function


'Connect
Public Function EssMenuVConnect() As Long
    EssMenuVConnect = EssbaseVBAFunctionWrapper.EssMenuVConnect(APS_URL, DODECA_ESSBASE_URL, Excel.application.Hwnd)
End Function

Public Function EssVConnect(ByVal sheetName As Variant, ByVal userName As Variant, ByVal password As Variant, ByVal server As Variant, ByVal application As Variant, ByVal database As Variant) As Long
    EssVConnect = EssbaseVBAFunctionWrapper.EssVConnect(sheetName, userName, password, server, application, database, APS_URL, DODECA_ESSBASE_URL, Excel.application.Hwnd)
End Function


'Disconnect
Public Function EssMenuVDisconnect() As Long
    EssMenuVDisconnect = EssbaseVBAFunctionWrapper.EssMenuVDisconnect(Excel.application.Hwnd)
End Function

Public Function EssVDisconnect(ByVal sheetName As Variant) As Long
    EssVDisconnect = EssbaseVBAFunctionWrapper.EssVDisconnect(sheetName, Excel.application.Hwnd)
End Function


'Retrieve
Public Function EssMenuVRetrieve() As Long
    EssMenuVRetrieve = EssbaseVBAFunctionWrapper.EssMenuVRetrieve(Excel.application.Hwnd)
End Function

Public Function EssVRetrieve(ByVal sheetName As Variant, ByVal range As Variant, ByVal lockFlag As Variant) As Long
    EssVRetrieve = EssbaseVBAFunctionWrapper.EssVRetrieve(sheetName, range, lockFlag, Excel.application.Hwnd)
End Function


'ZoomIn
Public Function EssMenuVZoomIn() As Long
    EssMenuVZoomIn = EssbaseVBAFunctionWrapper.EssMenuVZoomIn(Excel.application.Hwnd)
End Function

Public Function EssVZoomIn(ByVal sheetName As Variant, ByVal range As Variant, ByVal selection As Variant, ByVal level As Variant, ByVal across As Variant) As Long
    EssVZoomIn = EssbaseVBAFunctionWrapper.EssVZoomIn(sheetName, range, selection, level, across, Excel.application.Hwnd)
End Function


'ZoomOut
Public Function EssMenuVZoomOut() As Long
    EssMenuVZoomOut = EssbaseVBAFunctionWrapper.EssMenuVZoomOut(Excel.application.Hwnd)
End Function

Public Function EssVZoomOut(ByVal sheetName As Variant, ByVal range As Variant, ByVal selection As Variant) As Long
    EssVZoomOut = EssbaseVBAFunctionWrapper.EssVZoomOut(sheetName, range, selection, Excel.application.Hwnd)
End Function


'KeepOnly
Public Function EssMenuVKeepOnly() As Long
    EssMenuVKeepOnly = EssbaseVBAFunctionWrapper.EssMenuVKeepOnly(Excel.application.Hwnd)
End Function

Public Function EssVKeepOnly(ByVal sheetName As Variant, ByVal range As Variant, ByVal selection As Variant) As Long
    EssVKeepOnly = EssbaseVBAFunctionWrapper.EssVKeepOnly(sheetName, range, selection, Excel.application.Hwnd)
End Function


'RemoveOnly
Public Function EssMenuVRemoveOnly() As Long
    EssMenuVRemoveOnly = EssbaseVBAFunctionWrapper.EssMenuVRemoveOnly(Excel.application.Hwnd)
End Function

Public Function EssVRemoveOnly(ByVal sheetName As Variant, ByVal range As Variant, ByVal selection As Variant) As Long
    EssVRemoveOnly = EssbaseVBAFunctionWrapper.EssVRemoveOnly(sheetName, range, selection, Excel.application.Hwnd)
End Function


'Pivot
Public Function EssMenuVPivot() As Long
    EssMenuVPivot = EssbaseVBAFunctionWrapper.EssMenuVPivot(Excel.application.Hwnd)
End Function

Public Function EssVPivot(ByVal sheetName As Variant, ByVal range As Variant, ByVal startPoint As Variant, ByVal endPoint As Variant) As Long
    EssVPivot = EssbaseVBAFunctionWrapper.EssVPivot(sheetName, range, startPoint, endPoint, Excel.application.Hwnd)
End Function


'SendData
Public Function EssMenuVSend() As Long
    EssMenuVSend = EssbaseVBAFunctionWrapper.EssMenuVSend(Excel.application.Hwnd)
End Function

Public Function EssVSendData(ByVal sheetName As Variant, ByVal range As Variant) As Long
    EssVSendData = EssbaseVBAFunctionWrapper.EssVSendData(sheetName, range, Excel.application.Hwnd)
End Function


'Lock
Public Function EssMenuVRetrieveLock() As Long
    EssMenuVRetrieveLock = EssbaseVBAFunctionWrapper.EssMenuVRetrieveLock(Excel.application.Hwnd)
End Function

Public Function EssMenuVLock() As Long
    EssMenuVLock = EssbaseVBAFunctionWrapper.EssMenuVLock(Excel.application.Hwnd)
End Function


'Unlock
Public Function EssMenuVUnlock() As Long
    EssMenuVUnlock = EssbaseVBAFunctionWrapper.EssMenuVUnlock(Excel.application.Hwnd)
End Function

Public Function EssVUnlock(ByVal sheetName As Variant) As Long
    EssVUnlock = EssbaseVBAFunctionWrapper.EssVUnlock(sheetName, Excel.application.Hwnd)
End Function


'Global Options
Public Function EssVGetGlobalOption(ByVal item As Long) As Variant
    EssVGetGlobalOption = EssbaseVBAFunctionWrapper.EssVGetGlobalOption(item, Excel.application.Hwnd)
End Function

Public Function EssVSetGlobalOption(ByVal item As Long, ByVal globalOption As Variant) As Long
    EssVSetGlobalOption = EssbaseVBAFunctionWrapper.EssVSetGlobalOption(item, globalOption, Excel.application.Hwnd)
End Function


'Menu Options
Public Function EssMenuVOptions() As Long
    EssMenuVOptions = EssbaseVBAFunctionWrapper.EssMenuVOptions(Excel.application.Hwnd)
End Function


'Sheet Options
Public Function EssVGetSheetOption(ByVal sheetName As Variant, ByVal item As Variant) As Variant
    EssVGetSheetOption = EssbaseVBAFunctionWrapper.EssVGetSheetOption(sheetName, item, Excel.application.Hwnd)
End Function

Public Function EssVSetSheetOption(ByVal sheetName As Variant, ByVal item As Variant, ByVal sheetOption As Variant) As Long
    EssVSetSheetOption = EssbaseVBAFunctionWrapper.EssVSetSheetOption(sheetName, item, sheetOption, Excel.application.Hwnd)
End Function


'Calculations
Public Function EssMenuVCalculation() As Long
    EssMenuVCalculation = EssbaseVBAFunctionWrapper.EssMenuVCalculation(Excel.application.Hwnd)
End Function

Public Function EssVCalculate(ByVal sheetName As Variant, ByVal calcScript As Variant, ByVal synchronous As Variant) As Long
    EssVCalculate = EssbaseVBAFunctionWrapper.EssVCalculate(sheetName, calcScript, synchronous, Excel.application.Hwnd)
End Function

Public Function EssVCancelCalc(ByVal sheetName As Variant) As Long
    EssVCancelCalc = EssbaseVBAFunctionWrapper.EssVCancelCalc(sheetName, Excel.application.Hwnd)
End Function


'Cascade
Public Function EssMenuVCascade() As Long
    EssMenuVCascade = EssbaseVBAFunctionWrapper.EssMenuVCascade(Excel.application.Hwnd)
End Function

'Public Function EssVCascade() As Long
'    EssVCascade = EssbaseVBAFunctionWrapper.EssVCascade(sheetName, range, selection, Path, Prefix, suffix, level, openFile, copyFormats, overwrite, listFile, Excel.application.Hwnd)
'End Function


'Cell
Public Function EssVCell(ByVal sheetName As Variant, ParamArray memberList() As Variant) As Variant
    Dim memberListArgs As Variant
    
    'Copy the ParamArray to a variant
    memberListArgs = memberList
    
    'Get the value
    EssVCell = EssbaseVBAFunctionWrapper.EssVCell(sheetName, Excel.application.Hwnd, memberListArgs)
End Function


'FlashBack
Public Function EssMenuVFlashBack() As Long
    EssMenuVFlashBack = EssbaseVBAFunctionWrapper.EssMenuVFlashBack(Excel.application.Hwnd)
End Function

Public Function EssVFlashBack() As Long
    EssVFlashBack = EssbaseVBAFunctionWrapper.EssVFlashBack(sheetName, Excel.application.Hwnd)
End Function


'DataPoint
Public Function EssVGetDataPoint(ByVal sheetName As Variant, ByVal cell As Variant, ByVal range As Variant, ByVal aliases As Variant) As Variant
    EssVGetDataPoint = EssbaseVBAFunctionWrapper.EssVGetDataPoint(sheetName, cell, range, aliases, Excel.application.Hwnd)
End Function

Public Function EssVFreeDataPoint(ByRef Info As Variant) As Long
    EssVFreeDataPoint = EssbaseVBAFunctionWrapper.EssVFreeDataPoint(Info, Excel.application.Hwnd)
End Function


'MemberInfo
Public Function EssVGetMemberInfo(ByVal sheetName As Variant, ByVal mbrName As Variant, ByVal action As Variant, ByVal aliases As Variant) As Variant
    EssVGetMemberInfo = EssbaseVBAFunctionWrapper.EssVGetMemberInfo(sheetName, mbrName, action, aliases, Excel.application.Hwnd)
End Function

Public Function EssVFreeMemberInfo(ByRef memInfo As Variant) As Long
    EssVFreeMemberInfo = EssbaseVBAFunctionWrapper.EssVFreeMemberInfo(memInfo, Excel.application.Hwnd)
End Function


'MemberSelection
Public Function EssMenuVMemberSelection() As Long
    EssMenuVMemberSelection = EssbaseVBAFunctionWrapper.EssMenuVMemberSelection(Excel.application.Hwnd)
End Function


'DatalessNav
Public Function EssMenuVDatalessNav() As Long
    EssMenuVDatalessNav = EssbaseVBAFunctionWrapper.EssMenuVDatalessNav(Excel.application.Hwnd)
End Function


'Currency
Public Function EssMenuVCurrencyReport() As Long
    EssMenuVCurrencyReport = EssbaseVBAFunctionWrapper.EssMenuVCurrencyReport(Excel.application.Hwnd)
End Function

Public Function EssVGetCurrency(ByVal sheetName As Variant) As Variant
    EssVGetCurrency = EssbaseVBAFunctionWrapper.EssVGetCurrency(sheetName, Excel.application.Hwnd)
End Function

Public Function EssVSetCurrency(ByVal sheetName As Variant, ByVal currencyIdentifier As Variant) As Long
    EssVSetCurrency = EssbaseVBAFunctionWrapper.EssVSetCurrency(sheetName, currencyIdentifier, Excel.application.Hwnd)
End Function


'Styles
Public Function EssVGetStyle(ByVal sheetName As Variant, ByVal styleType As Variant, ByVal dimName As Variant, ByVal item As Long) As Variant
    EssVGetStyle = EssbaseVBAFunctionWrapper.EssVGetStyle(sheetName, styleType, dimName, item, Excel.application.Hwnd)
End Function

Public Function EssVSetStyle(ByVal sheetName As Variant, ByVal styleType As Variant, ByVal dimName As Variant, ByVal item As Long, ByVal newValue As Variant) As Long
    EssVSetStyle = EssbaseVBAFunctionWrapper.EssVSetStyle(sheetName, styleType, dimName, item, newValue, Excel.application.Hwnd)
End Function


'Hctx
Public Function EssVGetHctxFromSheet(ByVal sheetName As Variant) As Long
    EssVGetHctxFromSheet = EssbaseVBAFunctionWrapper.EssVGetHctxFromSheet(sheetName, Excel.application.Hwnd)
End Function


'LoginSetPassword
Public Function EssVLoginSetPassword(ByVal sheetName As Variant, ByVal newPassword As Variant, ByVal oldPassword As Variant, ByVal server As Variant, ByVal userName As Variant) As Long
    EssVLoginSetPassword = EssbaseVBAFunctionWrapper.EssVLoginSetPassword(sheetName, newPassword, oldPassword, server, userName, Excel.application.Hwnd)
End Function


'SetMenu
Public Function EssVSetMenu(ByVal setMenu As Boolean) As Long
    EssVSetMenu = EssbaseVBAFunctionWrapper.EssVSetMenu(setMenu, Excel.application.Hwnd)
End Function


'SetSKUMouse
Public Function EssVSetSKUMouse(ByVal fSKUWithMouse As Variant) As Long
    EssVSetSKUMouse = EssbaseVBAFunctionWrapper.EssVSetSKUMouse(fSKUWithMouse, Excel.application.Hwnd)
End Function


'LaunchTableau
Public Function EssMenuVLaunchTableau() As Long
    EssMenuVLaunchTableau = EssbaseVBAFunctionWrapper.EssMenuVLaunchTableau(Excel.application.Hwnd)
End Function


'LinkedObjects
Public Function EssMenuVLinkedObjects() As Long
    EssMenuVLinkedObjects = EssbaseVBAFunctionWrapper.EssMenuVLinkedObjects(Excel.application.Hwnd)
End Function


'QueryDesigner
Public Function EssMenuVQueryDesigner() As Long
    EssMenuVQueryDesigner = EssbaseVBAFunctionWrapper.EssMenuVQueryDesigner(Excel.application.Hwnd)
End Function


'RetrievalWizard
Public Function EssMenuVWizard() As Long
    EssMenuVWizard = EssbaseVBAFunctionWrapper.EssMenuVWizard(Excel.application.Hwnd)
End Function

''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''
' Utility functions not in the original Essbase add-in code
''''''''''''''''''''''''''''''''''''''''''''''''''''''''''''

'Returns the Essbase sheetname argument based on the worksheet object passed in.
'If the worksheet is null (ie Nothing), the ActiveSheet is used.
Public Function GetEssbaseSheetName(Optional oSheet As Worksheet) As String
    'Default to the activesheet
    If oSheet Is Nothing Then
        Set oSheet = ActiveSheet
    End If
    
    'Return the sheet name
    GetEssbaseSheetName = "[" & oSheet.Parent.Name & "]" & oSheet.Name
End Function


'Returns the Essbase error code from the previous call.
Public Function EssVGetLastErrorCode() As Long
    EssVGetLastErrorCode = EssbaseVBAFunctionWrapper.EssVGetLastErrorCode()
End Function


'Returns the Essbase error message from the previous call.
Public Function EssVGetLastErrorMessage() As String
    EssVGetLastErrorMessage = EssbaseVBAFunctionWrapper.EssVGetLastErrorMessage()
End Function


' Returns an array of servers, applications, and databases available to the user.  The list
' may be optionally filtered by application name, and database name.  If there is an Essbase error,
' the appropriate error code and error message are returned by the EssVGetLastErrorCode and
' EssVGetLastErrorMessage functions.
'
' The userName and password values may not be Null or Empty for this function.
Public Function EssVListDatabases(ByVal userName As Variant, ByVal password, ByVal serverName As Variant, ByVal applicationName As Variant, ByVal databaseName As Variant) As Variant
    EssVListDatabases = EssbaseVBAFunctionWrapper.EssVListDatabases(userName, password, serverName, applicationName, databaseName, APS_URL, DODECA_ESSBASE_URL, Excel.application.Hwnd)
End Function

'''''''''''''''''''''''''''''''''''
' VarianceTracker code begins here
'''''''''''''''''''''''''''''''''''

Sub RetrieveData()

'This is the master routine; it gets called from the buttons on the first two sheets in the workbook
'It refreshes either all of the "Before" sheets, or all of the "After" sheets (depending on the button)
'When the "After" sheets are refreshed, the new numbers are compared to the corresponding "Before" numbers
'All cells with variances on "After" sheets are marked yellow
'Also, a list of every column on the sheet that contains at least one variance cell is written to cell A2 on the "After" sheets

On Error GoTo ErrorHandler

    Dim i As Integer
    Dim SheetType As String
    Dim RetrievalFailure As Boolean
    Dim EssServer As String
    Dim check_dr As Variant
    Dim check_dr2 As Variant
    Dim ret As Variant
    Dim cube As Variant
    Dim cube_name As String
    Dim app_name As String
    
    application.ScreenUpdating = True
    
    'Capture the sheet type and specify the Essbase server it will pull from
    If UCase(ActiveSheet.Name) Like "*BEFORE" Then
        SheetType = "Before"
        EssServer = EssServer_Before
    Else
        SheetType = "After"
        EssServer = EssServer_After
    End If
    
    'Verify the intent to refresh all sheets
    ret = MsgBox("Are you sure you want to refresh all of the " & SheetType & " sheets?", vbYesNo + vbQuestion)
    If ret = vbNo Then
        MsgBox ("Operation cancelled")
        Exit Sub
    End If
    
    'Prompt to check Dept Reporting vs CC Reporting totals
    check_dr = MsgBox("Do you want to verify Dept Reporting vs. CC Reporting?", vbYesNo + vbQuestion)
    If check_dr = vbYes Then
        check_dr2 = MsgBox("Are you sure you want to verify Dept Reporting vs. CC Reporting?", vbYesNo + vbQuestion)
    End If
    
    For Each cube In CubeList.Keys
        
        cube_name = cube
        app_name = CubeList(cube)
        
        Sheets(cube_name & " " & SheetType).Select
        
        range("a1").Select
        range("a1").ClearContents
        With range("a2")
            .Value = ""
            .Interior.ColorIndex = xlNone
            .Font.ColorIndex = 1
            .Font.Bold = False
        End With

        'Clear all existing data values from the sheet
        If ActiveSheet.Name Like "CAP*" Or ActiveSheet.Name Like "CASHFLOW*" Then
            range(Cells(first_data_row, first_data_column_capital), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count)).ClearContents
            range(Cells(first_data_row, first_data_column_capital), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count)).Interior.ColorIndex = xlNone
        Else
            range(Cells(first_data_row, first_data_column), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count)).ClearContents
            range(Cells(first_data_row, first_data_column), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count)).Interior.ColorIndex = xlNone
        End If

        'Retrieve the new data
        If PerformEssbaseOperation(EssServer, app_name, cube_name, "Retrieve") = True Then
            range("a1").Value = "Data pulled from " & cube_name & " on " & EssServer & " at " & Now()
        Else
            ret = MsgBox("Data retrieval on " & ActiveSheet.Name & " failed.", vbExclamation)
            RetrievalFailure = True
        End If

        If SheetType = "After" Then
            'Mark any variances after a comparison with the "Before" sheet
            If HighlightVariances = True Then
                With range("a2")
                    .Value = "Variances found"
                    .Font.ColorIndex = 3
                    .Font.Bold = True
                End With
                
                'Write into cell A2 a list of the columns that contain variances
                Call ListVarianceColumns
            
            Else
                With range("a2")
                    .Value = "No variances found"
                    .Interior.ColorIndex = 4
                End With
            End If
        End If
        
        If RetrievalFailure = True Then Exit For
    
    Next
    
    
    If RetrievalFailure = True Then
        ret = MsgBox("Data retrievals incomplete.", vbExclamation)
    Else
        'Check net income and headcount between Dept Reporting and CC Reporting
        'The sheet contains conditional formatted to show any variances between those two hierarchies
        If check_dr2 = vbYes Then
            CheckDeptReporting
        End If

        ret = MsgBox("Data retrievals complete.")
    End If
    
    application.Calculate

Exit Sub
ErrorHandler:
    MsgBox ("A critical error occurred: " & Err.Description)
End Sub

Function HighlightVariances() As Boolean

'This compares values on the "After" sheets with the corresponding cells on the "Before" sheets
'Also, a list of every column on the "After" sheet that contains at least one variance cell is written to cell A2

    Dim cell As range
    Dim data_range As range
    Dim Before_Sheet As Worksheet
    
    HighlightVariances = False
    
    Set Before_Sheet = ActiveSheet.Previous
    If Not (Before_Sheet.Name Like "*Before" Or Before_Sheet.Name Like "*After") Then
        MsgBox ("Unable to locate appropriate Before sheet.")
        Exit Function
    End If
    
    If Left(ActiveSheet.Name, 3) = "CAP" Then
        Set data_range = range(Cells(first_data_row, first_data_column_capital), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count))
    Else
        Set data_range = range(Cells(first_data_row, first_data_column), Cells(ActiveSheet.UsedRange.Rows.Count, ActiveSheet.UsedRange.Columns.Count))
    End If
        
    For Each cell In data_range.Cells
        If Trim(Round(cell.Value, 0)) = Trim(Round(Before_Sheet.Cells(cell.Row, cell.Column).Value, 0)) Then
            'cell.Interior.ColorIndex = 4
        Else
            cell.Interior.ColorIndex = 6
            HighlightVariances = True
        End If
    Next cell

    Set data_range = Nothing
    
End Function

Sub ListVarianceColumns()

'A list of every column on the "After" sheet that contains at least one variance cell is written to cell A2

    Dim cell As range
    Dim col As Long
    Dim monthly_yellow_cells_found As Boolean
    Dim yellow_cells As String
    
    'Initialize
    monthly_yellow_cells_found = False
    
    'Loop through the columns and identify the ones containing variances (i.e., yellow cells)
    For col = 1 To ActiveSheet.UsedRange.Columns.Count
        For Each cell In ActiveSheet.UsedRange.Columns(col).Cells
            If cell.Interior.ColorIndex = 6 Then
                If (Cells(4, col).Value = "Working" And Cells(5, col).Value = "Forecast") Or Cells(5, col).Value = "FlashNonGAAP" Then
                    Select Case Cells(7, col).Value
                        Case "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
                            'Update the variable, but don't flag the by-month columns
                            monthly_yellow_cells_found = True
                        Case Else
                            yellow_cells = yellow_cells & Col_Letter(col) & ", "
                            Exit For
                    End Select
                Else
                    yellow_cells = yellow_cells & Col_Letter(col) & ", "
                    Exit For
                End If
            End If
        Next cell
    Next col
    
    'Clean up the string of column references
    If Len(yellow_cells) > 0 Then
        
        yellow_cells = Trim(yellow_cells)
        yellow_cells = Left(yellow_cells, (Len(yellow_cells) - 1))
    
        'Write the list of variance columns into cell A2
        If InStr(1, yellow_cells, ",") Then
            range("A2").Value = "Variances found in columns " & yellow_cells
        Else
            range("A2").Value = "Variances found in column " & yellow_cells
        End If
    
    Else
        
        If monthly_yellow_cells_found = True Then
        'Variances were found on the sheet, but not in the qualifying columns - issue a warning
            range("A2").Value = "WARNING: Variances found in monthly columns, but not in YearTotal columns."
        End If
        
    End If

End Sub

Function Col_Letter(lngCol As Long) As String
'Get the column letter associated with the column number
    Dim vArr
    vArr = Split(Cells(1, lngCol).Address(True, False), "$")
    Col_Letter = vArr(0)
End Function

Sub CheckDeptReporting()

'Optional routine that refreshes the data on the "Dept Reporting Check" sheet
'NOTE: The retrieval time on that sheet is about three minutes (because it contains dozens of dynamic calc members)

    Dim ret As Variant
    Dim RetrievalFailure As Boolean
    
    Sheets("Dept Reporting Check").Select
    range("a1").Select
    range("a1").ClearContents
        
    'Clear all existing data values
    range("dept_rptg_values").ClearContents
        
    'Retrieve the new data
    If PerformEssbaseOperation(EssServer_DR, "SWA_RPT", "SWA_RPT", "Retrieve") = True Then
        range("a1").Value = "Data pulled from " & "SWA_RPT" & " on " & EssServer_DR & " at " & Now()
    Else
        ret = MsgBox("Data retrieval on " & ActiveSheet.Name & " failed.", vbExclamation)
        RetrievalFailure = True
    End If

End Sub

Function PerformEssbaseOperation(EssServer As String, EssApplication As String, EssDatabase As String, OperationType As String, Optional ZoomLevel As Integer, Optional SuppressMissing As Boolean, Optional SheetRange, Optional ReportName As String, Optional CalcName As String) As Boolean

'This routine connects the active worksheet to Essbase and refreshes it with the latest data

'The possible values for OperationType are: "Retrieve", "Zoom", "Upload", "RunScript", and "RetrieveNewRole"
'The possible values for ZoomLevel are: 1 (Next Level), 2(All Levels), and 3 (Bottom Level)

On Error GoTo ErrorHandler

    Dim first_line_of_script As String
    Dim error_message As String
    Dim PauseTime As Integer
    Dim login_ret As Long
    Dim ret As Long
    Dim lock_ret As Long
    Dim disconnect_ret As Long
    Dim Timer As Variant
    Dim Start As Variant
    Dim Msg As Variant
    Dim essServerName As String
    Dim essAppName As String
    Dim essDbName As String
    Dim login_attempts As Integer
    Dim EssbaseOperationFailed As Boolean
    Dim EssbaseFailureType As Integer

    'Ensure these variables are always reset before any Essbase operations are attempted
    EssbaseOperationFailed = False
    EssbaseFailureType = 0
    
    'essServerName = EssServer
    'essAppName = EssApplication
    'essDbName = EssDatabase
    
    If essConnect(EssServer, EssApplication, EssDatabase) = True Then
        
        'Unprotect the sheet
        ActiveSheet.Unprotect
             
        If IsMissing(SheetRange) Then
            Select Case OperationType
            Case "Retrieve", "RunScript"
                ret = EssVRetrieve(Empty, Empty, 1)
            Case "Zoom"
                ret = EssVZoomIn(Empty, Empty, Empty, ZoomLevel, False)
            Case "Upload"
                lock_ret = EssVRetrieve(Empty, Empty, 3) 'Lock the blocks that will receive the data
                If lock_ret <> 0 Then
                    ret = 99  'Forcibly set ret to a code that indicates a failed lock attempt
                    'NOTE: 99 is an arbitrary value used to indicate a lock failure, it is not an Essbase-issued value
                    'The appropriate failure message will be issued below
                Else
                    ret = EssVSendData(Empty, Empty) 'Uploads the data and unlocks the blocks
                End If
            Case "Calculate"
                ret = EssVCalculate(Empty, CalcName, True)
            End Select
        Else
            Select Case OperationType
            Case "Retrieve", "RunScript"
                ret = EssVRetrieve(Empty, SheetRange, 1)
            Case "Zoom"
                ret = EssVZoomIn(Empty, SheetRange, Empty, ZoomLevel, False)
            Case "Upload"
                lock_ret = EssVRetrieve(Empty, Empty, 3) 'Lock the blocks that will receive the data
                If lock_ret <> 0 Then
                    ret = 99  'Forcibly set ret to a code that indicates a failed lock attempt
                    'NOTE: 99 is an arbitrary value used to indicate a lock failure, it is not an Essbase-issued value
                    'The appropriate failure message will be issued below
                Else
                    ret = EssVSendData(Empty, Empty) 'Uploads the data and unlocks the blocks
                End If
            Case "Calculate"
                ret = EssVCalculate(Empty, CalcName, True)
            End Select
        End If
        
        'Check the return code to determine if the operation was successful
        If ret <> 0 Then
            If ret = 1020010 Then
                'This "error" merely means the zoom generated no data--ignore it
            Else
                'This catches all general failures (i.e., where ret <> 0)
                EssbaseFailureType = 2
            
                'Update variable
                PerformEssbaseOperation = False
            
                application.ScreenUpdating = True
            End If
        
        Else  'The Essbase operation was successful
        
            'Do not issue a confirmation message at this point
            PerformEssbaseOperation = True
            
        End If
    
    Else  'The attempt to connect to Essbase failed
        
        application.ScreenUpdating = True
        
        'Update variables
        EssbaseFailureType = 1
        PerformEssbaseOperation = False
        
    End If
    
    'Disconnect the sheet
    ret = EssVDisconnect(Empty)
    
    'Calculate the workbook to refresh all Excel formulas and links
    application.Calculate

Exit Function
ErrorHandler:
    'Call DisplayErrorMessage
End Function

Function essConnect(essServerName As String, essAppName As String, essDbName As String) As Boolean

'Prompt the user for eID and password, then connect the active sheet

On Error GoTo ErrorHandler
    
    application.ScreenUpdating = False
        
    Dim EssUserName As String
    Dim EssPassword As String
    Dim x As Variant
    Dim hCtx As Variant
    Dim ret As Variant
   
    If frmLogin.txtUserName = "" Then
        application.ScreenUpdating = True
        frmLogin.Show
        application.ScreenUpdating = False
    End If
    
    EssUserName = frmLogin.txtUserName
    EssPassword = frmLogin.txtPassword
    
    ret = EssVConnect(Empty, EssUserName, EssPassword, essServerName, essAppName, essDbName)
    
    If ret = 0 Then
        essConnect = True
    Else
        essConnect = False
        application.DisplayAlerts = True
        MsgBox ("Essbase connection failed. Error: " + Str(x))
    End If

Exit Function
ErrorHandler:
    MsgBox ("Error occurred during connection attempt.")
    essConnect = False
End Function

Sub CopyAfterToBefore()

'This is an ad-hoc routine that copies data values on all "After" sheets to the corresponding "Before" sheets

    Dim resp As Variant
    
    resp = MsgBox("Are you sure you want to sync the 'Before' sheets with 'After'?", vbYesNo + vbDefaultButton2)
    If resp = vbNo Then Exit Sub
    
    application.ScreenUpdating = False
    
    Sheets("FIN_STMT After").Select
    Cells.Copy
    Sheets("FIN_STMT Before").Select
    Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    range("A2").ClearContents
    range("A1").Select

    Sheets("CAPITAL After").Select
    Cells.Copy
    Sheets("CAPITAL Before").Select
    Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    range("A2").ClearContents
    range("A1").Select

    Sheets("SWA_RPT After").Select
    Cells.Copy
    Sheets("SWA_RPT Before").Select
    Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    range("A2").ClearContents
    range("A1").Select

    Sheets("CAPT_RPT After").Select
    Cells.Copy
    Sheets("CAPT_RPT Before").Select
    Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    range("A2").ClearContents
    range("A1").Select

    Sheets("CASHFLOW After").Select
    Cells.Copy
    Sheets("CASHFLOW Before").Select
    Cells.PasteSpecial Paste:=xlPasteValues, Operation:=xlNone, SkipBlanks:=False, Transpose:=False
    range("A2").ClearContents
    range("A1").Select
    
    Sheets("FIN_STMT Before").Select
    resp = MsgBox("Sheets reset.", vbOKOnly)
    
    application.Calculate

End Sub

Sub ValidateMonthRoll()

'Run this routine at the end of the month-roll process to ensure all of the data ended up in the correct places:
'1. Forecast/Working values from the month that was just closed are copied to Forecast/Final
'2. Actual/Final values from the month that was just closed are copied to Forecast/Working

    Dim month_name As String
    Dim year_num As Variant
    Dim sheet As Worksheet
    Dim col_num As Integer
    Dim column_def As Variant
    Dim final_forecast_col As Integer
    Dim working_forecast_col As Integer
    Dim final_actual_col As Integer
    Dim cell As range
    
    'Prompt for the name of the month that was just closed
    month_name = InputBox("Which month was just closed?")
    
    'Validate the month name that was entered
    Select Case month_name
    Case "Jan", "Feb", "Mar", "Apr", "May", "Jun", "Jul", "Aug", "Sep", "Oct", "Nov", "Dec"
        'Name is correct as-is
    Case "January", "February", "March", "April", "May", "June", "July", "August", "September", "October", "November", "December"
        'Keep the left three characters only
        month_name = Left(month_name, 3)
    Case Else
        MsgBox ("Invalid month name entered")
        Exit Sub
    End Select
    
    'Set the year member
    If Not month_name = "Dec" Then
        year_num = "FY" & Right(Year(Now()), 2)
    Else
        year_num = "FY" & (Right(Year(Now()), 2) - 1)
    End If
    
    'Process all of the "After" sheets (except the one for the Cashflow cube)
    For Each sheet In ThisWorkbook.Sheets
    
        If sheet.Name Like "* After" And Not sheet.Name Like "CASHFLOW*" Then
            
            sheet.Activate
        
            'Find the specified month's Final Forecast column
            For col_num = 1 To ActiveSheet.UsedRange.Columns.Count
                column_def = Cells(4, col_num) & Cells(5, col_num) & Cells(6, col_num) & Cells(7, col_num)
                
                If column_def = "FinalForecast" & year_num & month_name Then
                    final_forecast_col = col_num
                    MsgBox ("Final Forecast is in column " & final_forecast_col)
                    Exit For
                End If
            Next col_num
            
            'Find the specified month's Working Forecast column
            For col_num = 1 To ActiveSheet.UsedRange.Columns.Count
                column_def = Cells(4, col_num) & Cells(5, col_num) & Cells(6, col_num) & Cells(7, col_num)
                
                If column_def = "WorkingForecast" & year_num & month_name Then
                    working_forecast_col = col_num
                    MsgBox ("Working Forecast is in column " & working_forecast_col)
                    Exit For
                End If
            Next col_num
            
            'Find the specified month's Final Actual column
            For col_num = 1 To ActiveSheet.UsedRange.Columns.Count
                column_def = Cells(4, col_num) & Cells(5, col_num) & Cells(6, col_num) & Cells(7, col_num)
                
                If column_def = "FinalActual" & year_num & month_name Then
                    final_actual_col = col_num
                    MsgBox ("Final Actual is in column " & final_actual_col)
                    Exit For
                End If
            Next col_num
                
                
            'Now color-code the values in the Final Forecast and Working Forecast columns
            'Note: The values in the selected month's Final Forecast column must be compared with the values in the Working Forecast column on the "Before" sheet
            ActiveSheet.UsedRange.Columns(final_forecast_col).EntireColumn.Interior.ColorIndex = xlNone
            For Each cell In ActiveSheet.UsedRange.Columns(final_forecast_col).Cells
                If cell.Row >= 8 Then
                    If cell.Value <> "" Then
                        If cell.Value = ActiveSheet.Previous.Cells(cell.Row, working_forecast_col).Value Then
                            cell.Interior.ColorIndex = 4
                        Else
                            cell.Interior.ColorIndex = 3
                        End If
                    End If
                End If
            Next cell
            
            ActiveSheet.UsedRange.Columns(working_forecast_col).EntireColumn.Interior.ColorIndex = xlNone
            For Each cell In ActiveSheet.UsedRange.Columns(working_forecast_col).Cells
                If cell.Row >= 8 Then
                    If cell.Value <> "" Then
                        If cell.Value = Cells(cell.Row, final_actual_col).Value Then
                            cell.Interior.ColorIndex = 4
                        Else
                            cell.Interior.ColorIndex = 3
                        End If
                    End If
                End If
            Next cell
        
        End If
    
    Next sheet
    
    
    MsgBox ("Validation complete")

End Sub

Sub IncrementYearHeaders()

'Run this routine at the beginning of every new year to increment the year numbers in the column headers

    Dim i As Integer
    Dim sheet As Worksheet

    For Each sheet In ThisWorkbook.Sheets
        If sheet.Name Like "* Before" Or sheet.Name Like "* After" Then
            For i = (Right(Trim(Year(Now)), 2) + 4) To 15 Step -1
                sheet.Cells.Replace What:="FY" & i, Replacement:="FY" & i + 1, LookAt:=xlPart, SearchOrder:=xlByRows, MatchCase:=False, SearchFormat:=False, ReplaceFormat:=False
            Next
        End If
    Next

End Sub

Sub HighlightVariances_Manual()

'This allows the active sheet (only) to be checked for variances without refreshing its data
'If an "After" sheet is active, it will compare against the "Before"
'If a "Before" sheet is active, it will compare against the "After"

    If HighlightVariances = True Then
        With range("a2")
            .Value = "Variances found"
            .Interior.ColorIndex = 3
        End With
    Else
        With range("a2")
            .Value = "No variances found"
            .Interior.ColorIndex = 4
        End With
    End If

End Sub

