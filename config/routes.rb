'*************************************************************************************************
' @File         classDataEntry.qfl

'@Author:       Giacomo Kavanagh
'@Date:         13/12/2010
'@Objective:    Manage populating pages
'@Prereqs:      
'@Resources:
'@Notes:

'@History:
'@Version   Date        Author              ChangeID    Description
'@0.1 GK    13/12/2010   Giacomo Kavanagh    0.1 GK      Initial functions changed to class format with full error handling
'*************************************************************************************************

'*************************************************************************************************
'@Name          fnctInitializeClsDataEntry
'@Description   Initialises the Events class
'*************************************************************************************************
Function fnctInitializeClsDataEntry
    Set fnctInitializeClsDataEntry = New clsDataEntry
    fnctInitializeClsDataEntry.onfnctGetObjName.addHandler array( "logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    fnctInitializeClsDataEntry.onfnctEnterData.addHandler array("fnctReportDataEntrySuccess", "logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    fnctInitializeClsDataEntry.onDataCheck.addHandler array("fnctReportDataEntrySuccess", "logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    fnctInitializeClsDataEntry.onfnctCheckData.addHandler array( "logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
End Function

Class clsDataEntry
    Private int_Count, arr_ObjectOnPageNames, arr_Values, str_CurrentMethod, obj_Page, _
       str_State
    Public strObjectOnPageName, strValue, strActualValue, strActualState, blnReadOnlyValue
    Public onfnctGetObjName, onfnctEnterData, onfnctCheckData, onDataCheck

    Private Sub Class_Initialize
        Set onfnctGetObjName = New clsHandlers
        Set onfnctEnterData = New clsHandlers
        Set onfnctCheckData = New clsHandlers
        Set onDataCheck = New clsHandlers
        Set obj_Page = Description.Create()
    End Sub

    Public Property Get properties ()
        set dictProperties = createObject ("scripting.dictionary")
        dictProperties.add "method", str_CurrentMethod
        dictProperties.add "Value", strValue
        dictProperties.add "strActualValue", strActualValue
        dictProperties.add "ObjectOnPageName", strObjectOnPageName
        set properties = dictProperties
    End Property

    Public Sub DataEntry(rs, objPage)
        str_CurrentMethod = "DataEntry"
        on error resume next
        Set obj_Page = objPage
        int_Count = rs.fields.count - 1

        Call fnctGetObjName(rs)
        Call fnctRemoveNullValues(rs)
        For i = 0 to int_Count
            strValue = arr_Values(i)
            strObjectOnPageName = arr_ObjectOnPageNames(i)
            If strObjectOnPageName <> "ID" And InStr(strObjectOnPageName, ".ID") = 0  Then
                Call fnctEnterData
            End If
        Next
    End Sub

    Public Sub DataCheck(rs, objPage, strState)
        str_CurrentMethod = "DataCheck"
        on error resume next
        Set obj_Page = objPage
        str_State = strState
        int_Count = rs.fields.count - 1

        Call fnctGetObjName(rs)
        Call fnctRemoveNullValues(rs)

        For i = 0 to int_Count
            strValue = arr_Values(i)
            strObjectOnPageName = arr_ObjectOnPageNames(i)
            If strObjectOnPageName <> "ID" And InStr(strObjectOnPageName, ".ID") = 0 Then
                Call fnctCheckData
    
                '@Check the retrieved values
                blnResult = fnctCheckValues(strActualValue, strValue)
                If Not(blnResult) Then
                    Err.Raise 2, , "strActualValue <" & strActualValue & "> Not matched strExpectedValue <" & strValue & "> For object of name <" & strObjectOnPageName & ">", ""
                End If
                onDataCheck.fire Me
            End If

            If str_State = "ReadOnly" And strObjectOnPageName <> "ID" And InStr(strObjectOnPageName, ".ID") = 0 Then
                '@Check the fields are read only (which is largely wildcat specific
                blnReadOnlyResult = fnctCheckValues(blnReadOnlyValue, 1)
                If Not(blnReadOnlyResult) Then
                    Err.Raise 2, , "Expecting Read only field for object name <" & strObjectOnPageName & ">, received actual result of <" & blnReadOnlyValue & ">", ""
                End If
                onDataCheck.fire Me
            End If
        Next
    End Sub

    Private Sub fnctGetObjName(rs)
        '@These functions will ignore fields called ID
        Dim strCharacter, intCharacters
        str_CurrentMethod = "fnctGetObjName"

        on error resume next

        ReDim arr_ObjectOnPageNames (int_Count)

        For i = 0 to int_Count
            strInput = rs.Fields(i).Name

            '@Record the initial length of the string to search
            intCharacters = Len(strInput)
            For j = 0 to intCharacters
                strCharacter = Left(strInput,1)

                If strCharacter = LCase(strCharacter) And Not(IsNumeric(strCharacter)) Then
                    strInput = Right(strInput,Len(strInput)-1)
                ElseIf (strCharacter = LCase(strCharacter) And IsNumeric(strCharacter)) Or strCharacter <> LCase(strCharacter) Then
                    arr_ObjectOnPageNames(i) = strInput
                    Exit For
                End If
            Next
            If j = intCharacters Then
                Err.Raise 2, , "No upper case Letter found in string, no object name will be given for object value <" & arr_ObjectOnPageNames(i) & ">"
            End If

            onfnctGetObjName.fire Me
        Next
    End Sub

    Private Sub fnctRemoveNullValues(rs)
        '@These functions will ignore fields called ID
        str_CurrentMethod = "fnctRemoveNullValues"
        on error resume next

        ReDim arr_Values (int_Count)

        For i = 0 to int_Count

            If IsNull(rs.fields(i).value) Or IsEmpty(rs.fields(i).value) Then
                arr_Values(i) = ""
            Else
                arr_Values(i) = rs.fields(i).value
            End If
        Next
    End Sub

    Private Sub fnctEnterData
        str_CurrentMethod = "fnctEnterData"
        On Error Resume Next
    
        Dim intTempReporterFilterSetting, blnResult
        '@Switch off the reporting as we try various objects until one works, and error handling on
        
        intTempReporterFilterSetting = Reporter.Filter
        Reporter.Filter = rfDisableAll
        With obj_Page
            .WebEdit(strObjectOnPageName).Set strValue
            If Err.Number <> 0 Then
                Err.Clear
                .WebList(strObjectOnPageName).Select strValue
            End If
            If Err.Number <> 0 Then
                Err.Clear
                .WebRadioGroup(strObjectOnPageName).Select strValue
            End If
            
            If Err.Number <> 0 Then
                Err.Clear
                .WebFile(strObjectOnPageName).Set strValue
            End If
    
            If Err.Number <> 0 Then
                Err.Clear
                If strValue = "True" or strValue = 1 or strValue = -1 or strValue = "1" or strValue = "-1" Then
                    strValue = "ON"
                ElseIf strValue = "False" or strValue = 0 or strValue = "0" Then
                    strValue = "OFF"
                End If
                .WebCheckBox(strObjectOnPageName).Set strValue
            End If
            If Err.Number = 0 Then
                blnResult = true
            Else
                Err.Raise 2, , "Could not find object <" & strObjectOnPageName & "> And set value <" & strValue & ">"
            End If
        End With
        Reporter.Filter = intTempReporterFilterSetting

        onfnctEnterData.fire Me
    End Sub

    Private Sub fnctCheckData
        Dim intTempReporterFilterSetting
        Dim blnReadOnlyCheckPerformed
        '@Switch off the reporting as we try various objects until one works, and error handling on

        str_CurrentMethod = "fnctCheckData"

        On Error Resume Next
        intTempReporterFilterSetting = Reporter.Filter
        Reporter.Filter = rfDisableAll
        blnReadOnlyCheckPerformed = false
        blnReadOnlyValue = false
        With obj_Page
            strActualValue = .WebEdit(strObjectOnPageName).GetROProperty("value")
            
            If Err.Number <> 0 Then
                Err.Clear
                strActualValue = .WebList(strObjectOnPageName).GetROProperty("selection")
            Else
                If str_State = "ReadOnly" Then
                    blnReadOnlyValue = .WebEdit(strObjectOnPageName).GetROProperty("disabled")
                    blnReadOnlyCheckPerformed = true
                End If
            End If
    
            If Err.Number <> 0 Then
                Err.Clear
                strActualValue = .WebRadioGroup(strObjectOnPageName).GetROProperty("value")
            Else
                If str_State = "ReadOnly" And Not(blnReadOnlyCheckPerformed) Then
                    blnReadOnlyValue = .WebList(strObjectOnPageName).GetROProperty("disabled")
                    blnReadOnlyCheckPerformed = true
                End If
            End If
                
            If Err.Number <> 0 Then
                Err.Clear
                strActualValue = .WebCheckBox(strObjectOnPageName).GetROProperty("checked")
                strActualValue = CBool(strActualValue)
            Else
                If str_State = "ReadOnly" And Not(blnReadOnlyCheckPerformed) Then
                    blnReadOnlyValue = .WebRadioGroup(strObjectOnPageName).GetROProperty("disabled")
                    blnReadOnlyCheckPerformed = true
                End If
            End If
    
            If Err.Number <> 0 Then
                Err.Clear
                strActualValue = .WebFile(strObjectOnPageName).GetROProperty("value")
            Else
                If str_State = "ReadOnly" And Not(blnReadOnlyCheckPerformed) Then
                    blnReadOnlyValue = .WebCheckBox(strObjectOnPageName).GetROProperty("disabled")
                    blnReadOnlyCheckPerformed = true
                End If
            End If
    
            If Err.Number = 0 Then
                If str_State = "ReadOnly" And Not(blnReadOnlyCheckPerformed) Then
                    blnReadOnlyValue = .WebFile(strObjectOnPageName).GetROProperty("disabled")
                    blnReadOnlyCheckPerformed = true
                End If
            Else
                Err.Raise 2, , "Could not find object <" & strObjectOnPageName & "> and retrieve any value"
            End If

        End With
        Reporter.Filter = intTempReporterFilterSetting

        onfnctCheckData.fire Me
        Err.Clear

    End Sub

End Class

Function fnctCheckValues(strActualValue, strExpectedValue)
  
	fnctCheckValues = false

    If IsNull(strExpectedValue) Then
		'@Replace nulls, as database will not store zero length string
		strExpectedValue = ""
	End If

	On Error Resume next
	If strActualValue = strExpectedValue Then
		Reporter.ReportEvent micDone, "<" & strActualValue & "> was equal to <" & strExpectedValue & ">", ""
        fnctCheckValues = true
	ElseIf Trim(strExpectedValue) = Trim(strActualValue) Then
		Reporter.ReportEvent micDone, "The expected value <" & strExpectedValue & "> was found to be within the actual value <" & strActualValue & ">", ""
        fnctCheckValues = true
	ElseIf LCase(strExpectedValue) = LCase(strActualValue) Then
		Reporter.ReportEvent micDone, "The actual value <" & strActualValue & "> was found to be the same as the case agnostic expected value <" & strExpectedValue & ">", ""
        fnctCheckValues = true
	ElseIf IsNumeric(Trim(strExpectedValue)) And IsNumeric(Trim(strActualValue)) Then
		If Round(strExpectedValue,5) = Round(strActualValue,5) Then
			Reporter.ReportEvent micDone, "The value matched the value expected, after rounding both to 5dp", "strActualValue <" & strActualValue & ">, strExpectedValue <" & strExpectedValue & "> " & ""
            fnctCheckValues = true
		End If
	ElseIf IsDate(Trim(strExpectedValue)) And IsDate(Trim(strActualValue)) Then
		If CDate(strExpectedValue) = CDate(strActualValue) Then
			Reporter.ReportEvent micDone, "The value matched the value expected, after converting both to dates", "strActualValue <" & strActualValue & ">, strExpectedValue <" & strExpectedValue & "> " & ""
            fnctCheckValues = true
		End If
    ElseIf VarType(strActualValue) = 11 Then
        If strActualValue Then
            If LCase(strExpectedValue) = "yes" Or strExpectedValue = 1 Or strExpectedValue = -1 Or strExpectedValue = "-1" Or strExpectedvalue = "true" Or strExpectedValue Then
                Reporter.ReportEvent micDone, "The value matched the value expected, after converting booleans to strings", "strActualValue <" & strActualValue & ">, strExpectedValue <" & strExpectedValue & "> " & ""
                fnctCheckValues = true
            End If
        Else
            If LCase(strExpectedValue) = "no" Or strExpectedValue = 0 Or strExpectedValue = "0" Then
                Reporter.ReportEvent micDone, "The value matched the value expected, after converting booleans to strings", "strActualValue <" & strActualValue & ">, strExpectedValue <" & strExpectedValue & "> " & ""
                fnctCheckValues = true
            End If
        End If
    ElseIf VarType(strExpectedValue) = 11 Then
        If strExpectedValue Then
            If Trim(LCase(strActualValue)) = "yes" Or strExpectedValue = 1 Or strExpectedValue = -1 Then
                Reporter.ReportEvent micDone, "The value matched the value expected, after converting booleans to strings", "strActualValue <" & strActualValue & ">, strExpectedValue <" & strExpectedValue & "> " & ""
                fnctCheckValues = true
            End If
        Else
            If Trim(LCase(strActualValue)) = "no" Or strExpectedValue = 0 Then
                Reporter.ReportEvent micDone, "The value matched the value expected, after converting booleans to strings", "strActualValue <" & strActualValue & ">, strExpectedValue <" & strExpectedValue & "> " & ""
                fnctCheckValues = true
            End If
        End If
	End If
End Function



'@   FRIENDS LIFE DORKING TEST AUTOMATION TEAM
'@   11/05/2011










'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'@   U S E R - D E F I N E D   C L A S S E S
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////










'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'@   C L A S S   I N I T I A L I Z A T O R S  (bridge between wscript and qtp namespace)
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////










'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'@   F U N C T I O N S
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


Function fnctBrowserExists(iTimeOut, objBrowser)
    Dim iSyncStart, intTempReporterFilterSetting
    fnctBrowserExists = false
    iSyncStart = Timer

    '@Switch off reporting, as this has a tendency to generate many useless warning events when reporting is on
    intTempReporterFilterSetting = Reporter.Filter
    Reporter.Filter = rfDisableAll

    Do 
         If (Timer - iSyncStart) > iTimeOut Then : Exit Function
    
    Loop While Not(objBrowser.Exist(0))

    fnctBrowserExists = true
    Reporter.Filter = intTempReporterFilterSetting
End Function

Function funcGetPropertyByTagName (objObject, strTag, intIndex, strProperty)
    Set oLocalObject = objObject.Object
    Set oTags        = oLocalObject.getElementsByTagName(strTag)
    funcGetPropertyByTagName = Eval ("oTags(intIndex)." & strProperty)
End Function ' funcGetPropertyByTagName


'*************************************************************************************************
'@Name          fnctRandomStr
'@Description   Returns a random string with the number of letters specified in chars
'*************************************************************************************************
Function fnctRandomStr(byVal intChars)
    Dim i
    For i = 1 to intChars step 1
        Randomize
        fnctRandomStr = fnctRandomStr & chr(Int((122 - 97 + 1) * Rnd + 97))
    Next
End Function

'*************************************************************************************************
'@Name        fnctAddLeadingZeroesToLength
'@Description 	Adds as many zeroes as are necessary to make up a given length, given 1 string input and a length.
'*************************************************************************************************
Function fnctAddLeadingZeroesToLength(byVal strInput, byVal intLength)
    Do while Len(strInput) < intLength
        strInput = "0" & strInput
    Loop
    fnctAddLeadingZeroesToLength = strInput
End Function

'*************************************************************************************************
'@Name          fnctGetNumbersFromString
'@Description   Gets all numeric values from a string
'************************************************************************************
Function fnctGetNumbersFromString(byVal strInputString)
    Dim intStringLength
    Dim ChCurrentChar
    Dim strFormattedString
    
    fnctGetNumbersFromString = ""
    '@ Check if string is null or empty
    If IsNull(strInputString) Or strInputString = "" Then
        Exit Function
    End If

    intStringLength = Len(strInputString)
    
    For i = 1 to intStringLength
        ChCurrentChar = Mid(strInputString, i, 1)
        If IsNumeric(ChCurrentChar) Then
            strFormattedString = strFormattedString & ChCurrentChar
        End If
    Next
    fnctGetNumbersFromString = strFormattedString
End Function

'*************************************************************************************************
'@Name          fnctReturnLeftOfStringUpToSubString
'@Description   Returns all of a string up to a contained substring
'************************************************************************************
Function fnctReturnLeftOfStringUpToSubString(byVal strString, byVal strText)
    Dim intStartPosition

    intStartPosition = InStr(strString, strText)
    fnctReturnLeftOfStringUpToSubString = Left(strString, intStartPosition - 1)
End Function




'***************************************************************************************************
'  C U S T O M  S Y N C	C U S T O M  S Y N C		C U S T O M  S Y N C		C U S T O M  S Y N C		C U S T O M  S Y N C		C U S T O M  S Y N C

' Example 1
'	hwnd = Browser("B").Object.HWND
'	Browser("B").Page("P").Link("L").Click
'	CustomSync hwnd, False, "", 30

' Example 2
'	hwnd = Browser("B").Object.HWND
'	Browser("B").Page("P").Link("L").Click
'	CustomSync hwnd, True, "", 30

' Example 3
'	hwnd = Browser("B").Object.HWND
'	Browser("B").Page("P").Link("L").Click
'	CustomSync hwnd, True, "html_element_id", 30

' Example 4
'	hwnd = Browser("B").Object.HWND
'	Browser("B").Page("P").WebButton("WB").Click
'	If NOT CustomSync (hwnd, True, "html_element_id", 30) Then : _
'	Reporter.ReportEvent micFail, "Submit Form", "The Form has not been submitted in 30 sec"

'***************************************************************************************************

Function CustomSync (hWnd, bPage, sElemId, iTimeOut)
Dim oShell, oBrowser, iSyncStart

CONST READYSTATE_BROWSER = 4
CONST READYSTATE_DOCUMENT = "complete"
CONST READYSTATE_ELEMENT = "complete"
CustomSync = FALSE
'asdfasdfsda
	' Get browser window by hWnd
	Set oShell = CreateObject ("Shell.Application")
	For each oBrowser in oShell.Windows
		' There could be items in the collection that do not have hwnd property
		On Error Resume Next
			winHwnd = oBrowser.hwnd
		On Error GoTo 0
		If winHwnd = hwnd Then : Exit For
	Next

	' Kick off Sync Timer
	iSyncStart = Timer

	' Sync browser window
	Do
		If (Timer - iSyncStart) > iTimeOut Then : Exit Function
	Loop While oBrowser.ReadyState <> READYSTATE_BROWSER

	' Sync browser window document
	If bPage Then
		Do
			If (Timer - iSyncStart) > iTimeOut Then : Exit Function
		Loop While oBrowser.Document.ReadyState <> READYSTATE_DOCUMENT
	End If

	If sElemId <> "" Then

		' Wait till element appears on the page
		Do
			If (Timer - iSyncStart) > iTimeOut Then : Exit Function
		Loop While oBrowser.Document.GetElementById (sElemId) Is Nothing

		' Sync element
		Do
			If (Timer - iSyncStart) > iTimeOut Then : Exit Function
		Loop While oBrowser.Document.GetElementById (sElemId).ReadyState <> READYSTATE_ELEMENT
	End If

CustomSync = TRUE
End Function


'***************************************************************************************************
'	C U S T O M  S Y N C  TO  O B J E C T  S T Y L E		C U S T O M  S Y N C  TO  O B J E C T  S T Y L E		C U S T O M  S Y N C  TO  O B J E C T  S T Y L E

' Example 1
'	hwnd = Browser("B").Object.HWND
'	Browser("B").Page("P").Link("L").Click
'	CustomSyncToObjectStyle hwnd, "html_element_id", "display", "none", 30

' Example 2
'	hwnd = Browser("B").Object.HWND
'	Browser("B").Page("P").WebElement("WEL").Click
'	If NOT CustomSyncToObjectStyle (hwnd, "html_element_id", "display", "none", 30) Then : _
'	Reporter.ReportEvent micFail, "Show list", "The List has not been displayed in 30 sec"

'***************************************************************************************************

Function funcSyncToObjectAttribute (hWnd, sElemId, sAttribute, sValue, iTimeOut)
    Dim oShell, oBrowser, iSyncStart, ret

    funcSyncToObjectAttribute = FALSE

    ' Get browser window by hWnd
    Set oShell = CreateObject ("Shell.Application")
    For each oBrowser in oShell.Windows
        ' There could be items in the collection that do not have hwnd property
        On Error Resume Next
                        winHwnd = oBrowser.hwnd
        On Error GoTo 0
        If winHwnd = hwnd Then : Exit For
    Next

    ' Kick off Sync Timer
    iSyncStart = Timer
    
    ' Wait till element appears on the page
    Do
       If (Timer - iSyncStart) > iTimeOut Then : Exit Function
    Loop While oBrowser.Document.GetElementById (sElemId) Is Nothing

    ' Sync element style property value
    Do
        If (Timer - iSyncStart) > iTimeOut Then : Exit Function
        On error resume next
                        ret = Eval ("oBrowser.Document.getElementById (sElemId).GetAttribute (sAttribute)")
        On error goto 0
    Loop While ret  <> sValue

    funcSyncToObjectAttribute = TRUE
End Function 'funcSyncToObjectAttribute



Function funcSyncToObjectAttributeByTagName (hWnd, sElemId, sAttribute, sValue, iTimeOut)
    Dim oShell, oBrowser, iSyncStart, ret

    ret = "not found"

    ' Get browser window by hWnd
    Set oShell = CreateObject ("Shell.Application")
    For each oBrowser in oShell.Windows
        ' There could be items in the collection that do not have hwnd property
        On Error Resume Next
                        winHwnd = oBrowser.hwnd
        On Error GoTo 0
        If winHwnd = hwnd Then : Exit For
    Next

    ' Kick off Sync Timer
    iSyncStart = Timer
    
    ' Wait till element appears on the page
    Do
        on error resume next
        ret = typename (oBrowser.Document.GetElementsBytagname (sElemId).item(0))
        on error goto 0
    Loop Until ret = "HTMLHeaderElement" or (Timer - iSyncStart) > iTimeOut

    ' Sync element style property value
    Do
        On error resume next
            ret = Eval ("oBrowser.Document.getElementsByTagName (sElemId).item(0).GetAttribute (sAttribute)")
        On error goto 0
    Loop Until inStr (ret, sValue) or (Timer - iSyncStart) > iTimeOut

    funcSyncToObjectAttributeByTagName = ret
End Function ' funcSyncToObjectAttributeByTagName


Function funcSyncToObjectAttributeByTagName (hWnd, sElemId, sAttribute, sValue, iTimeOut)
    Dim oShell, oBrowser, iSyncStart, ret

    ret = false

    ' Get browser window by hWnd
    Set oShell = CreateObject ("Shell.Application")
    For each oBrowser in oShell.Windows
        ' There could be items in the collection that do not have hwnd property
        On Error Resume Next
                        winHwnd = oBrowser.hwnd
                        If err.number <> 0 Then : print err.description
        On Error GoTo 0
        If winHwnd = hwnd Then : Exit For
    Next

    ' Kick off Sync Timer
    iSyncStart = Timer

	' Sync browser window
	Do
		If (Timer - iSyncStart) > iTimeOut Then : Exit Function
	Loop While oBrowser.ReadyState <> 4

	' Sync browser window document
	If bPage Then
		Do
			If (Timer - iSyncStart) > iTimeOut Then : Exit Function
		Loop While oBrowser.Document.ReadyState <> "complete"
	End If

    ' Sync element style property value
    Do
        On error resume next
            intColLen = oBrowser.Document.getElementsByTagName (sElemId).length
            for idxElem = 0 to intColLen -1
                do : loop until (Timer - iSyncStart) > iTimeOut or typename (oBrowser.Document.getElementsByTagName (sElemId).item (idxElem)) <> "Nothing"
                ret = eval (trim (oBrowser.Document.getElementsByTagName (sElemId).item (idxElem).getAttribute (sAttribute)) = sValue)
                If ret Then : exit for
            next
        On error goto 0
    Loop Until ret or (Timer - iSyncStart) > iTimeOut

    if ret then : funcSyncToObjectAttributeByTagName = sValue
End Function 'funcSyncToObjectAttributeByTagName

Function funcSyncToObjectPartAttributeByTagName (hWnd, sElemId, sAttribute, sValue, iTimeOut)
    Dim oShell, oBrowser, iSyncStart, ret

    ret = false
    funcSyncToObjectPartAttributeByTagName = false
    ' Kick off Sync Timer
    iSyncStart = Timer

    '@Keep trying to loop through until a value is returned
    Do
        ' Get browser window by hWnd
        Set oShell = CreateObject ("Shell.Application")
        For each oBrowser in oShell.Windows
            ' There could be items in the collection that do not have hwnd property
            On Error Resume Next
                            winHwnd = oBrowser.hwnd
                            If err.number <> 0 Then : print err.description
            On Error GoTo 0
            If winHwnd = hwnd Then : Exit For
        Next

    	'@Wait for the browser to appear ready
    	Do
    		If (Timer - iSyncStart) > iTimeOut Then : Exit Function
    	Loop While oBrowser.ReadyState <> 4

        '@Wait for a document to appear ready
        Do
			If (Timer - iSyncStart) > iTimeOut Then : Exit Function
		Loop While oBrowser.Document.ReadyState <> "complete"
    
        ' Sync element style property value
        On error resume next
            intColLen = oBrowser.Document.getElementsByTagName (sElemId).length
            for idxElem = 0 to intColLen -1
                ret = Instr( 1, oBrowser.Document.getElementsByTagName (sElemId).item (idxElem).getAttribute (sAttribute), sValue)
                If ret <> 0 Then : exit for
            next
        On error goto 0

        Set oShell = nothing
        Set oBrowser = nothing

    Loop Until ret <> 0 or (Timer - iSyncStart) > iTimeOut

    if ret <> 0 then : funcSyncToObjectPartAttributeByTagName = sValue
End Function 'funcSyncToObjectPartAttributeByTagName


Function funcSyncToPageHeaderAttribute (hWnd, sElemId, sAttribute, sValue, iTimeOut)
    Dim oShell, oBrowser, iSyncStart, ret

    ret = "not found"

    ' Get browser window by hWnd
    Set oShell = CreateObject ("Shell.Application")
    For each oBrowser in oShell.Windows
        ' There could be items in the collection that do not have hwnd property
        On Error Resume Next
                        winHwnd = oBrowser.hwnd
        On Error GoTo 0
        If winHwnd = hwnd Then : Exit For
    Next

    ' Kick off Sync Timer
    iSyncStart = Timer
    
    ' Wait till element appears on the page
    Do
        on error resume next
        ret = typename (oBrowser.Document.GetElementsBytagname (sElemId).item(0))
        on error goto 0
    Loop Until ret = "HTMLHeaderElement" or (Timer - iSyncStart) > iTimeOut

    ' Sync element style property value
    Do
        On error resume next
            ret = Eval ("oBrowser.Document.getElementsByTagName (sElemId).item(0).GetAttribute (sAttribute)")
        On error goto 0
    Loop Until inStr (ret, sValue) or (Timer - iSyncStart) > iTimeOut

    funcSyncToPageHeaderAttribute = ret
End Function ' funcSyncToPageHeaderAttribute

Function funcSyncIEHandler(iTimeOut, objBrowser)
    Dim iSyncStart
    funcSyncIEHandler = false
    iSyncStart = Timer

    Do 
         If (Timer - iSyncStart) > iTimeOut Then : Exit Function
            On error Resume Next
            strTempHWND = objBrowser.object.hwnd
            If Err.Number = 0 Then
                funcSyncIEHandler = strTempHWND
                Exit Do
            End If
    
    Loop While Not(funcSyncIEHandler)

End Function

'   FRIENDS LIFE DORKING TEST AUTOMATION TEAM
'   11/05/2011









'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'@   C L A S S E S
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'@   R e c o r d s e t

Class clsRecordSet

    '@ declare variables that will be accessed from the global script
    Public object, onConnect, onOpen, onUpdate, onAddRecord, onExport, onAddMissingHeaders
    
    '@ declare variables that will only be used within the class
    Private strCurrentMethod_, arrHeaders_, strConnect_, strOpen_, strUpdate_, strAddRecord_, strExport_, strAddMissingHeaders_
    
    '@ this function runs each time object is initialized from this class
    Private Sub Class_Initialize ()
        set object =  createObject ("ADODB.RecordSet") '@ create an object that will be an interface to Recordset class of ADODB library (COM)
        '@ create event objects (each event object is a collection of function refferences with an ability to be executed (fired))
        set onConnect = New clsHandlers
        set onOpen = New clsHandlers
        set onUpdate = New clsHandlers
        set onAddRecord = New clsHandlers
        set onExport = New clsHandlers
        set onAddMissingHeaders = New clsHandlers
    End Sub
    
    '@ this function runs each time the object  is released
    Private Sub Class_Terminate ()
        on error resume next '@ switch off the WSH error handling (in case when the object is released before the connection to DB is made)
        object.close         '@ close a recordset
    End Sub
    
    '@ a class property, that collects most important object properties and class variables into one object - for logging
    Public Property Get properties ()
        set dictProperties = createObject ("scripting.dictionary")
        dictProperties.add "method", strCurrentMethod_
        dictProperties.add "connect", strConnect_
        dictProperties.add "open", strOpen_
        dictProperties.add "update", strUpdate_
        dictProperties.add "addRecord", strAddRecord_
        dictProperties.add "export", strExport_
        dictProperties.add "addMissingHeaders", strAddMissingHeaders_
        set properties = dictProperties
    End Property
    
    '@ set recordset connection to DB
    Public Sub connect (strDataBaseFullPath)
        strCurrentMethod_ = "connect" '@ note the method that will be executed - for logging
        strConnect_ = strDataBaseFullPath
        on error resume next '@ switch off WSH error handling (custom error handling will be used)
  	object.ActiveConnection = "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = "& strDataBaseFullPath '@ set DB connection string, that automatically creates the connection object
		onConnect.fire Me '@ use fire method of the Event object; all functions (handlers) that are referenced by this object will be executed (Me object represents a current class)
    End sub

    '@ open a recordset by executing an sql query
    Public Sub open (strSql)
        strCurrentMethod_ = "open" '@ note the method that will be executed - for logging
        strOpen_ = strSql
        on error resume next '@ switch off WSH error handling (custom error handling will be used)
            '@ set recordset parameters before openning it (these parameters are properties of ADODB Recordset class and can be found in MSDN for reference)
            object.cursorType = 3 '@ a static cursor allowing forward and backward scrolling of a fixed, unchangeable set of records
            object.lockType = 3 '@ multiple users can modify the data which is not locked until Update method is called
            object.open strSql '@ return (open) a recordset by executing SQL
            onOpen.fire Me '@ use fire method of the Event object; all functions (handlers) that are referenced by this object will be executed (Me object represents a current class)
    End Sub
    
    '@ update a recordset field with new value
    Public Sub update (strField, strNewValue)
        strCurrentMethod_ = "update"
        strUpdate_ = strField & vbNewLine & strNewValue
        on error resume next
            object ( strField ) = strNewValue '@ assign new value to a field
            object.update '@ update a recordset (save)
            onUpdate.fire Me
    End Sub
    
    '@ add a new record of values to recordset
    Public Sub addRecord (arrHeaders, arrValues)
        strCurrentMethod_ = "addRecord"
        on error resume next
            strAddRecord_ = join (arrValues, vbNewLine)
            object.AddNew arrHeaders, arrValues '@ add a new record of values to a pre-set array of recordset fields
            onAddRecord.fire Me
    End Sub

    public sub export (intIterations, strFilePath)
        strCurrentMethod_ = "export"
        strExport_ = intIterations & vbNewLine & strFilePath
        set myFile = createObject ("scripting.FileSystemObject").createTextFile (strFilePath)
        myFile.WriteLine join (listHeaders, vbTab)
        a = intIterations \ object.recordCount
        b = intIterations mod object.recordCount
        for i = 1 to a
            myFile.Write object.getString (2,, vbTab,, "")
            object.moveFirst
        next
        if b <> 0 then : myFile.WriteLine object.getString (2, b, vbTab,, "")
        myFile.Close
        onExport.fire me
    end sub

    public function listHeaders ()
        ReDim arrHeaders (-1)
        for each field in object.fields
            Redim Preserve arrHeaders (UBound (arrHeaders) +1)
            arrHeaders (UBound(arrHeaders)) = field.name
        next
        listHeaders = arrHeaders
    end function

    public function listValues ()
        ReDim arrValues (-1)
        for each fld in object.Fields
            ReDim preserve arrValues (UBound (arrValues) +1)
            arrValues (UBound (arrValues)) = fld.value
        next
    listValues = arrValues
    end function

    '@ set up a disconnected recordset with array of fields and pre-set (constant) field parameters
    public Sub addMissingHeaders (arrHeaders)
        strCurrentMethod_ = "addMissingHeaders"
        strAddMissingHeaders_ = arrHeaders '@ note the array of recordset headers as an internal variable
        const FIELD_TYPE = 200 '@ numeric value for a recordset field (string type)
        const MAX_CHAR = 1024 '@ recordset field size
        on error resume next
            '@ add fields to recordset with pre-set parameters
            For Each strHeader in arrHeaders
                object.fields.append strHeader, FIELD_TYPE, MAX_CHAR
            Next
            onAddMissingHeaders.fire Me
    end Sub

    public sub addMissingTable ()
        strTheThing = join (arrFields, " TEXT, ")
        strTR = "CREATE TABLE "& tblName &" ("& strTheThing &" TEXT);"
        objConnection.Execute strTR
        objConnection.Close   
        set objConnection = nothing
        print err.description
        if err.number = 0 then : createTable = true
    end sub

    public function tableExists (strTableName)
        set rsTblSchema = object.activeConnection.OpenSchema (20)
        while not rsTblSchema.EOF and not rsTblSchema.BOF
            if tblName = rsTblSchema.fields("TABLE_NAME").value then : exit function
            rsTblSchema.MoveNext
        wend
        rsTblSchema.Close
        set rsTblSchema = nothing
    end function

End Class


'@   Da t a B a s e   C a t a l o g


class clsDataBaseCatalog

    '@ declare variables that will be accessed from the global script
    Public object, onCreate, onConnect, onTableExists, onAddTable
    
    '@ declare variables that will only be used within the class
    Private strCurrentMethod_, strCreate_, strConnect_, strTableExists_, strAddTable_, adoConnection
    
    '@ this function runs each time object is initialized from this class
    Private Sub Class_Initialize ()
        set object = CreateObject ("ADOX.Catalog") '@ create an object that will be an interface to Catalog class of ADOX library (COM)
        '@ create event objects (each event object is a collection of function refferences with an ability to be executed (fired))
        set onCreate = new clsHandlers
        set onConnect = new clsHandlers
        set onTableExists = new clsHandlers
        set onAddTable = new clsHandlers
    End Sub
    
    '@ this function runs each time the object is released (script stops running or object set to nothing)
    Private Sub Class_Terminate ()
        on error resume next '@ switch off the WSH error handling (in case when the object is released before the connection to DB is made)
        set object = nothing         '@ close a recordset
    End Sub
    
    '@ a class property, that collects most important object properties and class variables into one object - for logging
    Public Property Get properties ()
        set dictProperties = createObject ("scripting.dictionary")
        dictProperties.add "method", strCurrentMethod_
        dictProperties.add "create", strCreate_
        dictProperties.add "connect", strConnect_
        dictProperties.add "tableExists", strTableExists_
        dictProperties.add "addTable", strAddTable_
        set properties = dictProperties
    End Property
    
    Public sub create (strDataBassFullPath)
        strCurrentMethod_ = "create"
        strCreate_ = strDataBassFullPath
        on error resume next
            object.create "Provider = Microsoft.Jet.OLEDB.4.0; Data Source = "& strDataBassFullPath
            object.activeConnection.close
            onCreate.fire me
    End sub
    
    public sub connect (strDataBasePath)
        strCurrentMethod_ = "connect"
        strConnect_ = strDataBasePath
        on error resume next
            set adoConnection = createObject ("ADODB.connection")
            adoConnection.Open "Provider= Microsoft.Jet.OLEDB.4.0; Data Source= "& strDataBasePath 
            set object.activeConnection = adoConnection
            onConnect.fire me
    end sub
    
    public function tableExists (strTableName)
        strCurrentMethod_ = "tableExists"
        strtableExists_ = strTableName
        tableExists = false
        on error resume next
            for each tbl in object.tables
                if tbl.name = strTableName then : tableExists = true
            next
            onTableExists.fire me
    end function
    
    public sub addTable (strTableName, arrColumnNames)
        strCurrentMethod_ = "addTable"
        strAddTable_ = strTableName
        on error resume next
            set tbl = createObject ("ADOX.table")
            tbl.Name = strTableName
            for each strNm in arrColumnNames
                tbl.columns.append strNm, 200
            next
            object.tables.append tbl
            onAddTable.fire me
    end sub

end class










'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'@   C L A S S   I N I T I A L I Z A T O R S  (bridge between wscript and qtp namespace)
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////



Function funcCreateRecordSet ()
    set funcCreateRecordSet = new clsRecordSet
    funcCreateRecordSet.onConnect.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    funcCreateRecordSet.onOpen.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    funcCreateRecordSet.onUpdate.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    funcCreateRecordSet.onAddrecord.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    funcCreateRecordSet.onExport.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    funcCreateRecordSet.onAddMissingHeaders.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
End Function


function funcDataBaseCatalog ()
	set funcDataBaseCatalog = new clsDataBaseCatalog
    funcDataBaseCatalog.onCreate.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    funcDataBaseCatalog.onConnect.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    funcDataBaseCatalog.onTableExists.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    funcDataBaseCatalog.onAddTable.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
end function






'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'@   F U N C T I O N S
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'@   R e c o r d s e t   c o m p a r i s s o n

'@ Compare 2 recordsets by key field
Function funcCompareRecordsetsByField (rs1, rs2, sKeyField)
Dim sKeyFieldValue, sSearchCriteria, Field, bFound, sFieldName, sReportHeader, sReportBody, sReportTail

Dim dKeyFieldValues : Set dKeyFieldValues = CreateObject ("Scripting.Dictionary")
Dim dFieldNames : Set dFieldNames = CreateObject ("Scripting.Dictionary")

Dim ROW_DELIMITER : ROW_DELIMITER = vbNewLine
Dim ITEM_DELIMITER : ITEM_DELIMITER= "|"
Dim REPORT_DELIMITER : REPORT_DELIMITER = vbNewLine &"	***"& vbNewLine

    If Not rs1.BOF Then : rs1.MoveFirst
    If Not rs2.BOF Then : rs2.MoveFirst

    ' Get a dictionary of all possible fields in both recordsets and map them
    For each Field in rs1.Fields
        dFieldNames.Add Field.Name, 1
    Next
    For each Field in rs2.Fields
        If dFieldNames.Exists (Field.Name) Then
            dFieldNames.Item (Field.Name) = Field.Name
        Else
            dFieldNames.Add Field.Name, 2
        End If
    Next

    ' Report fields missing in one of the RS
    For each sFieldName in dFieldNames.Keys
        Select Case dFieldNames.Item (sFieldName) 
            Case "1"
                            sReportHeader = sReportHeader &"Field '"& sFieldName &"' not found in RS2" & ROW_DELIMITER
                            Reporter.ReportEvent micFail, "Recordset mismatch", sReportHeader
            Case "2"
                            sReportHeader = sReportHeader &"Field '"& sFieldName &"' not foud in RS1" & ROW_DELIMITER
                            Reporter.ReportEvent micFail, "Recordset mismatch", sReportHeader
            Case Else
                            'sReportHeader = sReportHeader &"Field '"& sFieldName &"' in RS1 matched '"& dFieldNames.Item (sFieldName) &"' in RS2"& vbNewLine
        End Select
    Next

    ' Get a dictionary of all possible records in both recordsets and map them
    While NOT rs1.EOF
                    sKeyFieldValue = rs1 (sKeyField).Value
                    If NOT dKeyFieldValues.Exists (sKeyFieldValue) Then : dKeyFieldValues.Add sKeyFieldValue, "a"
                    rs1.MoveNext
    Wend
    While NOT rs2.EOF
        sKeyFieldValue = rs2 (sKeyField).Value
        If dKeyFieldValues.Exists (sKeyFieldValue) Then 
            dKeyFieldValues.Item (sKeyFieldValue) = sKeyFieldValue
        Else
            dKeyFieldValues.Add sKeyFieldValue, "b"
        End If
        rs2.MoveNext
    Wend

    ' Report fields missing in one of the RS
    For each sKeyFieldValue in dKeyFieldValues.Keys
        Select Case dKeyFieldValues.Item (sKeyFieldValue)
            Case "a"
                sReportBody = sReportBody &"Record '"& sKeyFieldValue &"' not foud in RS2" & ROW_DELIMITER
                Reporter.ReportEvent micFail, "Recordset mismatch", sReportHeader
            Case "b"
                sReportBody = sReportBody &"Record '"& sKeyFieldValue &"' not foud in RS1" & ROW_DELIMITER
                Reporter.ReportEvent micFail, "Recordset mismatch", sReportHeader
            Case Else
                'sReportBody = sReportBody &"Record '"& sKeyFieldValue &"' in RS1 matched '"& dKeyFieldValues.Item (sKeyFieldValue) &"' in RS2"& vbNewLine
        End Select
    Next

    ' Report field mistmatches
    For each sKeyFieldValue in dKeyFieldValues
        If dKeyFieldValues.Item (sKeyFieldValue) <> "a" AND dKeyFieldValues.Item (sKeyFieldValue) <> "b" Then
                        
            rs1.Filter = "["& sKeyField &"]='"& sKeyFieldValue &"'"
            rs2.Filter = "["& sKeyField &"]='"& sKeyFieldValue &"'"

            For each sFieldName in dFieldNames.Keys
                If dFieldNames.Item (sFieldName) <> "1" AND dFieldNames.Item (sFieldName) <> "2" Then
                    If NOT rs1 (sFieldName).Value = rs2 (dFieldNames.Item (sFieldName)).Value Then
                        sReportTail = sReportTail & sKeyFieldValue & ITEM_DELIMITER & sFieldName & ITEM_DELIMITER & rs1 (sFieldName).Value & ROW_DELIMITER
                        sReportTail = sReportTail & sKeyFieldValue & ITEM_DELIMITER & dFieldNames.Item (sFieldName) & ITEM_DELIMITER & rs2 (dFieldNames.Item (sFieldName)).Value & ROW_DELIMITER & ROW_DELIMITER
                    End If
                End If
            Next
        End If
    Next

    If sReportHeader <> "" Then : sReportHeader = Left (sReportHeader, Len (sReportHeader) - Len (ROW_DELIMITER))
    If sReportBody <> "" Then : sReportBody = Left (sReportBody, Len (sReportBody) - Len (ROW_DELIMITER))
    If sReportTail <> "" Then : sReportTail = Left (sReportTail, Len (sReportTail) - Len (ROW_DELIMITER & ROW_DELIMITER))
                                
'funcCompareRecordsetsByField = sReportHeader & REPORT_DELIMITER & sReportBody & REPORT_DELIMITER & sReportTail

set dictResults = createObject ("scripting.dictionary")
dictResults.Add "FIELDS", sReportHeader
dictResults.Add "RECORDS", sReportBody
dictResults.Add "VALUES", sReportTail

set funcCompareRecordsetsByField = dictResults
End Function

function exportCsvtoDataBaseCreateTable (strCsvFileName, strCsvLocation, strTableName, strDataBasePath)

    set specFile = createObject ("scripting.fileSystemObject").createTextFile (strCsvLocation +"\Schema.ini")
    specFile.WriteLine "["+ strCsvFileName  +"]"
    specfile.WriteLine "ColNameHeader=True"
    specfile.Writeline "CharacterSet=ANSI"
    specFile.WriteLine "Format=TabDelimited"
    specFile.Close

    set cnn = CreateObject("ADODB.Connection")  
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="+ strDataBasePath
    sqlString = "SELECT * INTO "+ strTableName +" FROM [Text;HDR=YES;FMT=Delimited;DATABASE="+ strCsvLocation +"].["+ strCsvFileName +"]"  

    cnn.Execute sqlString  
    cnn.Close

end function

function exportCsvtoDataBaseExistingTable (strCsvFileName, strCsvLocation, strTableName, strDataBasePath)

    set specFile = createObject ("scripting.fileSystemObject").createTextFile (strCsvLocation +"\Schema.ini")
    specFile.WriteLine "["+ strCsvFileName  +"]"
    specfile.WriteLine "ColNameHeader=True"
    specfile.Writeline "CharacterSet=ANSI"
    specFile.WriteLine "Format=TabDelimited"
    specFile.Close

    set cnn = CreateObject("ADODB.Connection")  
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="+ strDataBasePath
    sqlString = "INSERT INTO "+ strTableName +" SELECT * FROM [Text;HDR=YES;FMT=Delimited;DATABASE="+ strCsvLocation +"].["+ strCsvFileName +"]"

    cnn.Execute sqlString  
    cnn.Close

end function

Sub subCreateLoggingTable (strDataBasePath)
    set cnn = CreateObject("ADODB.Connection")  
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="+ strDataBasePath
    sqlString = "CREATE TABLE " & _
        "tblFailure (pkeyFailureID AUTOINCREMENT UNIQUE PRIMARY KEY," & _
        "strFailureType VARCHAR(255)," & _
        "strMessage TEXT," & _
        "strIteration VARCHAR(255)," & _
        "strActionName VARCHAR(255)," & _
        "strTestName TEXT," & _
        "tstmp DATETIME," & _
        "strResultPath TEXT," & _
        "intDefectNumber VARCHAR(255)," & _
        "strRequirementReference VARCHAR(255))"
    cnn.Execute sqlString  
    cnn.Close
End Sub



'11:55 26/04/2011


'@                                                       C L A S S   D E F I N I T I O N

Function fnctInitializeClsHandlers
   Set fnctInitializeClsHandlers = New clsHandlers
End Function

'@   H a n d l e r   C l a s s
 
Class clsHandlers



    Private arrHandlers_ ()


    '@   d e c l a r e   h a n d l e r   s t o r a g e   a r r a y
    Private Sub Class_Initialize ()

        ReDim arrHandlers_ (-1) '@ dynamic array with no elements in it

    End Sub


    '@   a d d   h a n d l e r s   t o   s t o r a g e   a r r a y
    Public Sub addHandler ( arrFunctionNames )

        for each strFunctionName in arrFunctionNames
            ReDim Preserve arrHandlers_ ( UBound ( arrHandlers_ ) +1) '@ add an empty element to local storage array
            Set arrHandlers_ ( UBound ( arrHandlers_ ) ) = GetRef ( strFunctionName ) '@ set a new element to reference the current function in the list
        next

    End Sub


    '@   e x e c u t e   h a n d l e r s   f r o m   s t o  r a g e
    Public Sub fire ( args )

        for each refHandler in arrHandlers_
            refHandler args
        next

    End Sub



End Class


sub printMe (objCaller)
   ' print err.description
end sub


'@                                                       E R R O R   H A N D L E R S


'@   E x i t   t e s t   i f   e r r o r   e x i s t s
Sub exitTestOnError ( objCaller )
    if err.number <> 0 then : exitTest("onErrorExitTest")
End Sub


'@   E x i t   i t e r a t i o n   i f   e r r o r   e x i s t s
Sub exitIterationOnError ( objCaller )
    if err.number <> 0 then : exitTestIteration("onErrorExitIteration")
End Sub

Sub exitIterationAndCloseBrowserOnError ( objCaller )
    If err.number <> 0 then 
        On Error Resume Next
        Browser("index:=0").Close
        Err.Clear
        On Error Goto 0
        ExitTestIteration("exitIterationAndCloseBrowserOnError")
    End If
End Sub

'@   R e p o r t   e r r o r   i f   f o u n d
Sub ReportOnError ( objCaller )
    if err.number = 0 then : exit sub
    strMsg = objCaller.properties () ("method") & vbNewLine & err.description 
    reporter.ReportEvent micFail, TypeName ( objCaller ), strMsg
End Sub

'@   P r i n t   e r r o r   i f   f o u n d
sub printError (objCaller)
    if err.number = 0 then : exit sub
    print "error "+ cStr (err.number) +": "+ typeName (objCaller) +"."+ objCaller.properties () ("method") +" - "+ err.description
end sub

Sub logCallerPropertiesToDatabaseOnError ( objCaller )
    const EVNT_TYPE = 2
    if err.number = 0 then : exit sub
    set dictProperties = objCaller.properties
    strMsg = TypeName ( objCaller ) & vbNewLine
    for each strProperty in dictProperties.keys
        strMsg = strMsg & vbNewLine & strProperty &": "& dictProperties.Item ( strProperty )
    next
    subReportError "logCallerPropertiesToDatabaseOnError", strMsg
End Sub

Sub fnctReportDataEntrySuccess ( objCaller )
    If err.number = 0 then
        strMsg = "Set value into fields for object name <" & objCaller.strObjectOnPageName & ">, value <" & objCaller.strValue & ">"
        Reporter.ReportEvent micPass, strMsg, ""
    End If
End Sub

Sub logCallerPropertiesToQTPReportOnError ( objCaller )
    const EVNT_TYPE = 2
    if err.number = 0 then : exit sub
    set dictProperties = objCaller.properties
    strMsg = TypeName ( objCaller ) & vbNewLine
    for each strProperty in dictProperties.keys
        strMsg = strMsg & vbNewLine & strProperty &": "& dictProperties.Item ( strProperty )
    next
    Reporter.ReportEvent micFail, strMsg, ""
End Sub



'@                                                       E V E N T   H A N D L E R S


'@   R a i s e   e r r o r   i f   r e c o r d s e t  i s   e m p t y
Sub checkIfAnyRecordsReturned ( objCaller )

    if objCaller.object.recordCount = 0 then : on error resume next : err.raise 1, , "no records found"

End sub



'@                                                       U T I L I T Y   H A N D L E R S


'@   L o g   t e s t   d a t a
sub logDT ()

    'environment ("ResultDir") &"\"& environment ("ActionName") & ".txt"

end sub 


'@ update status field of the scheduled recordset with current date and time
Sub logToDB ( objCaller )
    if typename ( objCaller ) = "clsConnectedRecordSet" Then : rsSchedule.update "STATUS", now : exit sub
    if typename ( objCaller ) = "clsQuickTestApplication" Then : rsSchedule.update "STATUS", qtApp.object.test.lastRunResults.status
End Sub


'@ log methods to txt
Sub logMethodToTxt (objCaller)

  const READ = 1, WRITE = 2, APPEND = 8
	
	set fso = createobject ("scripting.filesystemobject")
	set file = fso.OpenTextFile (fso.buildPath (environment.Value ("TestDir"), createObject ("WScript.Network").ComputerName &".html"), APPEND, true)

	if err.number <> 0 then : file.writeline " <table><tr class='err'>"& err.description  &"</tr></table>"
	
	file.writeline "<table>"
	file.writeline "<tr>"
	file.write "<td class='date'>" : file.write now : file.write "</td>"
	file.write "<td class='class'>" : file.write typeName (objCaller) : file.write "</td>"
	file.write "<td class='method'>" : file.write objCaller.properties () ("method") : file.write "</td>"
	file.write "<td class='content'>" : file.write replace (objCaller.properties () (objCaller.properties () ("method")), vbNewLine, "<br />")  : file.write "</td>"
	file.writeline "</tr>"
	file.writeline "</table>"
	file.writeline "<link rel='stylesheet' type='text/css' href='mystyle.css' />"
	
End Sub


Sub subReportError(strFailureType, strMessage)

    strOutputDBPath = createobject ("scripting.filesystemobject").GetParentFolderName (environment.Value ("TestDir"))
    strOutputDBPath = createobject ("scripting.filesystemobject").BuildPath (strOutputDBPath, "logging.mdb")

    set cnn = CreateObject("ADODB.Connection")  
    cnn.Open "Provider=Microsoft.ACE.OLEDB.12.0; Data Source="+ strOutputDBPath
    sqlString = "INSERT INTO tblFailure (strFailureType, " & _
        "strMessage, " & _
        "strIteration, " & _
        "strActionName, " & _
        "strTestName, " & _
        "tstmp, " & _
        "strResultPath) " & _
        "VALUES ('" & strFailureType & "','" & _
        strMessage & "','" & _
        Environment.Value("TestIteration") & "','" & _
        Environment.Value("ActionName") & "','" & _
        Environment.Value("TestName") & "','" & _
        Now & "','" & _
        Environment.Value("ResultDir") & "');"

    cnn.Execute sqlString
    cnn.Close
End Sub



'@   FRIENDS LIFE DORKING TEST AUTOMATION TEAM
'@   11/05/2011










'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'@   C L A S S E S
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////


'@   Q t p   C l a s s   c l s Q u i c k T e s t A p p l i c a t i o n

Class clsQuickTestApplication

    '@ declare variables that will be accessed from the global script
    Public object, onAddCodeToAction, onOpen, onConnect

    '@ declare variables that will only be used within the class
    Private strCurrentMethod_

    '@ this function runs each time object is initialized from this class
    Private Sub Class_Initialize ()
        '@ create event objects (each event object is a collection of function refferences with an ability to be executed (fired))
        set onAddCodeToAction = new clsHandlers
        set onOpen = new clsHandlers
        set onConnect = new clsHandlers
    End Sub

    '@ a class property, that collects most important object properties and class variables into one object - for logging
    Public Property Get properties ()
        set dictProperties = createObject ("scripting.dictionary")
        dictProperties.add "method", strCurrentMethod_
        set properties = dictProperties
    End Property

    '@ add any string to an action by action index
    Public Sub addCodeToAction ( intActionIndex, strCode )
        strCurrentMethod_ = "addCodeToAction"                   '@ note the method that will be executed - for logging
        on error resume next                                    '@ switch off WSH error handling (custom error handling will be used)
            object.Test.Actions( intActionIndex ).SetScript strCode '@ set code to action
            onAddCodeToAction.fire Me                               '@ use fire method of the Event object; all functions (handlers) that are referenced by this object will be executed (Me object represents a current class)
    End Sub

    '@ add an array of lookup folders to QTP options
    Public Sub setFolders ( arrFolderPaths )
        on error resume next     '@ switch off WSH error handling (custom error handling will be used)
            object.folders.removeAll '@ clear pre-set folder list
            '@ add folder paths to the list one by one
            for each strPath in arrFolderPaths
                object.folders.add( strPath )
            next
    End Sub

    '@ launch QTP
    Public Sub open ()
        strCurrentMethod_ = "open"                          '@ note the method that will be executed - for logging
        on error resume next                                '@ switch off WSH error handling (custom error handling will be used)
            set object = createObject ("QuickTest.Application") '@ create an object that will be an interface to Application class of QuickTest library (COM)
            object.Launch                                       '@ launch application
            object.Visible = TRUE                               '@ make application UI visible
            onOpen.fire Me                                      '@ use fire method of the Event object; all functions (handlers) that are referenced by this object will be executed (Me object represents a current class)
    End Sub

    '@ force close QTP
    Public Sub close ()
        on error resume next '@ switch off WSH error handling (custom error handling will be used)
            object.quit          '@ quit application
            wscript.sleep 2000   '@ allow 2 seconds before checking if application is closed
            '@ kill all qtp related processes if they are still in the process list
            for each objProcess in getObject ("winmgmts:").InstancesOf ("Win32_process")
                if objProcess.name = "QTPro.exe" or objProcess.name = "QTReport.exe" then : objProcess.terminate
            next
    End Sub

    '@ combine data received from environment xml with QTP statements to get an Environment variable setup code
    Public Function generateEnvironmentConfigurationCode ( objEnvironmentSettingsNode )
        generateEnvironmentConfigurationCode = ""      '@ return an empty string on error
        on error resume next '@ switch off WSH error handling (custom error handling will be used)
            '@ add a new line of code for each child node
            for each xmlEnvSetting in objEnvironmentSettingsNode.ChildNodes
                strCode = strCode & "Environment.Value ( "& Chr(34) & xmlEnvSetting.nodeName & Chr(34) &" ) = "& Chr(34) & xmlEnvSetting.text & Chr(34) & vbNewLine
            next
            generateEnvironmentConfigurationCode = strCode '@ pass the result back to caller
    End Function

    '@ initiate a connection between QTP and QC
    Public Sub connect (strUrl, strDomain, strProject, strUserName, strPassword, blnEncrypted)
        strCurrentMethod_ = "connect" '@ note the method that will be executed - for logging
        on error resume next '@ switch off WSH error handling (custom error handling will be used)
            object.TDConnection.connect strUrl, strDomain, strProject, strUserName, strPassword, blnEncrypted '@ connect to QC project
            onConnect.fire Me '@ use fire method of the Event object; all functions (handlers) that are referenced by this object will be executed (Me object represents a current class)
    End Sub

    Public function listActions ()
        reDim arrActions (-1)
        for each objAction in object.Test.Actions
            reDim arrActions (uBound (arrActions) +1)
            arrActions (uBound (arrActions)) = objAction.name
        next
        listActions = arrActions
    end function

End class










'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'@   C L A S S   I N I T I A L I Z A T O R S  (bridge between wscript and qtp namespace)
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

function funcCreateQuickTestApplication ()
    set funcCreateQuickTestApplication = new clsQuickTestApplication
    funcCreateQuickTestApplication.onAddCodeToAction.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    funcCreateQuickTestApplication.onOpen.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
    funcCreateQuickTestApplication.onConnect.addHandler array ("logCallerPropertiesToQTPReportOnError", "logCallerPropertiesToDatabaseOnError")
end function









'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////
'@   F U N C T I O N S
'////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////////

sub exportRecordToGlobalSheet (rsToExport, arrFieldsToExport)
    dataTable.globalSheet.SetCurrentRow environment.Value ("TestIteration")
    for each objFld in rsToExport.fields
        for each strElem in arrFieldsToExport
            if uCase (strElem) = uCase (objFld.name) then
                on error resume next : dataTable.value ("a_"+ environment ("ActionName") +"_"+ objFld.name, dtGlobalSheet) = objFld.value
                if err.number <> 0 then : dataTable.globalSheet.addParameter "a_" & environment ("ActionName") +"_"+ objFld.name, objFld.value : on error goto 0
            End if
        next
    next
end sub

sub importRecordsetLogToLocalDataTable (objCaller)
    dataTable.importSheet environment ("ResultDir") &"\"& dataTable.localSheet.name, 1, dataTable.localSheet.name
end sub

function listDataSheetHeaders (strDataSheetName)
    ReDim arrParameters (-1)
    for idxColumns = 1 to dataTable.getSheet (strDataSheetName).GetParameterCount
        ReDim preserve arrParameters (UBound (arrParameters) +1)
        arrParameters (UBound (arrParameters)) = dataTable.getSheet (strDataSheetName).GetParameter (idxColumns).Name
    next
    listDataSheetHeaders = arrParameters
end function

function listDataSheetCurrentRowValues (strDataSheetName)
    ReDim arrValues (-1)
    for idxColumns = 1 to dataTable.getSheet (strDataSheetName).GetParameterCount
        ReDim preserve arrValues (UBound (arrValues) +1)
        arrValues (UBound (arrValues)) = dataTable.getSheet (strDataSheetName).GetParameter (idxColumns).Value
    next
    listDataSheetCurrentRowValues = arrValues
end function

Function fnctReturnGlobalSheetCurrentRowParamValueByInstr(strColumnPartName)
    Dim intColumnCount
    Dim strColumnName
    Dim blnResult
    blnResult = false

    intColumnCount = DataTable.GlobalSheet.GetParameterCount
    For i = 1 to intColumnCount
        strColumnName = DataTable.GlobalSheet.GetParameter(i).Name
        If InStr(strColumnName,strColumnPartName) <> 0 Then
            blnResult = DataTable.GlobalSheet.GetParameter(i)
            Exit For
        End If
    Next

    fnctReturnGlobalSheetCurrentRowParamValueByInstr = blnResult

End Function

Function fnctReplaceGlobalSheetCurrentRowParamValueByInstr(strColumnPartName, strValue)
    Dim intColumnCount
    Dim strColumnName
    Dim blnResult
    blnResult = false

    intColumnCount = DataTable.GlobalSheet.GetParameterCount
    For i = 1 to intColumnCount
        strColumnName = DataTable.GlobalSheet.GetParameter(i).Name
        If InStr(strColumnName,strColumnPartName) <> 0 Then
            DataTable(strColumnName, dtGlobalSheet) = strValue
            blnResult = true
            Exit For
        End If
    Next

    fnctReplaceGlobalSheetCurrentRowParamValueByInstr = blnResult

End Function




'*************************************************************************************************
'@Name        fnctReturnFNZTimeFormat
'@Description   Converts the standard timestamp string input to the FNZ format in the auth section
'*************************************************************************************************
Function fnctReturnFNZTimeFormat(byVal strTimeStamp)
	
	Dim strTime, strDate, strHour, intHour, strMinute, strMeridian, strAbrrevMonth, strFinalHour

	strDate = FormatDateTime(CDate(strTimeStamp),vbShortDate)
	strAbrrevMonth = MonthName(Month(strTimeStamp),true)
	'strDate = Left(strDate,2) & "-" & strAbrrevMonth & "-" & Right(strDate, 2)
    strDate = Left(strDate,2) & "-" & strAbrrevMonth & "-" & Right(strDate, 4)
	'@This may only work for the next 90 years or so for the year abbreviation

	strTime = FormatDateTime(CDate(strTimeStamp),vbShortTime)
	strHour = Left(strTime,2)
	strMinute = Right(strTime,2)
	intHour = CInt(strHour)
	If intHour > 12 Then
		intHour = intHour - 12
		strMeridian = "pm"
	ElseIf intHour = 12 Then
		strMeridian = "pm"
	Else
		strMeridian = "am"
	End If
	strFinalHour = CStr(intHour)
	strTime = strFinalHour & ":" & strMinute & " " & strMeridian & " "
	
	fnctReturnFNZTimeFormat = strDate & " " & strTime
End Function


'*************************************************************************************************
'@Name  	    fnctReturnFNZTimeFormat2YearDigits
'@Description   Converts the standard timestamp string input to the FNZ format in the user section
'*************************************************************************************************
Function fnctReturnFNZTimeFormat2YearDigits(byVal strTimeStamp)
	
	Dim strTime, strDate, strHour, intHour, strMinute, strMeridian, strAbrrevMonth, strFinalHour

	strDate = FormatDateTime(CDate(strTimeStamp),vbShortDate)
	strAbrrevMonth = MonthName(Month(strTimeStamp),true)
	strDate = Left(strDate,2) & "-" & strAbrrevMonth & "-" & Right(strDate, 2)
    'strDate = Left(strDate,2) & "-" & strAbrrevMonth & "-" & Right(strDate, 4)
	'@This may only work for the next 90 years or so for the year abbreviation

	strTime = FormatDateTime(CDate(strTimeStamp),vbShortTime)
	strHour = Left(strTime,2)
	strMinute = Right(strTime,2)
	intHour = CInt(strHour)
	If intHour > 12 Then
		intHour = intHour - 12
		strMeridian = "pm"
	ElseIf intHour = 12 Then
		strMeridian = "pm"
	Else
		strMeridian = "am"
	End If
	strFinalHour = CStr(intHour)
	strTime = strFinalHour & ":" & strMinute & " " & strMeridian & " "
	
	fnctReturnFNZTimeFormat2YearDigits = strDate & " " & strTime
End Function

'*************************************************************************************************
'@ Name         fnctFindCellWithTextInTableNoReportOnFail
'@Description   Searches a column of a table for a given text string, and returns the row
'*************************************************************************************************
Function fnctFindCellWithTextInTableNoReportOnFail(byVal strText, byVal objTable, byVal intColumn)
	Dim intRowCount, blnTextFound
	blnTextFound = false
	strText = Trim(strText)
    intRowCount = objTable.rowcount
	fnctFindCellWithTextInTableNoReportOnFail = null
	For i = 1 to intRowCount
		strCellText = Trim(objTable.GetCellData(i,intColumn))
		If strText = strCellText Then
			blnTextFound = true
			Exit For
		End If
	Next
	If blnTextFound Then
		fnctFindCellWithTextInTableNoReportOnFail = i
		Reporter.ReportEvent micPass,"strText <" & strText & "> found in table",""
	Else
		fnctFindCellWithTextInTableNoReportOnFail = -1
		'@Return an impossible result if the search failed
	End If
End Function


Function fnctReturnClosestTimeFromTable(byVal strActualTime, byVal objTable, byVal intFirstRow, byVal intColumn)
	Dim arrDateTimes(), arrDateTimeDifferences(), floatDifference, floatActualDifference, strClosestTime

	intRowCount = objTable.RowCount
	ReDim arrDateTimes(objTable.RowCount - 2)
	ReDim arrDateTimeDifferences(objTable.RowCount - 2)
	floatDifference = 1000
	For i = intFirstRow to objTable.RowCount
		k = i - 2
		arrDateTimes(k) = CDate(objTable.GetCellData(i,intColumn))
		floatActualDifference = Abs(CDate(strActualTime) - CDate(objTable.GetCellData(i,intColumn)))
		arrDateTimeDifferences(k) = floatActualDifference
		If floatActualDifference < floatDifference Then
			floatDifference = floatActualDifference
			strClosestTime = CStr(arrDateTimes(k))
		End If
	Next
	fnctReturnClosestTimeFromTable = strClosestTime
End Function


'*************************************************************************************************
'@Name          fnctFindMatchingTransactionInTransactionTableByNumberOfUnits
'@Description   Matches a number of units and a fund name to a transaction in a table
'@							Could be expanded if more data is needed to match
'************************************************************************************

Function fnctFindMatchingTransactionInTransactionTableByNumberOfUnits(objTable, strName, dblNumberOfUnits)
	Dim blnFound, intRow, strNumberOfUnits
    
	blnFound = false
	intRow = 0

	Do Until blnFound Or intRow = 0 Or intRow > objTable.RowCount

            intRow = objTable.GetRowWithCellText(strName, 2, intRow)
    
            strNumberOfUnits = objTable.GetCellData(intRow, 4)
            
            strNumberOfUnits = Round(strNumberOfUnits, 4)
            If dblNumberOfUnits = strNumberOfUnits Then
                blnFound = true
            Else
                intRow = intRow + 1
            End If
            
	Loop

	fnctFindMatchingTransactionInTransactionTableByNumberOfUnits = intRow

End Function


'*************************************************************************************************
'@Name          fnctFindMatchingTransactionInTransactionTableByValue
'@Description   Matches a number of units and a fund name to a transaction in a table
'@							Could be expanded if more data is needed to match
'************************************************************************************

Function fnctFindMatchingTransactionInTransactionTableByValue(rsTransactionData, objTable)

	blnFound = false
    intRow = -1
	Do Until blnFound Or intRow = 0 Or intRow > objTable.RowCount

        intRow = objTable.GetRowWithCellText(rsTransactionData.object.fields("strName").Value, 2, intRow)

        strValue =  objTable.GetCellData(intRow, 6)

        If rsTransactionData.object.fields("strValue").Value =  strValue Then
            blnFound = true
        Else
            intRow = intRow + 1
        End If
        
	Loop

	fnctFindMatchingTransactionInTransactionTableByValue = intRow

End Function

