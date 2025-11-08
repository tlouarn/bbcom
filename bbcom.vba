Option Explicit

Private Enum ENUM_REQUEST_TYPE
    REFERENCE_DATA = 1
    HISTORICAL_DATA = 2
End Enum

'Constants
Private Const CONST_SERVICE_TYPE_REF As String = "//blp/refdata"
Private Const CONST_SERVICE_TYPE_PRT As String = "//blp/tseapi"
Private Const CONST_REQUEST_TYPE_REFERENCE As String = "ReferenceDataRequest"
Private Const CONST_REQUEST_TYPE_HISTORICAL As String = "HistoricalDataRequest"

'Private data structures
Private bInputSecurityArray() As String
Private bInputFieldArray() As String
Private bOutputArray As Variant
Private bOverrideFieldArray() As String
Private bOverrideValueArray() As Variant

'Session Objects
Private bSession As blpapicomLib2.Session
Private bService As blpapicomLib2.Service
Private bRequest As blpapicomLib2.REQUEST
Private bEvent As blpapicomLib2.Event

'Request Data Objects
Private bRequestType As ENUM_REQUEST_TYPE
Private bSecurities() As String
Private bFields() As String
Private bStartDate As String
Private bEndDate As String
Private bErrorType As String
Private bEntityType As String
Private bEntityName As String

'Overrides Data Objects
Private bOverrides As Element
Private bOverrideFields() As String
Private bOverrideValues() As Variant
Private bCalendarCodeOverride As String
Private bCurrencyCode As String
Private bNonTradingDayFillOption As String
Private bNonTradingDayFillMethod As String
Private bPeriodicityAdjustment As String
Private bPeriodicitySelection As String
Private bMaxDataPoints As Long
Private bPricingOption As String
Private bIncludeCash As Boolean
Private bAdjustmentFollowDPDF As Boolean
Private bAdjustmentAbnormal As Boolean
Private bAdjustmentSplit As Boolean
Private bAdjustmentNormal As Boolean

Private Sub Class_Initialize()
    'Nothing
End Sub
Private Sub Class_Terminate()
    'Cleaning up or Terminate the Class without crashing Excel (in case Terminate did not trigger)
    Call Terminate
End Sub

Private Sub Terminate()
    'Cleans up the objects
    Set bOutputArray = Nothing
    Set bEvent = Nothing
    Set bRequest = Nothing
    Set bService = Nothing
    On Error Resume Next
        bSession.Stop
    On Error GoTo 0
    Set bSession = Nothing
End Sub

Public Function referenceData( _
    tickers() As String, _
    fields() As String, _
    Optional overrideFields As Variant, _
    Optional overrideValues As Variant _
) As Variant

    Dim i As Integer
    
    'Mandatory parameters
    bRequestType = REFERENCE_DATA
    bSecurities = tickers
    bFields = fields
    
    'Optional parameters
    If Not IsMissing(overrideFields) Then bOverrideFields = overrideFields
    If Not IsMissing(overrideValues) Then bOverrideValues = overrideValues
    If IsEmpty(bOverrideFields) <> IsEmpty(bOverrideValues) Then
        Err.Raise vbObjectError, "BBCOM", "Override Fields and Values are not defined properly!"
    ElseIf IsEmpty(bOverrideFields) = True Then
        If UBound(bOverrideFields, 1) <> UBound(bOverrideValues, 1) Then
            Err.Raise vbObjectError, "BBCOM", "Override Fields are of different Size!"
        End If
    End If
    
    'Process request
    ReDim bOutputArray(0 To UBound(bSecurities, 1), 0 To UBound(bFields))
    For i = 1 To UBound(bFields)
        bOutputArray(0, i) = bFields(i)
    Next i
    Call ProcessDataRequest
    
    'Return result
    referenceData = bOutputArray
    Call Terminate
    
End Function

Public Function historicalData( _
    tickers() As String, _
    fields() As String, _
    startDate As Date, _
    endDate As Date, _
    Optional calendarCodeOverride As Variant, _
    Optional currencyCode As Variant, _
    Optional nonTradingDayFillOption As Variant, _
    Optional nonTradingDayFillMethod As Variant, _
    Optional periodicityAdjustment As Variant, _
    Optional periodicitySelection As Variant, _
    Optional maxDataPoints As Variant, _
    Optional pricingOption As Variant, _
    Optional adjustmentFollowDPDF As Boolean = True, _
    Optional adjustmentAbnormal As Boolean, _
    Optional adjustmentSplit As Boolean, _
    Optional adjustmentNormal As Boolean, _
    Optional overrideFields As Variant, _
    Optional overrideValues As Variant _
) As Variant

    Dim i As Integer
    
    'Mandatory parameters
    bRequestType = HISTORICAL_DATA
    bSecurities = tickers
    bFields = fields
    bStartDate = Format(startDate, "YYYYMMDD")
    bEndDate = Format(endDate, "YYYYMMDD")
    If startDate > endDate Then
        Err.Raise vbObjectError, "BBCOM", "startDate must be < endDate"
    End If
    
    'Optional parameters
    If Not IsMissing(overrideFields) Then bOverrideFields = overrideFields
    If Not IsMissing(overrideValues) Then bOverrideValues = overrideValues
    If IsEmpty(bOverrideFields) <> IsEmpty(bOverrideValues) Then
        Err.Raise vbObjectError, "BBCOM2", "Override Fields and Values are not defined properly!"
    ElseIf IsEmpty(bOverrideFields) = True Then
        If UBound(bOverrideFields, 1) <> UBound(bOverrideValues, 1) Then
            Err.Raise vbObjectError, "BBCOM2", "Inconsistent override fields/values"
        End If
    End If
    If Not IsMissing(calendarCodeOverride) Then bCalendarCodeOverride = calendarCodeOverride              'CDR <GO>
    If Not IsMissing(currencyCode) Then bCurrencyCode = currencyCode                                      'CCY ISO-3
    If Not IsMissing(nonTradingDayFillOption) Then bNonTradingDayFillOption = nonTradingDayFillOption     'NON_TRADING_WEEKDAYS;ALL_CALENDAR_DAYS;ACTIVE_DAYS_ONLY
    If Not IsMissing(nonTradingDayFillMethod) Then bNonTradingDayFillMethod = nonTradingDayFillMethod    'PREVIOUS_VALUE;NIL_VALUE
    If Not IsMissing(periodicityAdjustment) Then bPeriodicityAdjustment = periodicityAdjustment           'ACTUAL;CALENDAR;FISCAL
    If Not IsMissing(periodicitySelection) Then bPeriodicitySelection = periodicitySelection              'DAILY;WEEKLY;MONTHLY;QUARTERLY;SEMI_ANNUALLY;YEARLY
    If Not IsMissing(maxDataPoints) Then bMaxDataPoints = maxDataPoints                                   '
    If Not IsMissing(pricingOption) Then bPricingOption = pricingOption                                  'PRINCING_OPTION_PRICE;PRICING_OPTION_YIELD
    If Not IsMissing(adjustmentFollowDPDF) Then bAdjustmentFollowDPDF = adjustmentFollowDPDF
    If Not IsMissing(adjustmentAbnormal) Then bAdjustmentAbnormal = adjustmentAbnormal
    If Not IsMissing(adjustmentSplit) Then bAdjustmentSplit = adjustmentSplit
    If Not IsMissing(adjustmentNormal) Then bAdjustmentNormal = adjustmentNormal

    'Process request
    ReDim bOutputArray(0 To UBound(bSecurities, 1), 0 To UBound(bFields) + 1) 'the +1 is to contain the Date
    bOutputArray(0, 1) = "date"
    For i = 1 To UBound(bFields)
        bOutputArray(0, i + 1) = bFields(i)
    Next i
    Call ProcessDataRequest
    
    'Return result
    historicalData = bOutputArray
    Call Terminate
    
End Function


Private Sub ProcessDataRequest()
    'Processes the Request
    Dim ValidStep As Boolean
    
    '1. Start the Session and Service
    ValidStep = OpenService
    If ValidStep = False Then
        Call Terminate
        Err.Raise vbObjectError, "BBCOM", "Error in Opening Session or Service!"
        Exit Sub
    End If
    
    '2. Send the Request
    ValidStep = SendRequest
    If ValidStep = False Then
        Call Terminate
        Err.Raise vbObjectError, "BBCOM", "Error: Could not send the Request!"
        Exit Sub
    End If
    
    '3. Process the Incoming Events
    Call catchServerEvent
    
End Sub

Private Function OpenService() As Boolean
    'Opens the Sessions and the Service
    Set bSession = New blpapicomLib2.Session
    bSession.QueueEvents = True
    On Error Resume Next
        bSession.Start
        If Err.Number <> 0 Or bSession Is Nothing Then
            OpenService = False
            Exit Function
        End If
    On Error GoTo 0
    
    On Error Resume Next
        If bRequestType = HISTORICAL_DATA Or bRequestType = REFERENCE_DATA Then
            bSession.OpenService (CONST_SERVICE_TYPE_REF)
            Set bService = bSession.GetService(CONST_SERVICE_TYPE_REF)
        End If
        If Err.Number <> 0 Or bService Is Nothing Then
            OpenService = False
            Exit Function
        End If
    On Error GoTo 0
    OpenService = True
End Function

Private Function SendRequest() As Boolean
    'Sends the Request
    Dim i As Integer, override As Element
    Select Case bRequestType
        Case REFERENCE_DATA
            Set bRequest = bService.CreateRequest(CONST_REQUEST_TYPE_REFERENCE)
            For i = 1 To UBound(bSecurities, 1)
                If bSecurities(i) <> "" Then
                    bRequest.GetElement("securities").AppendValue bSecurities(i)
                Else
                    bRequest.GetElement("securities").AppendValue "SomethingToStopCrashes"
                End If
            Next i
            For i = 1 To UBound(bFields, 1)
                If bFields(i) <> "" Then
                    bRequest.GetElement("fields").AppendValue bFields(i)
                Else
                    bRequest.GetElement("fields").AppendValue "SomethingToStopCrashes"
                End If
            Next i
            On Error Resume Next
            If UBound(bOverrideFields, 1) > 0 Then
                If Err.Number = 0 Then
                    On Error GoTo 0
                    Set bOverrides = bRequest.GetElement("overrides")
                    For i = 1 To UBound(bOverrideFields, 1)
                        Set override = bOverrides.AppendElment()
                        override.SetElement "fieldId", bOverrideFields(i)
                        override.SetElement "value", bOverrideValues(i)
                    Next i
                Else
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
            On Error Resume Next
                bSession.SendRequest bRequest
                If Err.Number <> 0 Then
                    SendRequest = False
                Else
                    SendRequest = True
                End If
            On Error GoTo 0
            
        Case HISTORICAL_DATA
            Set bRequest = bService.CreateRequest(CONST_REQUEST_TYPE_HISTORICAL)
            For i = 1 To UBound(bSecurities, 1)
                If bSecurities(i) <> "" Then
                    bRequest.GetElement("securities").AppendValue bSecurities(i)
                Else
                    bRequest.GetElement("securities").AppendValue "SomethingToStopCrashes"
                End If
            Next i
            For i = 1 To UBound(bFields, 1)
                If bFields(i) <> "" Then
                    bRequest.GetElement("fields").AppendValue bFields(i)
                Else
                    bRequest.GetElement("fields").AppendValue "SomethingToStopCrashes"
                End If
            Next i
            On Error Resume Next
            If UBound(bOverrideFields, 1) > 0 Then
                If Err.Number = 0 Then
                    On Error GoTo 0
                    Set bOverrides = bRequest.GetElement("overrides")
                    For i = 1 To UBound(bOverrideFields, 1)
                        Set override = bOverrides.AppendElment()
                        override.SetElement "fieldId", bOverrideFields(i)
                        override.SetElement "value", bOverrideValues(i)
                    Next i
                Else
                    Err.Clear
                    On Error GoTo 0
                End If
            End If
            bRequest.Set "startDate", bStartDate
            bRequest.Set "endDate", bEndDate
            If bCalendarCodeOverride <> "" Then bRequest.Set "calendarCodeOverride", bCalendarCodeOverride
            If bCalendarCodeOverride <> "" Then bRequest.Set "calendarCodeOverride", bCalendarCodeOverride
            If bCurrencyCode <> "" Then bRequest.Set "currencyCoden", bCurrencyCode
            If bNonTradingDayFillOption <> "" Then bRequest.Set "nonTradingDayFillOption", bNonTradingDayFillOption
            If bNonTradingDayFillMethod <> "" Then bRequest.Set "nonTradingDayFillMethod", bNonTradingDayFillMethod
            If bPeriodicityAdjustment <> "" Then bRequest.Set "periodicityAdjustment", bPeriodicityAdjustment
            If bPeriodicitySelection <> "" Then bRequest.Set "periodicitySelection", bPeriodicitySelection
            If bMaxDataPoints <> 0 Then bRequest.Set "maxDataPoints", bMaxDataPoints
            If bPricingOption <> "" Then bRequest.Set "maxDataPoints", bMaxDataPoints
            
            bRequest.Set "adjustmentFollowDPDF", bAdjustmentFollowDPDF
            bRequest.Set "adjustmentAbnormal", bAdjustmentAbnormal
            bRequest.Set "adjustmentSplit", bAdjustmentSplit
            bRequest.Set "adjustmentNormal", bAdjustmentNormal

            On Error Resume Next
                bSession.SendRequest bRequest
                If Err.Number <> 0 Then
                    SendRequest = False
                Else
                    SendRequest = True
                End If
            On Error GoTo 0
        
        End Select
        
End Function

Private Function catchServerEvent() As Boolean
    ' Catches the events coming from Bloomberg
    Dim TimeOut As Double: TimeOut = DateAdd("n", 5, Now()) 'MAX 5 MINUTES BEFORE TIMEOUT!
    Dim bExit As Boolean: bExit = False
    Do While Now() < TimeOut And bExit = False
        Set bEvent = bSession.NextEvent
        If (bEvent.EventType = PARTIAL_RESPONSE Or bEvent.EventType = RESPONSE) Then
            
            Select Case bRequestType
                Case ENUM_REQUEST_TYPE.REFERENCE_DATA: catchServerEvent = getServerData_reference
                Case ENUM_REQUEST_TYPE.HISTORICAL_DATA: catchServerEvent = getServerData_historical
            End Select
            
            If (bEvent.EventType = RESPONSE) Then bExit = True
        End If
    Loop
    If IsEmpty(bOutputArray) = True And Now() >= TimeOut Then
        catchServerEvent = False
    End If
End Function

Private Function getServerData_reference() As Boolean
    'Extracts the Data from a reference Data Request
    Dim i As Integer, j As Integer, k As Integer, l As Integer, Secu As String
    
    Dim it As blpapicomLib2.MessageIterator
    Set it = bEvent.CreateMessageIterator()
    
    Dim MSG As blpapicomLib2.Message, secData As blpapicomLib2.Element, _
        security As blpapicomLib2.Element, fields As blpapicomLib2.Element, _
        field As blpapicomLib2.Element, bError As blpapicomLib2.Element, _
        bBulkValues As blpapicomLib2.Element
    Dim BulkDataField() As Variant, BulkDataSubField() As Variant, SecNumber As Integer, FldNumber As Integer
    bErrorType = ""
    Do While it.Next()
        Set MSG = it.Message
        'Error Handling
        If MSG.AsElement.HasElement("responseError") Then
            Set bError = MSG.GetElement("responseError")
            bErrorType = bError.GetElement("subcategory")
            Select Case bErrorType
                Case "DAILY_LIMIT_REACHED", "MONTHLY_LIMIT_REACHED"
                    MsgBox "You have reached a data limit: " & bErrorType & vbNewLine & "Contact Bloomberg to unlock more data", vbOKOnly, "Bloomberg Wrapper"
                    getServerData_reference = False
                    Exit Function
                    
                Case "INVALID_SECURITY_IDENTIFIER", _
                     "INVALID_FIELD_DATA", _
                     "TOO_MANY_OVERRIDES", _
                     "INVALID_OVERRIDE_FIELD", _
                     "NOT_APPLICABLE_TO_REF_DATA"
                    MsgBox "Error in response detected: " & bErrorType & vbNewLine & "ABORTING", vbOKOnly, "Bloomberg Wrapper"
                    getServerData_reference = False
                    Exit Function
                    
                Case Else
                    MsgBox "Unknown Error in response detected: " & bErrorType & vbNewLine & "ABORTING", vbOKOnly, "Bloomberg Wrapper"
                    getServerData_reference = False
                    Exit Function
            End Select
        End If
        
        Set secData = MSG.AsElement.GetElement("securityData")
        For i = 1 To secData.NumValues
            Set security = MSG.GetElement("securityData").GetValue(i - 1)
            Secu = security.GetElement("security").Value
            For j = 1 To UBound(bSecurities, 1)
                If bSecurities(j) = Secu Then
                    SecNumber = j
                    bOutputArray(SecNumber, 0) = Secu
                    Exit For
                End If
            Next j
            If SecNumber = 0 Then
                'We might not find the security because it has a non-supported character
                For j = 1 To Len(Secu)
                    If Asc(Mid(Secu, j, 1)) = 63 Then
                        If j = 1 Then
                            Secu = Right(Secu, Len(Secu) - 1)
                            j = j - 1
                        ElseIf j = Len(Secu) Then
                            Secu = Left(Secu, Len(Secu) - 1)
                        Else
                            Secu = Left(Secu, j - 1) & Right(Secu, Len(Secu) - j)
                            j = j - 1
                        End If
                    End If
                Next j
                For j = 1 To UBound(bSecurities, 1)
                    If bSecurities(j) = Secu Then
                        SecNumber = j
                        bOutputArray(SecNumber, 0) = Secu
                        Exit For
                    End If
                Next j
            End If
            Set fields = security.GetElement("fieldData")
            For j = 1 To fields.NumElements
                Set field = fields.GetElement(j - 1)
                'Find where to place the data of this Field
                For k = 1 To UBound(bFields, 1)
                    If field.Name = bFields(k) Then
                        FldNumber = k
                        Exit For
                    End If
                Next k
                'Assign the Data
                If field.IsArray = False Then
                    If field.DataType = BLPAPI_INT32 Then   'Type not handled by current version of VBA
                        bOutputArray(SecNumber, FldNumber) = CInt(field.Value)
                    Else
                        bOutputArray(SecNumber, FldNumber) = field.Value
                    End If
                Else
                    'We have Bulk Data
                    ReDim BulkDataField(1 To field.NumValues)
                    For k = 1 To field.NumValues
                        Set bBulkValues = field.GetValue(k - 1)
                        ReDim BulkDataSubField(1 To bBulkValues.NumElements)
                        For l = 1 To bBulkValues.NumElements
                            If bBulkValues.GetElement(l - 1).DataType = BLPAPI_INT32 Then
                                BulkDataSubField(l) = CInt(bBulkValues.GetElement(l - 1).Value)
                            Else
                                BulkDataSubField(l) = bBulkValues.GetElement(l - 1).Value
                            End If
                        Next l
                        BulkDataField(k) = BulkDataSubField
                    Next k
                    bOutputArray(SecNumber, FldNumber) = BulkDataField
                End If
            Next j
        Next i
    Loop
    getServerData_reference = True
End Function
Private Function getServerData_historical() As Boolean
    'Extracts the Data from a historical Data Request
    Dim i As Integer, j As Integer, k As Integer, l As Integer
    Dim Secu As String
    
    Dim it As blpapicomLib2.MessageIterator
    Set it = bEvent.CreateMessageIterator()
    
    Dim MSG As blpapicomLib2.Message, security As blpapicomLib2.Element, fields As blpapicomLib2.Element, field As blpapicomLib2.Element, bError As blpapicomLib2.Element
    Dim HistData() As Variant, fieldHistData() As Variant, NumfieldHist As Integer, SecNumber As Integer, FldNumber As Integer
    bErrorType = ""
    
    Do While it.Next
        Set MSG = it.Message
        If MSG.AsElement.HasElement("responseError") Then
            Set bError = MSG.GetElement("responseError")
            bErrorType = bError.GetElement("subcategory")
            Select Case bErrorType
                Case "DAILY_LIMIT_REACHED", "MONTHLY_LIMIT_REACHED"
                    MsgBox "You have reached a data limit: " & bErrorType & vbNewLine & "Contact Bloomberg to unlock more data", vbOKOnly, "Risk Monitor - Bloomberg Data"
                    getServerData_historical = False
                    Exit Function
                    
                Case "INVALID_SECURITY_IDENTIFIER", _
                     "INVALID_START_END", _
                     "INVALID_CURRENCY", _
                     "NO_FIELDS", _
                     "TOO_MANY_OVERRIDES", _
                     "INVALID_FIELD", _
                     "INVALID_OVERRIDE_FIELD", _
                     "NOT_APPLICABLE_TO_HIST_DATA", _
                     "NOT_APPLICABLE_TO_SECTOR"
                    MsgBox "Error in response detected: " & bErrorType & vbNewLine & "ABORTING", vbOKOnly, "Bloomberg Wrapper"
                    getServerData_historical = False
                    Exit Function

                Case Else
                    MsgBox "Unknown Error in response detected: " & bErrorType & vbNewLine & "ABORTING", vbOKOnly, "Bloomberg Wrapper"
                    getServerData_historical = False
                    Exit Function
            End Select
        End If
    
        For i = 1 To MSG.AsElement.NumElements
            Set security = MSG.GetElement("securityData")
            Secu = security.GetElement("security").Value
            For j = 1 To UBound(bSecurities, 1)
                If bSecurities(j) = Secu Then
                    SecNumber = j
                    bOutputArray(SecNumber, 0) = Secu
                    Exit For
                End If
            Next j
            If SecNumber = 0 Then
                'We might not find the security because it has a non-supported character
                For j = 1 To Len(Secu)
                    If Asc(Mid(Secu, j, 1)) = 63 Then
                        If j = 1 Then
                            Secu = Right(Secu, Len(Secu) - 1)
                            j = j - 1
                        ElseIf j = Len(Secu) Then
                            Secu = Left(Secu, Len(Secu) - 1)
                        Else
                            Secu = Left(Secu, j - 1) & Right(Secu, Len(Secu) - j)
                            j = j - 1
                        End If
                    End If
                Next j
                For j = 1 To UBound(bSecurities, 1)
                    If bSecurities(j) = Secu Then
                        SecNumber = j
                        bOutputArray(SecNumber, 0) = Secu
                        Exit For
                    End If
                Next j
            End If
            Set fields = security.GetElement("fieldData")
            If fields.NumValues > 0 Then
                ReDim HistData(1 To fields.NumValues)
                For j = 1 To fields.NumValues
                    Set field = fields.GetValue(j - 1)
                    'Find where to place the data of this Field
                    ReDim fieldHistData(1 To UBound(bFields, 1) + 1)
                    For k = 1 To field.NumElements
                        If field.GetElement(k - 1).Name = "date" Then
                            fieldHistData(1) = field.GetElement(k - 1).Value
                        Else
                            For l = 1 To UBound(bFields, 1)
                                If field.GetElement(k - 1).Name = bFields(l) Then
                                    If field.GetElement(k - 1).DataType = BLPAPI_INT32 Then
                                        fieldHistData(l + 1) = CInt(field.GetElement(k - 1).Value)
                                    Else
                                        fieldHistData(l + 1) = field.GetElement(k - 1).Value
                                    End If
                                End If
                            Next l
                        End If
                    Next k
                    HistData(j) = fieldHistData
                Next j
                'Assign HistData to the bOutputArray
                ReDim fieldHistData(1 To fields.NumValues)
                For j = 1 To UBound(bFields, 1) + 1
                    For k = 1 To fields.NumValues
                        fieldHistData(k) = HistData(k)(j)
                    Next k
                    bOutputArray(SecNumber, j) = fieldHistData
                Next j
            End If
        Next i
    Loop
    getServerData_historical = True
End Function

