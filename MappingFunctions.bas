Attribute VB_Name = "Module1"
Option Explicit

Function mapDistance(address1 As String, address2 As String, apikey, Optional mode As String = "driving")
    Dim httpObject As Object
    Set httpObject = CreateObject("MSXML2.XMLHTTP")

    Dim encodedAddress1, encodedAddress2 As String
    encodedAddress1 = WorksheetFunction.EncodeURL(address1)
    encodedAddress2 = WorksheetFunction.EncodeURL(address2)
    
    
    Dim sURL As String
    sURL = "https://maps.googleapis.com/maps/api/directions/json?origin=" & encodedAddress1 & "&destination=" & encodedAddress2 & "&key=" & apikey & "&mode=" & mode
    
    Debug.Print "test"
    
    Debug.Print sURL
    
    Dim sRequest, sGetResult As String
    sRequest = sURL
    httpObject.Open "GET", sRequest, False
    httpObject.Send
    sGetResult = httpObject.ResponseText
    
    'Debug.Print sGetResult

    Dim vJSON
    Dim sState As String
    
    JSON.Parse sGetResult, vJSON, sState
    If sState = "Error" Then MsgBox "Invalid JSON": End
    'Debug.Print vJSON
    
    'Going to default to route 0, but this could be changed or customized later
    
    Dim distanceVal
    Dim distanceExists
    
    jsonExt.selectElement vJSON, ".routes[0].legs[0].distance.text", distanceVal, distanceExists
    
    Debug.Print distanceVal
    'MsgBox "Completed"
    
    mapDistance = distanceVal

End Function

Function mapDistanceRawVal(address1 As String, address2 As String, apikey, Optional mode As String = "driving")
    Dim httpObject As Object
    Set httpObject = CreateObject("MSXML2.XMLHTTP")

    Dim encodedAddress1, encodedAddress2 As String
    encodedAddress1 = WorksheetFunction.EncodeURL(address1)
    encodedAddress2 = WorksheetFunction.EncodeURL(address2)
    
    
    Dim sURL As String
    sURL = "https://maps.googleapis.com/maps/api/directions/json?origin=" & encodedAddress1 & "&destination=" & encodedAddress2 & "&key=" & apikey & "&mode=" & mode
    
    Debug.Print "test"
    
    Debug.Print sURL
    
    Dim sRequest, sGetResult As String
    sRequest = sURL
    httpObject.Open "GET", sRequest, False
    httpObject.Send
    sGetResult = httpObject.ResponseText
    
    'Debug.Print sGetResult

    Dim vJSON
    Dim sState As String
    
    JSON.Parse sGetResult, vJSON, sState
    If sState = "Error" Then MsgBox "Invalid JSON": End
    'Debug.Print vJSON
    
    'Going to default to route 0, but this could be changed or customized later
    
    Dim distanceVal
    Dim distanceExists
    
    'Returns value in meters
    jsonExt.selectElement vJSON, ".routes[0].legs[0].distance.value", distanceVal, distanceExists
    
    Debug.Print distanceVal
    'MsgBox "Completed"
    
    mapDistanceRawVal = distanceVal

End Function

Function mapTime(address1 As String, address2 As String, apikey, Optional mode As String = "driving")
    Dim httpObject As Object
    Set httpObject = CreateObject("MSXML2.XMLHTTP")

    Dim encodedAddress1, encodedAddress2 As String
    encodedAddress1 = WorksheetFunction.EncodeURL(address1)
    encodedAddress2 = WorksheetFunction.EncodeURL(address2)
    
    Dim sURL As String
    sURL = "https://maps.googleapis.com/maps/api/directions/json?origin=" & encodedAddress1 & "&destination=" & encodedAddress2 & "&key=" & apikey & "&mode=" & mode
    
    Debug.Print "test"
    
    Debug.Print sURL
    
    Dim sRequest, sGetResult As String
    sRequest = sURL
    httpObject.Open "GET", sRequest, False
    httpObject.Send
    sGetResult = httpObject.ResponseText
    
    'Debug.Print sGetResult

    Dim vJSON
    Dim sState As String
    
    JSON.Parse sGetResult, vJSON, sState
    If sState = "Error" Then MsgBox "Invalid JSON": End
    'Debug.Print vJSON
    
    'Going to default to route 0, but this could be changed or customized later
    
    Dim durationVal
    Dim durationExists
    
    jsonExt.selectElement vJSON, ".routes[0].legs[0].duration.text", durationVal, durationExists
    
    Debug.Print durationVal
    'MsgBox "Completed"
    
    mapTime = durationVal

End Function
Function mapTimeRawVal(address1 As String, address2 As String, apikey, Optional mode As String = "driving")
    Dim httpObject As Object
    Set httpObject = CreateObject("MSXML2.XMLHTTP")

    Dim encodedAddress1, encodedAddress2 As String
    encodedAddress1 = WorksheetFunction.EncodeURL(address1)
    encodedAddress2 = WorksheetFunction.EncodeURL(address2)
    
    
    Dim sURL As String
    sURL = "https://maps.googleapis.com/maps/api/directions/json?origin=" & encodedAddress1 & "&destination=" & encodedAddress2 & "&key=" & apikey & "&mode=" & mode
    
    Debug.Print "test"
    
    Debug.Print sURL
    
    Dim sRequest, sGetResult As String
    sRequest = sURL
    httpObject.Open "GET", sRequest, False
    httpObject.Send
    sGetResult = httpObject.ResponseText
    
    'Debug.Print sGetResult

    Dim vJSON
    Dim sState As String
    
    JSON.Parse sGetResult, vJSON, sState
    If sState = "Error" Then MsgBox "Invalid JSON": End
    'Debug.Print vJSON
    
    'Going to default to route 0, but this could be changed or customized later
    
    Dim durationVal
    Dim durationExists
    
    jsonExt.selectElement vJSON, ".routes[0].legs[0].duration.value", durationVal, durationExists
    
    Debug.Print durationVal
    'MsgBox "Completed"
    
    mapTimeRawVal = durationVal

End Function
Function mapAllVal(address1 As String, address2 As String, apikey, Optional mode As String = "driving")
    Dim httpObject As Object
    Set httpObject = CreateObject("MSXML2.XMLHTTP")

    Dim encodedAddress1, encodedAddress2 As String
    encodedAddress1 = WorksheetFunction.EncodeURL(address1)
    encodedAddress2 = WorksheetFunction.EncodeURL(address2)
    
    
    Dim sURL As String
    sURL = "https://maps.googleapis.com/maps/api/directions/json?origin=" & encodedAddress1 & "&destination=" & encodedAddress2 & "&key=" & apikey & "&mode=" & mode
    
    'Debug.Print "mapAllVal"
    
    'Debug.Print sURL
    
    Dim sRequest, sGetResult As String
    sRequest = sURL
    httpObject.Open "GET", sRequest, False
    httpObject.Send
    sGetResult = httpObject.ResponseText
    
    'Debug.Print sGetResult

    Dim vJSON
    Dim sState As String
    
    JSON.Parse sGetResult, vJSON, sState
    If sState = "Error" Then MsgBox "Invalid JSON": End
    'Debug.Print vJSON
    
    'Going to default to route 0, but this could be changed or customized later
    
    Dim distanceVal
    Dim distanceExists
    
    jsonExt.selectElement vJSON, ".routes[0].legs[0].distance.text", distanceVal, distanceExists
    
    Dim distanceRawValVal
    Dim distanceRawValExists
    
    jsonExt.selectElement vJSON, ".routes[0].legs[0].distance.value", distanceRawValVal, distanceRawValExists
    
    Dim durationVal
    Dim durationExists
    
    jsonExt.selectElement vJSON, ".routes[0].legs[0].duration.text", durationVal, durationExists
    
    Dim durationRawValVal
    Dim durationRawValExists
    
    jsonExt.selectElement vJSON, ".routes[0].legs[0].duration.value", durationRawValVal, durationRawValExists
    
    Debug.Print distanceVal & ":" & distanceRawValVal & ":" & durationVal & ":" & durationRawValVal
    'MsgBox "Completed"
    
    mapAllVal = distanceVal & ":" & distanceRawValVal & ":" & durationVal & ":" & durationRawValVal
    

End Function

Function mapClosest(address1 As String, addresses As Range, apikey, Optional mode As String = "driving")
    'Debug.Print "Row count: " & addresses.Rows.Count
    'Indexed from 1? Who does that???
    'Debug.Print addresses(1)
    
    Dim Responses()
    ReDim Responses(addresses.Rows.Count)
    Dim closest
    Dim shortestTime As Long
    shortestTime = 2147483647
    'Debug.Print shortestTime
    Dim closestAddress
    
    Dim encodedAddress1, encodedAddress2 As String
    encodedAddress1 = WorksheetFunction.EncodeURL(address1)
    
    Dim i As Integer
    For i = 1 To addresses.Rows.Count
        Dim httpObject As Object
        Set httpObject = CreateObject("MSXML2.XMLHTTP")
        
        encodedAddress2 = WorksheetFunction.EncodeURL(addresses(i))
        
        Dim sURL As String
        sURL = "https://maps.googleapis.com/maps/api/directions/json?origin=" & encodedAddress1 & "&destination=" & encodedAddress2 & "&key=" & apikey & "&mode=" & mode
        
        'Debug.Print "test"
        
        'Debug.Print sURL
        
        Dim sRequest, sGetResult As String
        sRequest = sURL
        httpObject.Open "GET", sRequest, False
        httpObject.Send
        sGetResult = httpObject.ResponseText
        
        Responses(i) = sGetResult
        
        Dim vJSON
        Dim sState As String
        
        JSON.Parse sGetResult, vJSON, sState
        If sState = "Error" Then MsgBox "Invalid JSON": End
        
        Dim durationRawValVal
        Dim durationRawValExists
    
        jsonExt.selectElement vJSON, ".routes[0].legs[0].duration.value", durationRawValVal, durationRawValExists
        
        If (durationRawValVal < shortestTime) Then
            shortestTime = durationRawValVal
            closest = i
            closestAddress = addresses(i)
            'Debug.Print "test in if"
        End If
    Next i
    
    JSON.Parse Responses(closest), vJSON, sState
    Dim durationRawValVal2
    Dim durationRawValExists2
    
    jsonExt.selectElement vJSON, ".routes[0].legs[0].duration.value", durationRawValVal2, durationRawValExists2
    
    'Debug.Print "Shortest duration from json " & durationRawValVal2
    'Debug.Print "Shortest time val: " & shortestTime / 60
    
    'Range to normal VBA array
    'Dim X
    
    'X = Application.Transpose(addresses)
    'Debug.Print "X" & X(2)

    mapClosest = (shortestTime / 60) & ":" & closestAddress
    
    'Debug.Print sGetResult

End Function

Sub test()

    MsgBox mapAllVal("Disneyland", "Universal Studios Hollywood", "APIKEYGOESHERE")
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"
Debug.Print "No API Key for you!"



End Sub
