Attribute VB_Name = "VBA-ESRI-REST"
Option Compare Database
Option Explicit
' @name QueryLocator
' @author Bill DeVoe - william.devoe@maine.gov - bdevoe@gmail.com
' @description Queries an ESRI REST locator service using a Street, City, and State.
' @param URL {String} The URL to the locator REST service; for example, the ESRI World Locator would
' be: http://geocode.arcgis.com/arcgis/rest/services/World
' @param Street {String} The street to pass to the locator service.
' @param City {String} The city to pass to the locator service.
' @param State {String} The state to pass to the locator service.
' @return {String} A semicolon delimited string containing latitude, longitude, score, and found
' address for the highest scoring result from the locator service.
' @references Microsoft Scripting Runtime
' @depends VBA-JSON from VBA-JSON v2.3.1 (c) Tim Hall - https://github.com/VBA-tools/VBA-JSON
Function QueryLocator(URL As String, Street As String, City As String, State As String) As String
    Dim xmlhttp As Object
    Set xmlhttp = CreateObject("MSXML2.serverXMLHTTP")
    Dim query As String
    Dim json_Text As String
    ' Build query to REST API
    On Error GoTo Access
    ' This method works with Excel
        query = URL + "/GeocodeServer/findAddressCandidates?" _
            & "Street=" & Application.EncodeURL(Street) _
            & "&City=" & Application.EncodeURL(City) _
            & "&State=" & Application.EncodeURL(State) _
            & "&Single+Line+Input=&category=&outFields=&maxLocations=&outSR=4326&searchExtent=&location=&distance=&magicKey=&f=pjson"
        GoTo Execute
Access:
    ' This method with Access
        query = URL + "/GeocodeServer/findAddressCandidates?" _
                & "Street=" & Application.HtmlEncode(Street) _
                & "&City=" & Application.HtmlEncode(City) _
                & "&State=" & Application.HtmlEncode(State) _
                & "&Single+Line+Input=&category=&outFields=&maxLocations=&outSR=4326&searchExtent=&location=&distance=&magicKey=&f=pjson"
Execute:
    On Error GoTo errorHandler
    ' Execute query
    xmlhttp.Open "GET", query, False
    xmlhttp.Send
    If xmlhttp.Status = 200 Then
       json_Text = xmlhttp.responseText
    Else ' Could not connect to service, 404 error or 400 bad request - ie, no internet connection or invalid url
       MsgBox xmlhttp.Status & ": " & xmlhttp.statusText
       MsgBox "Invalid Response From Server"
       Exit Function
    End If
    
    ' Parse the response JSON
    Dim json As Object
    Set json = ParseJson(json_Text)
    
    ' Check the length of the candidates array
    Dim address_count As Integer
    address_count = 0
    Dim Value As Dictionary
    For Each Value In json("candidates")
        address_count = address_count + 1
    Next Value
    ' If it's empty, return no result
    If address_count = 0 Then
        QueryLocator = "NA"
        Exit Function
    End If
    ' Parse lat, long and score
    Dim Lat As String
    Dim Lon As String
    Dim score As String
    Dim found_address As String
    Lat = json("candidates")(1)("location")("y")
    Lon = json("candidates")(1)("location")("x")
    score = json("candidates")(1)("score")
    found_address = json("candidates")(1)("address")
    QueryLocator = Lat & ";" & Lon & ";" & score & ";" & found_address
    Exit Function
    'In case of error
errorHandler:
    MsgBox "Error in: " & Err.Source & "  Description: " & Err.Description
End Function


' @name SpatialIntersect
' @author Bill DeVoe - william.devoe@maine.gov - bdevoe@gmail.com
' @descripton Passes a lat/lon value to an ESRI REST Service of polygon data, returning the value/values
'   of the given field name for polygons the point intersects.
' @dependencies
'    - VBA-JSON from https://github.com/VBA-tools/VBA-JSON
'    - Alpha Array from https://www.tek-tips.com/viewthread.cfm?qid=1134076
' @param Lat {Single} latitude of the point
' @param Lon  {Single} longitude of the point
' @param Service {String} URL to an ESRI REST Service containing polygon data
' @param Field {String} Name of the field within the service that will be used
'   to return intersecting features
' @return {String} A string representing the value/values from the field parameter of
'   polygons intersecting the input point
Public Function SpatialIntersect(ByVal Lat As Single, _
                                ByVal Lon As Single, _
                                ByVal Service As String, _
                                ByVal Field As String) As String
    On Error GoTo errorHandler
     Dim xhr As Object
     Dim query As String
     Dim thisRequest As String
     Dim json_Text As String
     Dim json As Object
     ' Change this value if you want to return a value other than "NA" when a match is not found. This prevents duplicate
     ' calls to the REST API for points that are outside of an area, by inserting a value indicating a match was not found
     Dim NA As String
     NA = "NA"
     
     ' Build query to REST API
     query = "/query?geometry=" & Lon & "%2C" & Lat & "&geometryType=esriGeometryPoint&inSR=4326&spatialRel=esriSpatialRelIntersects&outFields=" & Field & "&returnGeometry=false&f=geojson"
     thisRequest = Service & query
    
     ' Use late binding
     Set xhr = CreateObject("Microsoft.XMLHTTP")
     xhr.Open "GET", thisRequest, False
     xhr.Send
     If xhr.Status = 200 Then
       json_Text = xhr.responseText
     Else ' Could not connect to service, 404 error or 400 bad request - ie, no internet connection or invalid url
       MsgBox xhr.Status & ": " & xhr.statusText
       MsgBox "Invalid Response From Server, Unable to Intersect Point"
       Exit Function
     End If
     Set xhr = Nothing
     
     ' Parse the response JSON
     Set json = ParseJson(json_Text)
     Dim Value As Dictionary
     Dim result As String
     Dim results() As String
     Dim i As Long
     i = 0
     ' For each feature get the feature property for the field specified and push it into the results array
     For Each Value In json("features")
        result = Value("properties")(Field)
        ReDim Preserve results(i)
        results(i) = result
        i = i + 1
    Next Value
    ' If no results
    If Len(Join(results)) = 0 Then
        SpatialIntersect = NA
    Else ' Sort and collapse to string
        SpatialIntersect = Join(AlphaArray(results), "/")
    End If
    Exit Function
errorHandler:
    MsgBox "Error in: " & Err.Source & "  Description: " & Err.Description
End Function

