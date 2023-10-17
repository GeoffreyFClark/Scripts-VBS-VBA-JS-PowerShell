' XMLHTTP object for HTTP requests
Dim objXMLHTTP, strURL, strResponse

strURL = "http://www.somerandomwebsite.com"

Set objXMLHTTP = CreateObject("MSXML2.ServerXMLHTTP.6.0")

' Send GET request to the specified URL
objXMLHTTP.Open "GET", strURL, False
objXMLHTTP.setRequestHeader "User-Agent", "Mozilla/5.0"
objXMLHTTP.send ""

' Check for good response
If objXMLHTTP.Status = 200 Then
    ' Store response
    strResponse = objXMLHTTP.responseText
    
    ' Can add parsing of response here

    ' Can display response with
    MsgBox strResponse
Else
    MsgBox "Error fetching the content."
End If

' Cleanup
Set objXMLHTTP = Nothing
