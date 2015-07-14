Attribute VB_Name = "Pushover"
'PushOver request. see: https://pushover.net/api
Const PUSHOVER_URL As String = "https://api.pushover.net/1/messages.json"

' Sends a post via PushOver with optional options
' @param In app as String: The application's token
' @param In user/group as String: The user/group token
' @param In message as String: The message that you want to send
' @return as String(): True if the message was sent successfully, otherwise false
Public Function Post(ByVal app As String, ByVal group As String, ByVal message As String) As Boolean
    Dim response As String
    
    response = PostOp(app, group, message)
    If InStr(response, """status"":1") > 0 Then
        Post = True
    Else
        Debug.Print response
    End If
End Function

' Sends a post via PushOver with optional options
' @param In app as String: The application's token
' @param In user/group as String: The user/group token
' @param In message as String: The message that you want to send
' @param In options as String Array: Other options that you can send to Pushover message like 'title' and 'sound'.
' @return as String(): The post response
' @remark see more about parameters on: https://pushover.net/api
Public Function PostOp(ByVal app As String, ByVal group As String, ByVal message As String, ParamArray options() As Variant) As String
    Dim xhttp As Object, params As String, i As Integer
    
    If (UBound(options) + 1) Mod 2 = 1 Then
        Debug.Print ("option needs to have [key, value]")
        PostOp = Empty
        Exit Function
    End If

    params = StringFormat("token={0}&user={1}&message={2}", app, group, message)
    
    For i = 0 To UBound(options) Step 2
        params = params & StringFormat("&{0}={1}", options(i), options(i + 1))
    Next i
    
    Set xhttp = CreateObject("MSXML2.ServerXMLHTTP")
    With xhttp
        .Open "POST", PUSHOVER_URL, False
        .setRequestHeader "Content-type", "application/x-www-form-urlencoded"
        .send params
         
        PostOp = .responseText
    End With
End Function


' ----- copied from ExcelUtils.
' ----- from: https://github.com/makah/ExcelHelper
'
' Generates a string using .NET format, i.e. {0}, {1}, {2} ...
' @param In strValue as String: A composite format string that includes one or more format items
' @param In arrParames as Variant: Zero or more objects to format.
' @return as String: A copy of format in which the format items have been replaced by the string representations of the corresponding arguments.
' @example: Debug.Print StringFormat("My name is {0} {1}. Hey!", "Mauricio", "Arieira")
Private Function StringFormat(ByVal strValue As String, ParamArray arrParames() As Variant) As String
    Dim i As Integer

    For i = LBound(arrParames()) To UBound(arrParames())
        strValue = Replace(strValue, "{" & CStr(i) & "}", CStr(arrParames(i)))
    Next

    StringFormat = strValue
End Function
