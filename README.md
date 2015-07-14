Pushover-vba
============

VBA pushover module for https://pushover.net/api.

Usage
=====
```vba
Sub Example1()
    Dim myApp As String, theGroup As String
    
    myApp = "my_application_token"
    theGroup = "my_Group_or_user_token"
    
    Pushover.Post _
        app:=myApp, _
        group:=theGroup, _
        message:="Hello my PushoOver-VBA Helper"
End Sub

Sub Example2()
    Dim myApp As String, theGroup As String
    
    myApp = "my_application_token"
    theGroup = "my_Group_or_user_token"
    
    ok = Pushover.Post(myApp, theGroup, "Hello my PushoOver-VBA Helper 2")
            
    MsgBox IIf(ok, "Poshover OK", "Pushover Error!")
End Sub

Sub Example3()
    Dim myApp As String, theGroup As String
    
    myApp = "my_application_token"
    theGroup = "my_Group_or_user_token"
    message = "Hello my PushoOver-VBA Helper 3"
    
    Debug.Print Pushover.PostOp(myApp, theGroup, message, "sound", "cosmic", "title", "My new Title")
End Sub
```
   
This library has two diferent ways to post. Post() that is very simple and only returns true or false, and PostOp(), a function that returns the JSON response and allow user to add optional parameters.

Dependencies
============
* None

Copyright notice
============

This software is published under MIT license 

Permission is hereby granted, free of charge, to any person obtaining a copy of this software and associated documentation files (the "Software"), to deal in the Software without restriction, including without limitation the rights to use, copy, modify, merge, publish, distribute, sublicense, and/or sell copies of the Software, and to permit persons to whom the Software is furnished to do so, subject to the following conditions:

The above copyright notice and this permission notice shall be included in all copies or substantial portions of the Software.

THE SOFTWARE IS PROVIDED "AS IS", WITHOUT WARRANTY OF ANY KIND, EXPRESS OR IMPLIED, INCLUDING BUT NOT LIMITED TO THE WARRANTIES OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE AND NONINFRINGEMENT. IN NO EVENT SHALL THE AUTHORS OR COPYRIGHT HOLDERS BE LIABLE FOR ANY CLAIM, DAMAGES OR OTHER LIABILITY, WHETHER IN AN ACTION OF CONTRACT, TORT OR OTHERWISE, ARISING FROM, OUT OF OR IN CONNECTION WITH THE SOFTWARE OR THE USE OR OTHER DEALINGS IN THE SOFTWARE.
