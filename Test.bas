Attribute VB_Name = "Test"

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

