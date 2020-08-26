Sub WFH_CheckIn_Email_To_FileContact()
    Dim obApp As Object
    Dim NewMail As MailItem
 
    Set obApp = Outlook.Application
    Set NewMail = obApp.CreateItem(olMailItem)

    
    Dim FileNum As Integer
    Dim DataLine As String
    Dim ToAddress As String
    Dim contactList As String
    
    
    FileNum = FreeFile()
    Open "c:\contactList.txt" For Input As #FileNum
    
    While Not EOF(FileNum)
        Line Input #FileNum, DataLine ' read in data 1 line at a time
        ' decide what to do with dataline,
        ' depending on what processing you need to do for each case
        ToAddress = ToAddress & ";" & DataLine
    Wend
       
    With NewMail
         .Subject = "WFH - " & Day(Now) & "-" & MonthName(Month(Now), False) & "-" & Year(Now)
         .To = ToAddress
         .Body = "Time In:  " & FormatTime(Now, "True")
         .Display
    End With
 
    Set obApp = Nothing
    Set NewMail = Nothing
End Sub

Sub WFH_CheckIn_Email_To_Contact()
    Dim obApp As Object
    Dim NewMail As MailItem
 
    Set obApp = Outlook.Application
    Set NewMail = obApp.CreateItem(olMailItem)
    
    Dim contactList As String
        
    contactList = GetAddresses
   
    With NewMail
         .Subject = "WFH - " & Day(Now) & "-" & MonthName(Month(Now), False) & "-" & Year(Now)
         .To = contactList
         .Body = "Time In:  " & FormatTime(Now, "True")
         .Display
    End With
 
    Set obApp = Nothing
    Set NewMail = Nothing
End Sub


Function FormatTime(dateIn, asTwelveHourFormat)
  Dim result, h, M, t
  If IsDate(dateIn) Then
    If Len(asTwelveHourFormat & "") < 1 Then asTwelveHourFormat = False
    h = Hour(dateIn)
    M = Minute(dateIn)
    t = ""
    If asTwelveHourFormat Then
      t = " "
      If h >= 12 Then
        t = t & "PM"
        If h > 12 Then h = h Mod 12
      Else
        t = t & "AM"
      End If
    End If
    h = Right("00" & h, 2)
    M = Right("00" & M, 2)
    result = h & ":" & M & t
  Else
    result = "NA"
  End If
  FormatTime = result
End Function


Sub ReplyAll_WFH()
    Dim olItem As Outlook.MailItem
    Dim olReply As MailItem ' Reply
    Dim olRecip As Recipient ' Add Recipient
    
    Dim checkInTime As String
    Dim timeDifference As String
    Dim checkInDate As String
    Dim checkInDateTime As Date
    Dim timeDifferenceInHours As String
        
    Dim oContact As Outlook.ContactItem
    
            
    Dim empCode As String
    empCode = InputBox("Enter Employee Code", "(K.U(r)) Employee Code", "00279")
    
    checkInDate = FormatDateTime(ActiveExplorer().Selection.item(1).SentOn, vbShortDate)
        
    If Not IsEmpty(empCode) And Len(empCode & "") > 1 Then
      
        For Each olItem In Application.ActiveExplorer.Selection
            Set olReply = olItem.ReplyAll
            Set olRecip = olReply.Recipients.Add("crkr-admin@contour-software.com") ' Recipient Address
            
            If (InStr(olReply.Body, "Time In:")) > 0 Then
                olRecip.Type = olCC
                checkInTime = getTimeFromBody(olReply.Body)
                ' Create Complete Check In Date Time
                ' ==================================
                
                If IsEmpty(checkInTime) Then
                    MsgBox ("Please, enter correct check-in time")
                    checkInTime = InputBox("Enter Check-in time", "(K.U(r)) Check-in date time", "12:00 PM")
                End If
                
                checkInDateTime = CDate(checkInDate & " " & checkInTime)
                
                timeDifference = DateDiff("h", checkInDateTime, Now)
                
                If (timeDifference > 24 Or timeDifference < 1) Then
                               
                MsgBox ("Please, enter correct check-in complete date and time")
                checkInDateTime = InputBox("Enter Check-in date/time", "(K.U(r)) Check-in date time", Now)
                
                End If
                
                timeDifference = DateDiff("n", checkInDateTime, Now)
                
                timeDifferenceInHours = timeDifference / 60
                
                olReply.Body = "Hi," & vbCrLf & vbCrLf & _
                            "Time In:" & "                  " & checkInTime & vbCrLf & _
                            "Time Out:" & "               " & FormatTime(Now, "True") & vbCrLf & _
                            "Hours: " & "                    " & Round(timeDifferenceInHours, 2) & vbCrLf & _
                            "Employee Code:    " & empCode & vbCrLf
                            'olReply.Body
                olReply.Display
            Else
                MsgBox ("Please, select your last Time-In email and then click")
            End If
            'olReply.Send
        Next olItem
        
    End If
End Sub

Function getTimeFromBody(strIn As String) As String
    Dim result As String
    Dim allMatches As Object
    Dim objRegex As Object
    Set objRegex = CreateObject("vbscript.regexp")
    
    objRegex.Pattern = "Time In:\s*((1[0-2]|0?[1-9]):([0-5][0-9]) ?([AaPp][Mm]))"
    objRegex.Global = True
    objRegex.IgnoreCase = True
    
    Set allMatches = objRegex.Execute(strIn)

    If allMatches.Count <> 0 Then
       result = allMatches.item(0).submatches.item(0)
    End If
    
    getTimeFromBody = result
   
End Function



Private Function GetAddresses() As String

Dim o, AddressList, AddressEntry
Dim AddressName As String
Dim result As String

Set o = CreateObject("Outlook.Application")

Set AddressList = o.Session.AddressLists("Contacts")
    For Each AddressEntry In AddressList.AddressEntries
        result = result & ExtractEmailFun(CStr(AddressEntry)) & ";"
    Next AddressEntry
 GetAddresses = result & "crkr-admin@contour-software.com"
End Function



Function ExtractEmailFun(extractStr As String) As String
'Update by extendoffice
Dim CharList As String
On Error Resume Next
CheckStr = "[A-Za-z0-9._-]"
OutStr = ""
Index = 1
Do While True
    Index1 = VBA.InStr(Index, extractStr, "@")
    getStr = ""
    If Index1 > 0 Then
        For p = Index1 - 1 To 1 Step -1
            If Mid(extractStr, p, 1) Like CheckStr Then
                getStr = Mid(extractStr, p, 1) & getStr
            Else
                Exit For
            End If
        Next
        getStr = getStr & "@"
        For p = Index1 + 1 To Len(extractStr)
            If Mid(extractStr, p, 1) Like CheckStr Then
                getStr = getStr & Mid(extractStr, p, 1)
            Else
                Exit For
            End If
        Next
        Index = Index1 + 1
        If OutStr = "" Then
            OutStr = getStr
        Else
            OutStr = OutStr & Chr(10) & getStr
        End If
    Else
        Exit Do
    End If
Loop
ExtractEmailFun = OutStr
End Function
