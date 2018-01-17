Private Sub Application_ItemSend(ByVal Item As Object, Cancel As Boolean)

If TypeOf Item Is MailItem Then

    Dim myOlApp As Outlook.Application
    Dim myNamespace As Outlook.NameSpace
    Set myOlApp = CreateObject("Outlook.Application")
    Set myNamespace = myOlApp.GetNamespace("MAPI")
     
      Dim re As Object
      Dim s As String
	  
      Const aPat As String = "attach"
      Const ePat As String = "enclosed"
    
    Set myMailToSend = Item
    Set re = CreateObject("vbscript.regexp")
        re.Pattern = aPat
    Set se = CreateObject("vbscript.regexp")
        se.Pattern = ePat
    
    s = LCase(myMailToSend.HTMLBody) & " " & LCase(myMailToSend.Subject)
    
    If se.Test(s) = True Or re.Test(s) = True Then
        If myMailToSend.Attachments.Count = 0 Then
			' Non-breaking space used to as divider between current/previous message
            If InStr(1, s, "<o:p>&nbsp;</o:p>") > InStr(1, s, ePat) Or InStr(1, s, "<o:p>&nbsp;</o:p>") > InStr(1, s, aPat)Then
				answer = MsgBox("Text in this email indicates that you may have intended to attach a file, but there are no files attached.  Do you still want to send this email?", vbYesNo)
				If answer = vbNo Then
					Cancel = True
				End If
            End If
        End If
    End If
    
    'SSN filter
    'inital + last four filter
      
    On Error Resume Next
        
    Dim mc As Object, sc As Object
    Const ssnPat As String = "\b\d{3}-\d{2}-\d{4}\b"
    Const lfPat As String = "[A-WY-Za-wy-z]" & "\d{4}"
    Const acNumPat As String = "a1c"
    Const acFormatPat As String = "\b\d\b" & "." & "\b\d\b"
         
    Set myMailToSend = Item
    Set re = CreateObject("vbscript.regexp")
    Set sc = CreateObject("vbscript.regexp")
	
    re.Pattern = lfPat
    
    If re.Test(s) = True Then
        Set mc = re.Execute(s)
        result = MsgBox("This email may have PII/PHI included. Are you sure you still want to send it? See First Initial of Last Name + Last Four: " & mc(0), vbYesNo)
        
        If result = vbNo Then
         Cancel = True
        End If
    End If
    
    re.Pattern = ssnPat
    
    If re.Test(s) = True Then
        Set mc = re.Execute(s)
        result = MsgBox("This email may have PII/PHI included. Are you sure you still want to send it? See SSN: " & mc(0), vbYesNo)
        
        If result = vbNo Then
         Cancel = True
        End If
    End If
        
    re.Pattern = acNumPat
    sc.Pattern = acFormatPat
        
    s = LCase(myMailToSend.Body)
        
    If re.Test(s) = True And sc.Test(s) = True Then
        Set mc = sc.Execute(s)
        If (Abs(InStr(1, s, "a1c") - InStr(1, s, mc(0))) < 30) Then
            result = MsgBox("This email may have PII/PHI included. Are you sure you still want to send it? See A1C level: A1C " & mc(0), vbYesNo)
         If result = vbNo Then
                Cancel = True
         End If
        End If
        
    End If
    
End If

End Sub
