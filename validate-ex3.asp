<%@ Language=VBScript %>
<!DOCTYPE html>
<html>
<head>
    <title>Saved Form Data</title>
<link href="buy.css" type="text/css" rel="stylesheet" />
</head>
<body>
    <% 
    ' Function to check if a field is empty
    Function IsFieldEmpty(fieldValue)
        IsFieldEmpty = Len(Trim(fieldValue)) = 0 ' Check if the field is empty or only contains spaces
    End Function
    
    ' Function to display error message and link to try again
    Sub ShowErrorMessage(errorMessage)
    %>
        <h1>Error: <%= errorMessage %></h1>
        <a href="index.html">Try again</a>
    <% 
    End Sub
    
    ' Function to perform Luhn Algorithm validation
    Function IsLuhnValid(cardNumber)
        Dim sum, i, digit, isEven
        sum = 0
        isEven = False
        
        For i = Len(cardNumber) To 1 Step -1
            digit = CInt(Mid(cardNumber, i, 1))
            
            If isEven Then
                digit = digit * 2
                If digit >= 10 Then
                    digit = digit - 9
                End If
            End If
            
            sum = sum + digit
            isEven = Not isEven
        Next
        
        IsLuhnValid = (sum Mod 10 = 0)
    End Function
    
    ' Checking if form fields are empty
    Dim fullName, section, creditCardNumber, cardType
    fullName = Request.Form("fullName")
    section = Request.Form("section")
    creditCardNumber = Request.Form("creditCardNumber")
    cardType = Request.Form("cardType")
    
    If IsFieldEmpty(fullName) Or IsFieldEmpty(section) Or IsFieldEmpty(creditCardNumber) Or IsFieldEmpty(cardType) Then
        ShowErrorMessage("Incomplete Form: Please fill in all fields before submitting.")
    ElseIf Len(creditCardNumber) <> 16 Or Not IsLuhnValid(creditCardNumber) Then
        ShowErrorMessage("Invalid Credit Card: Please provide a valid credit card number.")
    Else
        ' All fields are filled and credit card is valid by Luhn algorithm, proceed with form data processing
        
        ' Format the data
        Dim formData
        formData = fullName & ";" & section & ";" & creditCardNumber & ";" & cardType
        
        ' Save form data to file validates.txt
        Dim fs, file
        Set fs = Server.CreateObject("Scripting.FileSystemObject")
        Set file = fs.OpenTextFile(Server.MapPath("validates.txt"), 8, True)
        file.WriteLine(formData)
        file.Close
        
        ' Read contents of validates.txt and display in <pre> element
        Dim fileContents, readFile
        Set readFile = fs.OpenTextFile(Server.MapPath("validates.txt"), 1)
        fileContents = readFile.ReadAll
        readFile.Close
    %>
    
     <!-- Display the success message and recorded information -->
    <h1>Thanks, sucker!</h1>
    <p>Your information has been recorded.</p>
    
    <h2>Name</h2>
    <p><%= fullName %></p>
    
    <h2>Credit Card</h2>
    <p><%= creditCardNumber %> (<%= cardType %>)</p>
    
    <% End If %>
</body>
</html>