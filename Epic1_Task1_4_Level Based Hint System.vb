Option Explicit

Sub PlayGuessTheNumberGameWithLevels()
    Dim level As Integer
    Dim maxLevel As Integer
    Dim minNumber As Integer
    Dim maxNumber As Integer
    Dim maxAttempts As Integer
    Dim targetNumber As Integer
    Dim userGuess As Integer
    Dim guessCount As Integer
    Dim userInput As String
    Dim hintUsed As Boolean
    
    ' Initialize game parameters
    level = 1
    maxLevel = 5  ' Total number of levels
    
    Do While level <= maxLevel
        ' Set range and attempts based on level
        minNumber = 1
        maxNumber = 50 * level  ' Increase range with level
        maxAttempts = 10 - (level - 1)  ' Decrease attempts as level increases
        
        ' Generate random target number
        Randomize
        targetNumber = Int((maxNumber - minNumber + 1) * Rnd + minNumber)
        
        ' Game instructions for current level
        MsgBox "Level " & level & " Start!" & vbCrLf & _
               "I'm thinking of a number between " & minNumber & " and " & maxNumber & "." & vbCrLf & _
               "You have " & maxAttempts & " attempts to guess it correctly." & vbCrLf & _
               "Type 'hint' to receive a hint (limited per level).", vbInformation
        
        ' Initialize guess counter and hint usage
        guessCount = 0
        hintUsed = False
        
        ' Main guessing loop for current level
        Do While guessCount < maxAttempts
            userInput = InputBox("Enter your guess (" & minNumber & " to " & maxNumber & "):", "Guess the Number - Level " & level)
            
            ' Check if the user requested a hint
            If LCase(userInput) = "hint" Then
                If Not hintUsed Then
                    ProvideHint level, targetNumber, minNumber, maxNumber
                    hintUsed = True
                Else
                    MsgBox "You have already used your hint for this level.", vbExclamation
                End If
                Continue Do  ' Skip incrementing guessCount
            End If
            
            ' Validate input
            If IsNumeric(userInput) Then
                userGuess = CInt(userInput)
                guessCount = guessCount + 1
                
                ' Check the guess
                If userGuess = targetNumber Then
                    MsgBox "Congratulations! You've guessed the correct number (" & targetNumber & ") in " & guessCount & " attempts!", vbExclamation
                    Exit Do  ' Proceed to next level
                ElseIf userGuess < targetNumber Then
                    MsgBox "Too low! Try again.", vbInformation
                ElseIf userGuess > targetNumber Then
                    MsgBox "Too high! Try again.", vbInformation
                End If
            Else
                MsgBox "Invalid input. Please enter a number or type 'hint'.", vbCritical
            End If
        Loop
        
        ' Check if the player guessed correctly
        If userGuess = targetNumber Then
            If level = maxLevel Then
                MsgBox "Congratulations! You've completed all levels!", vbExclamation
                Exit Sub
            Else
                MsgBox "Get ready for Level " & (level + 1) & "!", vbInformation
                level = level + 1
            End If
        Else
            ' Player failed to guess within attempts
            MsgBox "Sorry! You've used all " & maxAttempts & " attempts. The correct number was " & targetNumber & ".", vbExclamation
            MsgBox "Game Over! You reached Level " & level & ".", vbExclamation
            Exit Sub
        End If
    Loop
End Sub

Sub ProvideHint(currentLevel As Integer, targetNumber As Integer, minNum As Integer, maxNum As Integer)
    Select Case currentLevel
        Case 1
            MsgBox "Hint: The number is " & IIf(targetNumber Mod 2 = 0, "even.", "odd."), vbInformation
        Case 2
            MsgBox "Hint: The number is divisible by " & GetDivisor(targetNumber) & ".", vbInformation
        Case 3
            MsgBox "Hint: The number is between " & (targetNumber - 10) & " and " & (targetNumber + 10) & ".", vbInformation
        Case 4
            MsgBox "Hint: The sum of the digits is " & GetDigitSum(targetNumber) & ".", vbInformation
        Case 5
            MsgBox "Hint: The number is a " & GetPrimeStatus(targetNumber) & ".", vbInformation
        Case Else
            MsgBox "No hints available for this level.", vbInformation
    End Select
End Sub

Function GetDivisor(number As Integer) As Integer
    ' Returns a divisor of the number other than 1 and itself
    Dim i As Integer
    For i = 2 To number \ 2
        If number Mod i = 0 Then
            GetDivisor = i
            Exit Function
        End If
    Next i
    GetDivisor = 1  ' If prime, return 1
End Function

Function GetDigitSum(number As Integer) As Integer
    ' Returns the sum of the digits of the number
    Dim sum As Integer
    sum = 0
    Do While number > 0
        sum = sum + (number Mod 10)
        number = number \ 10
    Loop
    GetDigitSum = sum
End Function

Function GetPrimeStatus(number As Integer) As String
    ' Returns whether the number is prime or not
    Dim i As Integer
    If number < 2 Then
        GetPrimeStatus = "composite number."
        Exit Function
    End If
    For i = 2 To Sqr(number)
        If number Mod i = 0 Then
            GetPrimeStatus = "composite number."
            Exit Function
        End If
    Next i
    GetPrimeStatus = "prime number."
End Function
