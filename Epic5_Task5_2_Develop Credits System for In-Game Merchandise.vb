Option Explicit

Dim ball As Shape
Dim paddle As Shape
Dim ballSpeedX As Single, ballSpeedY As Single
Dim paddleSpeed As Single
Dim gameOn As Boolean
Dim credits As Integer

Sub StartGame()
    ' Initialize game settings
    ballSpeedX = 2
    ballSpeedY = -2
    paddleSpeed = 10
    credits = 3  ' Set initial credits
    gameOn = True
    
    ' Create ball
    Set ball = ActiveSheet.Shapes.AddShape(msoShapeOval, 100, 100, 20, 20)
    ball.Fill.ForeColor.RGB = RGB(255, 0, 0)
    
    ' Create paddle
    Set paddle = ActiveSheet.Shapes.AddShape(msoShapeRectangle, 150, 300, 80, 10)
    paddle.Fill.ForeColor.RGB = RGB(0, 0, 255)
    
    ' Display initial credits
    UpdateCredits
    
    ' Run the game loop
    Do While gameOn
        MoveBall
        DoEvents
        Application.Wait Now + TimeValue("00:00:00.02")
    Loop
End Sub

Sub MoveBall()
    ' Update ball position
    ball.Left = ball.Left + ballSpeedX
    ball.Top = ball.Top + ballSpeedY
    
    ' Ball collision with walls
    If ball.Left <= 0 Or ball.Left + ball.Width >= ActiveSheet.UsedRange.Width Then
        ballSpeedX = -ballSpeedX
    End If
    If ball.Top <= 0 Then
        ballSpeedY = -ballSpeedY
    End If
    
    ' Ball collision with paddle
    If (ball.Top + ball.Height >= paddle.Top) And _
       (ball.Left + ball.Width >= paddle.Left) And _
       (ball.Left <= paddle.Left + paddle.Width) Then
        ballSpeedY = -ballSpeedY
    End If
    
    ' Ball touches the bottom of the screen - Lose a credit
    If ball.Top + ball.Height >= ActiveSheet.UsedRange.Height Then
        LoseCredit
    End If
End Sub

Sub MovePaddleLeft()
    ' Move paddle left
    If paddle.Left > 0 Then
        paddle.Left = paddle.Left - paddleSpeed
    End If
End Sub

Sub MovePaddleRight()
    ' Move paddle right
    If paddle.Left + paddle.Width < ActiveSheet.UsedRange.Width Then
        paddle.Left = paddle.Left + paddleSpeed
    End If
End Sub

Sub LoseCredit()
    ' Reduce credits by 1 and reset ball position
    credits = credits - 1
    UpdateCredits
    
    If credits <= 0 Then
        ' End the game if no credits are left
        GameOver
    Else
        ' Reset ball position if credits are still available
        ball.Left = 100
        ball.Top = 100
        ballSpeedX = 2
        ballSpeedY = -2
    End If
End Sub

Sub UpdateCredits()
    ' Display credits on the sheet
    Range("A1").Value = "Credits: " & credits
End Sub

Sub GameOver()
    ' End the game
    gameOn = False
    MsgBox "Game Over! No credits left.", vbExclamation
    ' Cleanup
    ball.Delete
    paddle.Delete
End Sub
