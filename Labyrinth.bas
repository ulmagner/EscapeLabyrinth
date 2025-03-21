Attribute VB_Name = "Module1"
Sub lab()
    Dim PlayerRow As Integer, PlayerColumn As Integer
    Dim NewR As Integer, NewC As Integer
    Dim Direction As String
    Dim DestinationR As Integer, DestinationC As Integer

    DestinationR = 15
    DestinationC = 29
    PlayerRow = 15
    PlayerColumn = 12
    While (1)
        If PlayerRow = DestinationR And PlayerColumn = DestinationC Then
            Exit Sub
        End If
        NewR = 0
        NewC = 0
        Direction = InputBox("Next step :")
    
        If Direction = "Up" Then
            NewR = -1
        ElseIf Direction = "Down" Then
            NewR = 1
        ElseIf Direction = "Right" Then
            NewC = 1
        ElseIf Direction = "Left" Then
            NewC = -1
        ElseIf Direction = "Exit" Then
            Exit Sub
        End If
        
        If Cells(PlayerRow + NewR, PlayerColumn + NewC).Interior.Color = RGB(0, 0, 0) Then
            MsgBox ("Impossible Direction.")
        Else
            Cells(PlayerRow + NewR, PlayerColumn + NewC) = Cells(PlayerRow, PlayerColumn).Value
            Cells(PlayerRow, PlayerColumn) = ""
            If NewR = 1 Then
                PlayerRow = PlayerRow + NewR
            ElseIf NewC = 1 Then
                PlayerColumn = PlayerColumn + NewC
            End If
        End If
    Wend
End Sub
