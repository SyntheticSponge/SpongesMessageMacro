Imports Tulpep.NotificationWindow

Public Class Form1
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short
    Dim objPopup As New PopupNotifier
    Dim filePath As String = "..\..\..\Macros\macros.txt"
    Dim fileLines As ArrayList
    Dim pressedKey As String
    Dim fileKeyList(11, 1) As String 'list of keys attached to macros in the file
    Dim keyList(11, 1) As String 'full list of macro keys available
    Dim macroListChanging(11) As String 'Loads list of macros as full text (line and key) to prepare to change
    Dim notifications As Boolean = True
    Dim alfa As String = "abcdefghijkl"
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Label5.Text = ""
        Label5.ForeColor = Color.Black
    End Sub

    'Returns the key pressed when GetAsyncKeyState is triggered, else return 0
    Private Function GetKeyPressed(ByVal intStart As Integer, ByVal intEnd As Integer) As Integer
        If Not (intStart = Nothing Or intEnd = Nothing) Then
            For keyCode = intStart To intEnd
                If GetAsyncKeyState(keyCode) Then
                    Return keyCode
                End If
            Next
        End If
        Return 0
    End Function

    Private Function KeyPressRange(ByVal intStart As Integer, ByVal intEnd As Integer) As Boolean
        If Not (intStart = Nothing Or intEnd = Nothing) Then
            For keyCode = intStart To intEnd
                If GetAsyncKeyState(keyCode) Then
                    Return True
                End If
            Next
        End If
        Return False
    End Function

    Private Sub TmrHotkey_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim macroKeyPressed As Boolean
        If KeyPressRange(Keys.F1, Keys.F12) Then
            'Sends success notification
            Notifier("Action Complete", "Macro successfully copied to clipboard")
            'Open macro file for reading
            FileOpen(1, filePath, OpenMode.Input)
            'Gets the key code which is pressed
            Dim keyPressed As Integer = GetKeyPressed(Keys.F1, Keys.F12)
            Select Case keyPressed
                Case Keys.F1
                    Do Until Strings.

                    Loop
                Case Keys.F2
                Case Keys.F3
                Case Keys.F4
                Case Keys.F5
                Case Keys.F6
                Case Keys.F7
                Case Keys.F8
                Case Keys.F9
                Case Keys.F10
                Case Keys.F11
                Case Keys.F12
                Case Else
                    MsgBox("Great job you broke the program!")
            End Select
            Dim FileLine As String = Nothing

            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim MacroKey As String = FileLine.Substring(FileLine.Count() - 1)
                'Gets and sets the macroline from the file without identifier key
                If MacroKey = pressedKey Then
                    Dim macroLine As String = Strings.Left(FileLine, Len(FileLine) - 1)
                    My.Computer.Clipboard.SetText(macroLine)
                    macroKeyPressed = True
                    Exit Do
                End If
            Loop
            FileClose(1)
            If macroKeyPressed = False Then
                'Label8.Text = "Debug: Macro key is wrong, posted- " & MacroKey
                Notifier("Action Failed", "No macro found, try adding a new macro to the pressed key and try again.")
                Label8.Text = "Debug: F1 F Pressed"
            End If

            While KeyPressRange(Keys.F1, Keys.F12)
            End While
        End If

    End Sub


    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim newMacroKey As String = ComboBox1.SelectedItem
        Dim addToFile As String = ""
        Dim FileLine As String
        FileOpen(1, filePath, OpenMode.Input)
        Dim macroUsed As Boolean
        If TextBox1.Text = "" Then
            Label5.Text = "Error: Macro cannot contain less than 1 char."
            Label5.ForeColor = Color.Crimson
            FileClose(1)
            Exit Sub
        End If
        Do Until EOF(1)
            FileLine = LineInput(1)
            Dim LineSize As Integer = FileLine.Count()
            Dim subCount As Integer = LineSize - 1
            If subCount < 0 Then
                Label8.Text = "Debug: subCount < 0"
                FileClose(1)
                Exit Sub
            End If
            Dim MacroKey As String = FileLine.Substring(subCount)
            For intCount1 = 0 To 11
                If fileKeyList(intCount1, 0) = newMacroKey Then
                    Label5.Text = "Error! You already have a macro set to that key!"
                    Label5.ForeColor = Color.Crimson
                    macroUsed = True
                    Exit Do
                End If
            Next
        Loop
        If macroUsed = False Then
            Dim newLine As String = TextBox1.Text
            FileClose(1)
            FileOpen(1, filePath, OpenMode.Append)
            For intCount1 = 0 To 11
                If keyList(intCount1, 0) = newMacroKey Then
                    fileKeyList(intCount1, 0) = newMacroKey
                    fileKeyList(intCount1, 1) = keyList(intCount1, 1)
                    newLine += keyList(intCount1, 1)
                    Exit For
                End If
            Next

            PrintLine(1, newLine)
            Label5.Text = "Macro successfully added!"
            Label5.ForeColor = Color.Green
            FileClose(1)
        End If
        FileClose(1)
    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Getting and setting the macro key and identifier for future use in the programs runtime.* All possible keys.
        For intCount1 = 0 To 11
            keyList(intCount1, 0) = "F" & (intCount1 + 1)
        Next
        'Getting and setting current macro and identifier keys used in file.
        Dim curLine As Integer = 0
        FileOpen(1, filePath, OpenMode.Input)
        Do Until EOF(1)
            Dim fileKey As String
            'Gets the last character of a string, this is the attached macro identifier.
            Dim curLineText As String = LineInput(1)
            fileKey = curLineText.Substring(Len(curLineText) - 1)
            'Debugs to see if the correct line is being printed and if the correct character is being got
            'Label8.Text = "Debug: curLine prints as- " & curLine & " " & fileKey

            'Sets the identifying keys, a-l
            fileKeyList(curLine, 1) = fileKey
            For intCount1 = 0 To 11
                If keyList(intCount1, 1) = fileKey Then
                    fileKeyList(curLine, 0) = keyList(intCount1, 0)
                End If
            Next
            curLine += 1
        Loop
        FileClose(1)
        'Getting and setting macro text for later use in program runtime.
        FileOpen(1, filePath, OpenMode.Input)
        curLine = 0
        Do Until EOF(1)
            Dim curLineText = LineInput(1)
            macroListChanging(curLine) = curLineText
            ComboBox2.Items.Add(Strings.Left(curLineText, Len(curLineText) - 1))
            If curLineText.Substring(Len(curLineText) - 1) = fileKeyList(curLine, 1) Then
                ComboBox3.Items.Add(fileKeyList(curLine, 0))
            End If
            curLine += 1
        Loop
        FileClose(1)
        Timer1.Start()
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged, ComboBox3.SelectedIndexChanged
        TextBox2.Text = ComboBox2.SelectedItem
        ComboBox3.SelectedIndex = ComboBox2.SelectedIndex
        ComboBox2.SelectedIndex = ComboBox3.SelectedIndex
    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        'If the user deletes the macro
        Dim macroError As Boolean = False
        Dim fileText As String = ""
        Dim lineThing As String = ""
        Dim lineText As String = ""
        Dim fullFileText As String = ""
        If TextBox2.Text = "" Then
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                lineText = LineInput(1)
                lineThing += lineText & (ComboBox3.SelectedItem = keyList(alfa.IndexOf(Strings.Right(lineText, 1)), 0)) & vbCrLf
                Label10.Text = lineThing
                If Not (ComboBox3.SelectedItem = keyList(alfa.IndexOf(Strings.Right(lineText, 1)), 0)) Then
                    fileText += lineText & vbCrLf
                End If
                fullFileText += lineText & vbCrLf
            Loop
            FileClose(1)
            If fullFileText = fileText Then
                Label9.Text = "Error: Macro already deleted OR no macro selected!"
                Label9.ForeColor = Color.Crimson
                macroError = True
            End If
            If macroError = False Then
                fileText = Strings.Left(fileText, Len(fileText) - 2)
                FileOpen(1, filePath, OpenMode.Output)
                PrintLine(1, fileText)
                Label9.Text = "Macro successfully deleted!"
                Label9.ForeColor = Color.Green
            End If
            FileClose(1)
        End If
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            notifications = True
        Else
            notifications = False
        End If
    End Sub

    Private Sub Notifier(ByVal title As String, ByVal content As String)
        If notifications = True Then
            objPopup.TitleText = title
            objPopup.ContentText = content
            objPopup.Popup()
        End If
    End Sub
End Class