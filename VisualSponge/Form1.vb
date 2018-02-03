Imports Tulpep.NotificationWindow

Public Class Form1
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short
    Dim objPopup As New PopupNotifier
    Dim filePath As String = "C:\Users\Synth\Desktop\Sponge's Macro\macros.txt"
    Dim fileLines As ArrayList
    Dim pressedKey As String
    Dim fileKeyList(11, 1) As String 'list of keys attached to macros in the file
    Dim keyList(11, 1) As String 'full list of macro keys available
    Dim macroListChanging(11) As String 'Loads list of macros as full text (line and key) to prepare to change
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Label5.Text = ""
        Label5.ForeColor = Color.Black

    End Sub

    Private Sub TmrHotkey_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim macroKeyPressed As Boolean
        If GetAsyncKeyState(Keys.F1) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F1 Pressed"
            Dim FileLine As String
            pressedKey = "a"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F1 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F1)
            End While
        End If
        If GetAsyncKeyState(Keys.F2) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F2 Pressed"
            Dim FileLine As String
            pressedKey = "b"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F2 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F2)
            End While
        End If
        If GetAsyncKeyState(Keys.F3) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F3 Pressed"
            Dim FileLine As String
            pressedKey = "c"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F3 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F3)
            End While
        End If
        If GetAsyncKeyState(Keys.F4) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F4 Pressed"
            Dim FileLine As String
            pressedKey = "d"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F4 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F4)
            End While
        End If
        If GetAsyncKeyState(Keys.F5) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F5 Pressed"
            Dim FileLine As String
            pressedKey = "e"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F5 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F5)
            End While
        End If
        If GetAsyncKeyState(Keys.F6) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F6 Pressed"
            Dim FileLine As String
            pressedKey = "f"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F6 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F6)
            End While
        End If
        If GetAsyncKeyState(Keys.F7) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F7 Pressed"
            Dim FileLine As String
            pressedKey = "g"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F7 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F7)
            End While
        End If
        If GetAsyncKeyState(Keys.F8) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F8 Pressed"
            Dim FileLine As String
            pressedKey = "h"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F8 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F8)
            End While
        End If
        If GetAsyncKeyState(Keys.F9) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F9 Pressed"
            Dim FileLine As String
            pressedKey = "i"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F9 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F9)
            End While
        End If
        If GetAsyncKeyState(Keys.F10) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F10 Pressed"
            Dim FileLine As String
            pressedKey = "j"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F10 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F10)
            End While
        End If
        If GetAsyncKeyState(Keys.F11) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F11 Pressed"
            Dim FileLine As String
            pressedKey = "k"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F11 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F11)
            End While
        End If
        If GetAsyncKeyState(Keys.F12) Then
            objPopup.TitleText = "Action Complete"
            objPopup.ContentText = "Macro successfully copied to clipboard"
            Label8.Text = "Debug: F12 Pressed"
            Dim FileLine As String
            pressedKey = "l"
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                FileLine = LineInput(1)
                Dim LineSize As Integer = FileLine.Count()
                Dim subCount As Integer = LineSize - 1
                Dim MacroKey As String = FileLine.Substring(subCount)
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
                objPopup.TitleText = "Action Failed"
                objPopup.ContentText = "No macro found, try adding a new macro to the pressed key and try again."
                Label8.Text = "Debug: F12 F Pressed"
            End If
            If CheckBox1.Checked Then
                objPopup.Popup()
            End If
            While GetAsyncKeyState(Keys.F12)
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

    Private Sub TextBox1_TextChanged(sender As Object, e As EventArgs) Handles TextBox1.TextChanged

    End Sub

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Getting and setting the macro key and identifier for future use in the programs runtime.* All possible keys.
        For intCount1 = 0 To 11
            For intCount2 = 0 To 1
                If intCount2 = 0 Then
                    keyList(intCount1, 0) = "F" & (intCount1 + 1)
                End If
                If intCount2 = 1 Then
                    If intCount1 = 0 Then
                        keyList(intCount1, 1) = "a"
                    End If
                    If intCount1 = 1 Then
                        keyList(intCount1, 1) = "b"
                    End If
                    If intCount1 = 2 Then
                        keyList(intCount1, 1) = "c"
                    End If
                    If intCount1 = 3 Then
                        keyList(intCount1, 1) = "d"
                    End If
                    If intCount1 = 4 Then
                        keyList(intCount1, 1) = "e"
                    End If
                    If intCount1 = 5 Then
                        keyList(intCount1, 1) = "f"
                    End If
                    If intCount1 = 6 Then
                        keyList(intCount1, 1) = "g"
                    End If
                    If intCount1 = 7 Then
                        keyList(intCount1, 1) = "h"
                    End If
                    If intCount1 = 8 Then
                        keyList(intCount1, 1) = "i"
                    End If
                    If intCount1 = 9 Then
                        keyList(intCount1, 1) = "j"
                    End If
                    If intCount1 = 10 Then
                        keyList(intCount1, 1) = "k"
                    End If
                    If intCount1 = 11 Then
                        keyList(intCount1, 1) = "l"
                    End If
                End If
            Next intCount2
        Next intCount1
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

    Private Sub TabPage1_Click(sender As Object, e As EventArgs) Handles TabPage1.Click

    End Sub

    Private Sub Label1_Click(sender As Object, e As EventArgs) Handles Label1.Click

    End Sub

    Private Sub TableLayoutPanel1_Paint(sender As Object, e As PaintEventArgs)

    End Sub

    Private Sub TextBox2_TextChanged(sender As Object, e As EventArgs) Handles TextBox2.TextChanged

    End Sub

    Private Sub TabPage2_Click(sender As Object, e As EventArgs) Handles TabPage2.Click
    End Sub

    Private Sub Label5_Click(sender As Object, e As EventArgs) Handles Label5.Click
    End Sub

    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged, ComboBox3.SelectedIndexChanged
        TextBox2.Text = ComboBox2.SelectedItem
        ComboBox3.SelectedIndex = ComboBox2.SelectedIndex
        ComboBox2.SelectedIndex = ComboBox3.SelectedIndex
    End Sub

    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        Dim curLineText As String
        Dim MacroKey As String
        Dim curLine As Integer = 0
        FileOpen(1, filePath, OpenMode.Input)
        Do Until EOF(1)
            curLineText = LineInput(1)
            MacroKey = curLineText.Substring(Len(curLineText) - 1)
            If TextBox2.Text = "" Then
                If fileKeyList(indexInt, 0) = ComboBox3.SelectedItem Then
                    For curInt = curLine To 10
                        macroListChanging(curInt) = macroListChanging(curInt + 1)
                    Next
                    macroListChanging(11) = ""
                    Label9.Text = "Macro Successfully deleted!"
                    Label9.ForeColor = Color.Green
                    FileClose(1)
                End If
            Else
                If fileKeyList(indexInt, 0) = ComboBox3.SelectedItem Then
                    macroListChanging(curLine) = TextBox2.Text
                    Label9.Text = "Macro Successfully Edited!"
                    Label9.ForeColor = Color.Green
                    FileClose(1)
                End If
            End If
            curLine += 1
        Loop
        FileClose(1)
    End Sub
End Class

'Where I left off: Trying to figure out how to get the fileKeyList value from the identifier to the key.
