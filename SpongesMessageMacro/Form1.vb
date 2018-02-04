Imports Tulpep.NotificationWindow

Public Class Form1
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short
    Dim objPopup As New PopupNotifier
    Dim filePath As String = "..\..\..\Macros\macros.txt"
    Dim keyList(11) As String 'full list of macro keys available
    Dim regKeys(11) As String 'list of currently registered macro keys
    Dim regMacros(11) As String 'list of currently registered macros
    Dim notifications As Boolean = True 'notifications are enabled by default

    'Returns the keycode pressed after GetAsyncKeyState is triggered, else return 0
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

    'Determines whether a key press was within a certain range of keyboard keys
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

    'Finds the macro key associated with a line from the macro text document or finds function key assigned code
    Private Function GetHotkey(ByVal input As String, ByVal code As Integer) As String
        Dim alfa As String = "abcdefghijkl"
        If code = 0 Then
            Return keyList(alfa.IndexOf(Strings.Right(input, 1)))
        ElseIf code = 1 Then
            Return GetChar(alfa, Array.IndexOf(keyList, input))
        Else
            Return Nothing
        End If
    End Function

    'Notifications system
    Private Sub Notifier(ByVal title As String, ByVal content As String)
        If notifications Then
            objPopup.TitleText = title
            objPopup.ContentText = content
            objPopup.Popup()
        End If
    End Sub

    'Initializes the macros from the macro text document
    Private Sub InitMacros()
        Dim increment As Integer = 0
        Dim lineText As String = ""
        FileOpen(1, filePath, OpenMode.Input)
        ComboBox2.Items.Clear()
        ComboBox3.Items.Clear()
        Do Until EOF(1)
            lineText = LineInput(1)
            ComboBox2.Items.Add(lineText)
            ComboBox3.Items.Add(GetHotkey(lineText, 0))
            regKeys(increment) = GetHotkey(lineText, 0)
            regMacros(increment) = lineText
            increment += 1
            If Not regKeys.Contains(keyList(increment)) Then
                ComboBox1.Items.Add(keyList(increment))
            End If
        Loop
        FileClose(1)
    End Sub

    'Key listener and macro copier
    Private Sub TmrHotkey_Tick(sender As Object, e As EventArgs) Handles Timer1.Tick
        Dim keyPressed As Boolean = KeyPressRange(Keys.F1, Keys.F12)
        Dim keyCode As Integer = GetKeyPressed(Keys.F1, Keys.F12)
        If keyPressed Then
            'Opens macro file for reading
            FileOpen(1, filePath, OpenMode.Input)
            Dim keyName As String = [Enum].GetName(GetType(System.Windows.Forms.Keys), CType(keyCode, System.Windows.Forms.Keys))
            If regKeys.Contains(keyName.ToString()) Then
                Do Until GetHotkey(LineInput(1), 0) = keyName.ToString()
                    If GetHotkey(LineInput(1), 0) = keyName.ToString() Then
                        My.Computer.Clipboard.SetText(Strings.Left(LineInput(1), Len(LineInput(1)) - 1))
                        'Sends success notification
                        Notifier("Action Complete", "Macro successfully copied to clipboard")
                    End If
                Loop
                FileClose(1)
                While KeyPressRange(Keys.F1, Keys.F12)
                End While
            End If
        End If
    End Sub

    'Required functions ran when program starts
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Getting and setting the macro key and identifier for future use in the programs runtime.* All possible keys.
        For counter = 0 To 11
            keyList(counter) = "F" & (counter + 1)
        Next
        'Initialize macros
        InitMacros()
        'Initialize key listener
        Timer1.Start()
    End Sub

    'Handles notifications settings
    Private Sub CheckBox1_CheckedChanged(sender As Object, e As EventArgs) Handles CheckBox1.CheckedChanged
        If CheckBox1.Checked Then
            notifications = True
        Else
            notifications = False
        End If
    End Sub

    'Clears message label
    Private Sub ComboBox1_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox1.SelectedIndexChanged
        Label5.Text = ""
    End Sub

    'Switches macros in the macro editor
    Private Sub ComboBox2_SelectedIndexChanged(sender As Object, e As EventArgs) Handles ComboBox2.SelectedIndexChanged
        ComboBox3.Enabled = True
        ComboBox3.SelectedItem = regKeys(Array.IndexOf(regMacros, ComboBox2.SelectedItem))
        TextBox2.Text = ComboBox2.SelectedItem
    End Sub

    'Adds macros
    Private Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        If Not ComboBox1.SelectedItem = Nothing Then
            Dim decision As Integer = Nothing
            FileOpen(1, filePath, OpenMode.Input)
            If TextBox1.Text = "" Then
                MsgBox("Macro cannot be blank!", "ERROR")
                FileClose(1)
                Exit Sub
            ElseIf regMacros.Contains(TextBox1.Text) Then
                decision = MsgBox("You already have a macro with the same content. Are you sure you want to continue?", MessageBoxButtons.YesNo, "Conflict")
                If decision = DialogResult.No Then
                    FileClose(1)
                    Exit Sub
                End If
            End If
            FileClose(1)
            'Append the new macro
            FileOpen(1, filePath, OpenMode.Append)
            PrintLine(1, TextBox1.Text & GetHotkey(ComboBox1.SelectedItem, 1) & vbCrLf)
            FileClose(1)
        End If
    End Sub

    'Edits or deletes macros
    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        If Not ComboBox2.SelectedItem = Nothing Then
            'If the user deletes the macro
            Dim fileText As String = ""
            Dim lineText As String = ""
            If TextBox2.Text = "" Then
                FileOpen(1, filePath, OpenMode.Input)
                Do Until ComboBox3.SelectedItem = GetHotkey(lineText, 0)
                    lineText = LineInput(1)
                    If ComboBox3.SelectedItem = GetHotkey(lineText, 0) Then
                        lineText = Nothing
                    End If
                    fileText += lineText & vbCrLf
                Loop
                FileClose(1)

                'fileText = Strings.Left(fileText, Len(fileText) - 2)
                FileOpen(1, filePath, OpenMode.Output)
                PrintLine(1, fileText)
                Label9.Text = "Macro successfully deleted!"
                Label9.ForeColor = Color.Green

                FileClose(1)
            End If

            'Reinitialize macros
            InitMacros()
        Else
            MsgBox("Select a macro from the drop down list and try again!")
        End If
    End Sub
End Class