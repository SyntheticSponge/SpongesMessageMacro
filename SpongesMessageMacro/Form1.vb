Imports System.Environment
Imports System.IO

Public Class FrmMacro
    Private Declare Function GetAsyncKeyState Lib "user32" (ByVal vKey As Integer) As Short
    Dim oldMacro As String = ""
    Dim interval As Integer = 0
    Dim macroFile As String = "macros.txt"
    Dim MyAppData As String = Path.Combine(GetFolderPath(SpecialFolder.ApplicationData), "SMM") 'SMM Data folder
    Dim filePath As String = Path.Combine(MyAppData, macroFile)
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
            Return GetChar(alfa, Array.IndexOf(keyList, input) + 1)
        Else
            Return Nothing
        End If
    End Function

    Private Sub BtnToggle(ByVal btnObject As Button)
        If txtAdd.Text = Nothing Or cmbAddKey.Text = Nothing Then
            btnObject.Enabled = False
        Else
            btnObject.Enabled = True
        End If
    End Sub

    'Notifications system
    Private Sub Notifier(ByVal title As String, ByVal content As String)
        If notifications Then
            NotifyIcon.Icon = New System.Drawing.Icon("app.ico")
            NotifyIcon.BalloonTipIcon = ToolTipIcon.Info
            NotifyIcon.Text = title
            NotifyIcon.BalloonTipTitle = title
            NotifyIcon.BalloonTipText = content
            NotifyIcon.ShowBalloonTip(1000)
        End If
    End Sub

    'Initializes the macros from the macro text document
    Private Sub InitMacros()
        Dim increment As Integer = 0
        Dim lineText As String = ""
        Dim fileText As String = ""
        cmbAddKey.Items.Clear()
        cmbList.Items.Clear()
        cmbEditKey.Items.Clear()
        interval = 0
        oldMacro = 0
        cmbEditKey.Enabled = False
        txtEdit.Enabled = False
        txtEdit.Text = Nothing
        btnDel.Enabled = False
        btnEdit.Enabled = False

        'Loads macro file into ram and resets drop down lists
        FileOpen(1, filePath, OpenMode.Input)
        Do Until EOF(1)
            lineText = LineInput(1)
            If Not lineText = "" Then
                cmbList.Items.Add(Strings.Left(lineText, Len(lineText) - 1))
                regKeys(increment) = GetHotkey(lineText, 0)
                regMacros(increment) = Strings.Left(lineText, Len(lineText) - 1)
                increment += 1
            End If
        Loop
        FileClose(1)
        For count = 0 To 11
            If Not regKeys.Contains(keyList(count)) Then
                cmbAddKey.Items.Add(keyList(count))
                cmbEditKey.Items.Add(keyList(count))
            End If
        Next
        cmbEditKey.Items.Insert(0, "")
    End Sub

    'Key listener and macro copier
    Private Sub tmrKeyListener_Tick(sender As Object, e As EventArgs) Handles tmrKeyListener.Tick
        Dim keyPressed As Boolean = KeyPressRange(Keys.F1, Keys.F12)
        Dim keyCode As Integer = GetKeyPressed(Keys.F1, Keys.F12)
        If keyPressed Then
            'Opens macro file for reading
            Dim keyName As String = [Enum].GetName(GetType(System.Windows.Forms.Keys), CType(keyCode, System.Windows.Forms.Keys))
            'MsgBox(keyName.ToString)
            If regKeys.Contains(keyName.ToString()) Then
                For counter = 0 To 11
                    'MsgBox(regKeys(counter) & " " & keyName.ToString)
                    'MsgBox(counter)
                    If regKeys(counter) = keyName.ToString() Then
                        My.Computer.Clipboard.SetText(regMacros(counter))
                        'Sends success notification
                        Notifier("Action Complete", "Macro successfully copied to clipboard")
                    End If
                Next
                FileClose(1)
            End If
            While GetAsyncKeyState(keyCode)
            End While
        End If
    End Sub

    'Required functions ran when program starts
    Private Sub FrmMacro_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        'Create the SMM appdata directory if it doesn't exist
        If Not Directory.Exists(MyAppData) Then
            Directory.CreateDirectory(MyAppData)
        End If
        'If macros file doesn't exist, create it in AppData Roaming
        If Not (File.Exists(Path.Combine(MyAppData, macroFile))) Then
            Dim fs As FileStream = File.Create(Path.Combine(MyAppData, macroFile))
            fs.Close()
        End If
        'Getting and setting the macro key and identifier for future use in the programs runtime.* All possible keys.
        For counter = 0 To 11
            keyList(counter) = "F" & (counter + 1)
        Next
        'Initialize macros
        InitMacros()
        'Initialize key listener
        tmrKeyListener.Start()
    End Sub

    'Handles notifications settings
    Private Sub chkNotify_CheckedChanged(sender As Object, e As EventArgs) Handles chkNotify.CheckedChanged
        If chkNotify.Checked Then
            notifications = True
        Else
            notifications = False
        End If
    End Sub

    'Switches macros in the macro editor
    Private Sub cmbList_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbList.SelectedIndexChanged
        cmbEditKey.Items.RemoveAt(0)
        cmbEditKey.Items.Insert(0, regKeys(Array.IndexOf(regMacros, cmbList.SelectedItem)))
        cmbEditKey.SelectedIndex = 0
        cmbEditKey.Enabled = True
        txtEdit.Text = cmbList.SelectedItem
        btnEdit.Enabled = True
        btnDel.Enabled = True
        lblEditKey.Enabled = True
        txtEdit.Enabled = True
    End Sub

    'Adds macros
    Private Sub btnAdd_Click(sender As Object, e As EventArgs) Handles btnAdd.Click
        If Not cmbAddKey.SelectedItem = Nothing Then
            Dim decision As Integer = Nothing
            If txtAdd.Text = "" Then
                MsgBox("Macro cannot be blank!", MessageBoxButtons.OK, "ERROR")
                Exit Sub
            ElseIf regMacros.Contains(txtAdd.Text) Then
                If decision = DialogResult.No Then
                    Exit Sub
                End If
            End If
            'Append the new macro
            FileOpen(1, filePath, OpenMode.Append)
            PrintLine(1, txtAdd.Text & GetHotkey(cmbAddKey.SelectedItem, 1))
            FileClose(1)
            MsgBox("Macro added successfully!", MessageBoxIcon.Information, "Success")
        End If
        'reinitialize macros
        InitMacros()
        txtAdd.Text = ""
    End Sub

    'Edits or deletes macros
    Private Sub btnEdit_Click(sender As Object, e As EventArgs) Handles btnEdit.Click, btnDel.Click
        If Not cmbList.SelectedItem = Nothing Then
            Dim fileText As String = ""
            Dim lineText As String = ""
            Dim searchDone As Boolean = False
            'If the user deletes the macro
            FileOpen(1, filePath, OpenMode.Input)
            Do Until EOF(1)
                Dim line As Integer = 0
                lineText = LineInput(1)
                If Not lineText = Nothing Then
                    If cmbList.SelectedIndex = Array.IndexOf(regMacros, Strings.Left(lineText, Len(lineText) - 1)) Then
                        Select Case DirectCast(sender, Button).Name
                            Case "btnEdit"
                                lineText = txtEdit.Text & GetHotkey(cmbEditKey.Text, 1)
                                MsgBox("Macro edited successfully!", MessageBoxIcon.Information, "Success")
                            Case "btnDel"
                                lineText = Nothing
                                MsgBox("Macro deleted successfully!", MessageBoxIcon.Information, "Success")
                        End Select
                        searchDone = True
                    End If
                End If
                If Not lineText = Nothing Then
                    If EOF(1) Then
                        fileText += lineText
                    Else
                        fileText += lineText & vbNewLine
                    End If
                End If
                line += 1
            Loop
            FileClose(1)

            'Export to file and reinitialize macros
            FileOpen(1, filePath, OpenMode.Output)
            PrintLine(1, fileText)
            FileClose(1)
            InitMacros()
        Else
            'Display error message
            MsgBox("Select a macro from the drop down list and try again.", MessageBoxIcon.Error, "Error")
        End If
        txtEdit.Text = ""
    End Sub

    Private Sub txtAdd_TextChanged(sender As Object, e As EventArgs) Handles txtAdd.TextChanged
        BtnToggle(btnAdd)
    End Sub

    Private Sub cmbAddKey_SelectedIndexChanged(sender As Object, e As EventArgs) Handles cmbAddKey.SelectedIndexChanged
        BtnToggle(btnAdd)
    End Sub

    Private Sub EditKeyToggle(sender As Object, e As EventArgs) Handles cmbEditKey.SelectedIndexChanged, txtEdit.TextChanged
        If txtEdit.Text = Nothing Then
            btnEdit.Enabled = False
        Else
            btnEdit.Enabled = True
        End If
    End Sub

    Private Sub tabSMM_SelectedIndexChanged(sender As Object, e As EventArgs) Handles tabSMM.SelectedIndexChanged
        InitMacros()
    End Sub
End Class