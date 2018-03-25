<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class frmMacro
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()>
    Protected Overrides Sub Dispose(ByVal disposing As Boolean)
        Try
            If disposing AndAlso components IsNot Nothing Then
                components.Dispose()
            End If
        Finally
            MyBase.Dispose(disposing)
        End Try
    End Sub

    'Required by the Windows Form Designer
    Private components As System.ComponentModel.IContainer

    'NOTE: The following procedure is required by the Windows Form Designer
    'It can be modified using the Windows Form Designer.  
    'Do not modify it using the code editor.
    <System.Diagnostics.DebuggerStepThrough()>
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmMacro))
        Me.Panel1 = New System.Windows.Forms.Panel()
        Me.tabSMM = New System.Windows.Forms.TabControl()
        Me.TabPage1 = New System.Windows.Forms.TabPage()
        Me.txtAdd = New System.Windows.Forms.TextBox()
        Me.btnAdd = New System.Windows.Forms.Button()
        Me.lblAddKey = New System.Windows.Forms.Label()
        Me.lblMacroAdd = New System.Windows.Forms.Label()
        Me.cmbAddKey = New System.Windows.Forms.ComboBox()
        Me.TabPage2 = New System.Windows.Forms.TabPage()
        Me.cmbEditKey = New System.Windows.Forms.ComboBox()
        Me.cmbList = New System.Windows.Forms.ComboBox()
        Me.btnDel = New System.Windows.Forms.Button()
        Me.btnEdit = New System.Windows.Forms.Button()
        Me.txtEdit = New System.Windows.Forms.TextBox()
        Me.lblEditKey = New System.Windows.Forms.Label()
        Me.lblList = New System.Windows.Forms.Label()
        Me.lblContents = New System.Windows.Forms.Label()
        Me.TabPage3 = New System.Windows.Forms.TabPage()
        Me.chkNotify = New System.Windows.Forms.CheckBox()
        Me.tmrKeyListener = New System.Windows.Forms.Timer(Me.components)
        Me.NotifyIcon = New System.Windows.Forms.NotifyIcon(Me.components)
        Me.Panel1.SuspendLayout()
        Me.tabSMM.SuspendLayout()
        Me.TabPage1.SuspendLayout()
        Me.TabPage2.SuspendLayout()
        Me.TabPage3.SuspendLayout()
        Me.SuspendLayout()
        '
        'Panel1
        '
        resources.ApplyResources(Me.Panel1, "Panel1")
        Me.Panel1.Controls.Add(Me.tabSMM)
        Me.Panel1.Name = "Panel1"
        '
        'tabSMM
        '
        resources.ApplyResources(Me.tabSMM, "tabSMM")
        Me.tabSMM.Controls.Add(Me.TabPage1)
        Me.tabSMM.Controls.Add(Me.TabPage2)
        Me.tabSMM.Controls.Add(Me.TabPage3)
        Me.tabSMM.Name = "tabSMM"
        Me.tabSMM.SelectedIndex = 0
        Me.tabSMM.SizeMode = System.Windows.Forms.TabSizeMode.FillToRight
        Me.tabSMM.TabStop = False
        '
        'TabPage1
        '
        Me.TabPage1.Controls.Add(Me.txtAdd)
        Me.TabPage1.Controls.Add(Me.btnAdd)
        Me.TabPage1.Controls.Add(Me.lblAddKey)
        Me.TabPage1.Controls.Add(Me.lblMacroAdd)
        Me.TabPage1.Controls.Add(Me.cmbAddKey)
        resources.ApplyResources(Me.TabPage1, "TabPage1")
        Me.TabPage1.Name = "TabPage1"
        Me.TabPage1.UseVisualStyleBackColor = True
        '
        'txtAdd
        '
        resources.ApplyResources(Me.txtAdd, "txtAdd")
        Me.txtAdd.Name = "txtAdd"
        '
        'btnAdd
        '
        resources.ApplyResources(Me.btnAdd, "btnAdd")
        Me.btnAdd.Name = "btnAdd"
        Me.btnAdd.UseVisualStyleBackColor = True
        '
        'lblAddKey
        '
        resources.ApplyResources(Me.lblAddKey, "lblAddKey")
        Me.lblAddKey.Name = "lblAddKey"
        '
        'lblMacroAdd
        '
        resources.ApplyResources(Me.lblMacroAdd, "lblMacroAdd")
        Me.lblMacroAdd.Name = "lblMacroAdd"
        '
        'cmbAddKey
        '
        resources.ApplyResources(Me.cmbAddKey, "cmbAddKey")
        Me.cmbAddKey.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbAddKey.FormattingEnabled = True
        Me.cmbAddKey.Name = "cmbAddKey"
        Me.cmbAddKey.Sorted = True
        '
        'TabPage2
        '
        Me.TabPage2.Controls.Add(Me.cmbEditKey)
        Me.TabPage2.Controls.Add(Me.cmbList)
        Me.TabPage2.Controls.Add(Me.btnDel)
        Me.TabPage2.Controls.Add(Me.btnEdit)
        Me.TabPage2.Controls.Add(Me.txtEdit)
        Me.TabPage2.Controls.Add(Me.lblEditKey)
        Me.TabPage2.Controls.Add(Me.lblList)
        Me.TabPage2.Controls.Add(Me.lblContents)
        resources.ApplyResources(Me.TabPage2, "TabPage2")
        Me.TabPage2.Name = "TabPage2"
        Me.TabPage2.UseVisualStyleBackColor = True
        '
        'cmbEditKey
        '
        resources.ApplyResources(Me.cmbEditKey, "cmbEditKey")
        Me.cmbEditKey.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbEditKey.FormattingEnabled = True
        Me.cmbEditKey.Name = "cmbEditKey"
        '
        'cmbList
        '
        Me.cmbList.DropDownStyle = System.Windows.Forms.ComboBoxStyle.DropDownList
        Me.cmbList.FormattingEnabled = True
        resources.ApplyResources(Me.cmbList, "cmbList")
        Me.cmbList.Name = "cmbList"
        '
        'btnDel
        '
        resources.ApplyResources(Me.btnDel, "btnDel")
        Me.btnDel.Name = "btnDel"
        Me.btnDel.UseVisualStyleBackColor = True
        '
        'btnEdit
        '
        resources.ApplyResources(Me.btnEdit, "btnEdit")
        Me.btnEdit.Name = "btnEdit"
        Me.btnEdit.UseVisualStyleBackColor = True
        '
        'txtEdit
        '
        resources.ApplyResources(Me.txtEdit, "txtEdit")
        Me.txtEdit.Name = "txtEdit"
        '
        'lblEditKey
        '
        resources.ApplyResources(Me.lblEditKey, "lblEditKey")
        Me.lblEditKey.Name = "lblEditKey"
        '
        'lblList
        '
        resources.ApplyResources(Me.lblList, "lblList")
        Me.lblList.Name = "lblList"
        '
        'lblContents
        '
        resources.ApplyResources(Me.lblContents, "lblContents")
        Me.lblContents.Name = "lblContents"
        '
        'TabPage3
        '
        Me.TabPage3.Controls.Add(Me.chkNotify)
        resources.ApplyResources(Me.TabPage3, "TabPage3")
        Me.TabPage3.Name = "TabPage3"
        Me.TabPage3.UseVisualStyleBackColor = True
        '
        'chkNotify
        '
        resources.ApplyResources(Me.chkNotify, "chkNotify")
        Me.chkNotify.Checked = True
        Me.chkNotify.CheckState = System.Windows.Forms.CheckState.Checked
        Me.chkNotify.Name = "chkNotify"
        Me.chkNotify.UseVisualStyleBackColor = True
        '
        'tmrKeyListener
        '
        Me.tmrKeyListener.Interval = 20
        '
        'NotifyIcon
        '
        resources.ApplyResources(Me.NotifyIcon, "NotifyIcon")
        '
        'frmMacro
        '
        resources.ApplyResources(Me, "$this")
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.Controls.Add(Me.Panel1)
        Me.Name = "frmMacro"
        Me.Panel1.ResumeLayout(False)
        Me.tabSMM.ResumeLayout(False)
        Me.TabPage1.ResumeLayout(False)
        Me.TabPage1.PerformLayout()
        Me.TabPage2.ResumeLayout(False)
        Me.TabPage2.PerformLayout()
        Me.TabPage3.ResumeLayout(False)
        Me.TabPage3.PerformLayout()
        Me.ResumeLayout(False)

    End Sub
    Friend WithEvents Panel1 As Panel
    Friend WithEvents tmrKeyListener As Timer
    Friend WithEvents tabSMM As TabControl
    Friend WithEvents TabPage1 As TabPage
    Friend WithEvents txtAdd As TextBox
    Friend WithEvents btnAdd As Button
    Friend WithEvents lblAddKey As Label
    Friend WithEvents lblMacroAdd As Label
    Friend WithEvents cmbAddKey As ComboBox
    Friend WithEvents TabPage2 As TabPage
    Friend WithEvents cmbEditKey As ComboBox
    Friend WithEvents cmbList As ComboBox
    Friend WithEvents btnDel As Button
    Friend WithEvents btnEdit As Button
    Friend WithEvents txtEdit As TextBox
    Friend WithEvents lblEditKey As Label
    Friend WithEvents lblList As Label
    Friend WithEvents lblContents As Label
    Friend WithEvents TabPage3 As TabPage
    Friend WithEvents chkNotify As CheckBox
    Friend WithEvents NotifyIcon As NotifyIcon
End Class
