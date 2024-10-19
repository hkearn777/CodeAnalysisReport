<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()>
Partial Class Form1
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
    btnCodeAnalysisReport = New Button()
    btnClose = New Button()
    txtModelName = New TextBox()
    txtColumn1 = New TextBox()
    Label1 = New Label()
    txtWorkFolder = New TextBox()
    Label2 = New Label()
    Label3 = New Label()
    Label4 = New Label()
    txtContains1 = New TextBox()
    Label5 = New Label()
    Label6 = New Label()
    txtContains2 = New TextBox()
    Label8 = New Label()
    txtColumn2 = New TextBox()
    Label7 = New Label()
    txtTab = New TextBox()
    lblStatus = New Label()
    lbIndexes = New ListBox()
    Label9 = New Label()
    lblStatus2 = New Label()
    SuspendLayout()
    ' 
    ' btnCodeAnalysisReport
    ' 
    btnCodeAnalysisReport.Location = New Point(681, 391)
    btnCodeAnalysisReport.Name = "btnCodeAnalysisReport"
    btnCodeAnalysisReport.Size = New Size(208, 34)
    btnCodeAnalysisReport.TabIndex = 5
    btnCodeAnalysisReport.Text = "Code Analysis Report"
    btnCodeAnalysisReport.UseVisualStyleBackColor = True
    ' 
    ' btnClose
    ' 
    btnClose.Location = New Point(904, 391)
    btnClose.Name = "btnClose"
    btnClose.Size = New Size(112, 34)
    btnClose.TabIndex = 6
    btnClose.Text = "Close"
    btnClose.UseVisualStyleBackColor = True
    ' 
    ' txtModelName
    ' 
    txtModelName.Location = New Point(16, 126)
    txtModelName.Name = "txtModelName"
    txtModelName.Size = New Size(895, 31)
    txtModelName.TabIndex = 2
    txtModelName.Text = "C:\Users\906074897\Documents\All Projects\Lowes\Sandbox\ECATS\OUTPUT\eCATS Model.xlsx"
    ' 
    ' txtColumn1
    ' 
    txtColumn1.Location = New Point(103, 215)
    txtColumn1.Name = "txtColumn1"
    txtColumn1.Size = New Size(150, 31)
    txtColumn1.TabIndex = 3
    txtColumn1.Text = "Statement"
    ' 
    ' Label1
    ' 
    Label1.AutoSize = True
    Label1.Location = New Point(18, 14)
    Label1.Name = "Label1"
    Label1.Size = New Size(110, 25)
    Label1.TabIndex = 4
    Label1.Text = "Work folder:"
    ' 
    ' txtWorkFolder
    ' 
    txtWorkFolder.Location = New Point(16, 47)
    txtWorkFolder.Name = "txtWorkFolder"
    txtWorkFolder.Size = New Size(894, 31)
    txtWorkFolder.TabIndex = 1
    txtWorkFolder.Text = "C:\Users\906074897\Documents\All Projects\Lowes\Sandbox\ECATS\OUTPUT"
    ' 
    ' Label2
    ' 
    Label2.AutoSize = True
    Label2.Location = New Point(14, 92)
    Label2.Name = "Label2"
    Label2.Size = New Size(168, 25)
    Label2.TabIndex = 6
    Label2.Text = "Model spreadsheet:"
    ' 
    ' Label3
    ' 
    Label3.AutoSize = True
    Label3.Location = New Point(19, 215)
    Label3.Name = "Label3"
    Label3.Size = New Size(78, 25)
    Label3.TabIndex = 7
    Label3.Text = "Column:"
    ' 
    ' Label4
    ' 
    Label4.AutoSize = True
    Label4.Location = New Point(571, 221)
    Label4.Name = "Label4"
    Label4.Size = New Size(50, 25)
    Label4.TabIndex = 8
    Label4.Text = "AND"
    ' 
    ' txtContains1
    ' 
    txtContains1.Location = New Point(397, 218)
    txtContains1.Name = "txtContains1"
    txtContains1.Size = New Size(150, 31)
    txtContains1.TabIndex = 4
    txtContains1.Text = "CST"
    ' 
    ' Label5
    ' 
    Label5.AutoSize = True
    Label5.Location = New Point(259, 221)
    Label5.Name = "Label5"
    Label5.Size = New Size(132, 25)
    Label5.TabIndex = 10
    Label5.Text = "Contains Value:"
    ' 
    ' Label6
    ' 
    Label6.AutoSize = True
    Label6.Location = New Point(259, 271)
    Label6.Name = "Label6"
    Label6.Size = New Size(132, 25)
    Label6.TabIndex = 15
    Label6.Text = "Contains Value:"
    ' 
    ' txtContains2
    ' 
    txtContains2.Location = New Point(397, 268)
    txtContains2.Name = "txtContains2"
    txtContains2.Size = New Size(150, 31)
    txtContains2.TabIndex = 12
    txtContains2.Text = "T332"
    ' 
    ' Label8
    ' 
    Label8.AutoSize = True
    Label8.Location = New Point(19, 265)
    Label8.Name = "Label8"
    Label8.Size = New Size(78, 25)
    Label8.TabIndex = 13
    Label8.Text = "Column:"
    ' 
    ' txtColumn2
    ' 
    txtColumn2.Location = New Point(103, 265)
    txtColumn2.Name = "txtColumn2"
    txtColumn2.Size = New Size(150, 31)
    txtColumn2.TabIndex = 11
    txtColumn2.Text = "Table"
    ' 
    ' Label7
    ' 
    Label7.AutoSize = True
    Label7.Location = New Point(51, 178)
    Label7.Name = "Label7"
    Label7.Size = New Size(46, 25)
    Label7.TabIndex = 16
    Label7.Text = "TAB:"
    ' 
    ' txtTab
    ' 
    txtTab.Location = New Point(106, 177)
    txtTab.Name = "txtTab"
    txtTab.Size = New Size(147, 31)
    txtTab.TabIndex = 17
    txtTab.Text = "ExecSQL"
    ' 
    ' lblStatus
    ' 
    lblStatus.AutoSize = True
    lblStatus.Location = New Point(21, 308)
    lblStatus.Name = "lblStatus"
    lblStatus.Size = New Size(60, 25)
    lblStatus.TabIndex = 18
    lblStatus.Text = "Status"
    ' 
    ' lbIndexes
    ' 
    lbIndexes.FormattingEnabled = True
    lbIndexes.ItemHeight = 25
    lbIndexes.Location = New Point(681, 204)
    lbIndexes.Name = "lbIndexes"
    lbIndexes.Size = New Size(208, 129)
    lbIndexes.TabIndex = 21
    ' 
    ' Label9
    ' 
    Label9.AutoSize = True
    Label9.Location = New Point(680, 170)
    Label9.Name = "Label9"
    Label9.Size = New Size(76, 25)
    Label9.TabIndex = 22
    Label9.Text = "Indexes:"
    ' 
    ' lblStatus2
    ' 
    lblStatus2.AutoSize = True
    lblStatus2.Location = New Point(27, 344)
    lblStatus2.Name = "lblStatus2"
    lblStatus2.Size = New Size(69, 25)
    lblStatus2.TabIndex = 23
    lblStatus2.Text = "status2"
    ' 
    ' Form1
    ' 
    AutoScaleDimensions = New SizeF(10F, 25F)
    AutoScaleMode = AutoScaleMode.Font
    ClientSize = New Size(1039, 450)
    Controls.Add(lblStatus2)
    Controls.Add(Label9)
    Controls.Add(lbIndexes)
    Controls.Add(lblStatus)
    Controls.Add(txtTab)
    Controls.Add(Label7)
    Controls.Add(Label6)
    Controls.Add(txtContains2)
    Controls.Add(Label8)
    Controls.Add(txtColumn2)
    Controls.Add(Label5)
    Controls.Add(txtContains1)
    Controls.Add(Label4)
    Controls.Add(Label3)
    Controls.Add(Label2)
    Controls.Add(txtWorkFolder)
    Controls.Add(Label1)
    Controls.Add(txtColumn1)
    Controls.Add(txtModelName)
    Controls.Add(btnClose)
    Controls.Add(btnCodeAnalysisReport)
    Name = "Form1"
    Text = "Code Analysis Report"
    ResumeLayout(False)
    PerformLayout()
  End Sub

  Friend WithEvents btnCodeAnalysisReport As Button
  Friend WithEvents btnClose As Button
  Friend WithEvents txtModelName As TextBox
  Friend WithEvents txtColumn1 As TextBox
  Friend WithEvents Label1 As Label
  Friend WithEvents txtWorkFolder As TextBox
  Friend WithEvents Label2 As Label
  Friend WithEvents Label3 As Label
  Friend WithEvents Label4 As Label
  Friend WithEvents txtContains1 As TextBox
  Friend WithEvents Label5 As Label
  Friend WithEvents Label6 As Label
  Friend WithEvents txtContains2 As TextBox
  Friend WithEvents Label8 As Label
  Friend WithEvents txtColumn2 As TextBox
  Friend WithEvents Label7 As Label
  Friend WithEvents txtTab As TextBox
  Friend WithEvents lblStatus As Label
  Friend WithEvents lbIndexes As ListBox
  Friend WithEvents Label9 As Label
  Friend WithEvents lblStatus2 As Label

End Class
