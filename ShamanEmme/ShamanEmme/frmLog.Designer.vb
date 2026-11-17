<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class frmLog
    Inherits System.Windows.Forms.Form

    'Form overrides dispose to clean up the component list.
    <System.Diagnostics.DebuggerNonUserCode()> _
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
    <System.Diagnostics.DebuggerStepThrough()> _
    Private Sub InitializeComponent()
        Me.components = New System.ComponentModel.Container()
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(frmLog))
        Me.tmrInit = New System.Windows.Forms.Timer(Me.components)
        Me.pgbProceso = New System.Windows.Forms.ProgressBar()
        Me.txtLog = New System.Windows.Forms.TextBox()
        Me.SuspendLayout()
        '
        'tmrInit
        '
        Me.tmrInit.Interval = 3000
        '
        'pgbProceso
        '
        Me.pgbProceso.Dock = System.Windows.Forms.DockStyle.Bottom
        Me.pgbProceso.Location = New System.Drawing.Point(0, 272)
        Me.pgbProceso.Name = "pgbProceso"
        Me.pgbProceso.Size = New System.Drawing.Size(683, 23)
        Me.pgbProceso.TabIndex = 1
        '
        'txtLog
        '
        Me.txtLog.Dock = System.Windows.Forms.DockStyle.Fill
        Me.txtLog.Location = New System.Drawing.Point(0, 0)
        Me.txtLog.Multiline = True
        Me.txtLog.Name = "txtLog"
        Me.txtLog.ReadOnly = True
        Me.txtLog.ScrollBars = System.Windows.Forms.ScrollBars.Vertical
        Me.txtLog.Size = New System.Drawing.Size(683, 272)
        Me.txtLog.TabIndex = 2
        '
        'frmLog
        '
        Me.AutoScaleDimensions = New System.Drawing.SizeF(6.0!, 13.0!)
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.Font
        Me.ClientSize = New System.Drawing.Size(683, 295)
        Me.Controls.Add(Me.txtLog)
        Me.Controls.Add(Me.pgbProceso)
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.Name = "frmLog"
        Me.Text = "Log de Actividades"
        Me.ResumeLayout(False)
        Me.PerformLayout()

    End Sub
    Friend WithEvents tmrInit As System.Windows.Forms.Timer
    Friend WithEvents pgbProceso As System.Windows.Forms.ProgressBar
    Friend WithEvents txtLog As System.Windows.Forms.TextBox

End Class
