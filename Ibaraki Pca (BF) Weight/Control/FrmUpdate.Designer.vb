<Global.Microsoft.VisualBasic.CompilerServices.DesignerGenerated()> _
Partial Class FrmUpdate
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
        Dim resources As System.ComponentModel.ComponentResourceManager = New System.ComponentModel.ComponentResourceManager(GetType(FrmUpdate))
        Me.tmrMain = New System.Windows.Forms.Timer(Me.components)
        Me.pnlMain = New System.Windows.Forms.Panel()
        Me.pnlProgressBar = New System.Windows.Forms.Panel()
        Me.lblPercent = New System.Windows.Forms.Label()
        Me.lblText = New System.Windows.Forms.Label()
        Me.picUpdating = New System.Windows.Forms.PictureBox()
        Me.lblCapacity = New System.Windows.Forms.Label()
        Me.pnlMain.SuspendLayout()
        CType(Me.picUpdating, System.ComponentModel.ISupportInitialize).BeginInit()
        Me.SuspendLayout()
        '
        'pnlMain
        '
        Me.pnlMain.BackColor = System.Drawing.Color.White
        Me.pnlMain.Controls.Add(Me.pnlProgressBar)
        Me.pnlMain.Controls.Add(Me.lblPercent)
        Me.pnlMain.Controls.Add(Me.lblText)
        Me.pnlMain.Controls.Add(Me.picUpdating)
        Me.pnlMain.Controls.Add(Me.lblCapacity)
        Me.pnlMain.Dock = System.Windows.Forms.DockStyle.Fill
        Me.pnlMain.Font = New System.Drawing.Font("Segoe UI Light", 9.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.pnlMain.Location = New System.Drawing.Point(0, 0)
        Me.pnlMain.Margin = New System.Windows.Forms.Padding(0)
        Me.pnlMain.Name = "pnlMain"
        Me.pnlMain.Size = New System.Drawing.Size(360, 240)
        Me.pnlMain.TabIndex = 0
        '
        'pnlProgressBar
        '
        Me.pnlProgressBar.BackColor = System.Drawing.Color.LightSkyBlue
        Me.pnlProgressBar.Dock = System.Windows.Forms.DockStyle.Left
        Me.pnlProgressBar.Location = New System.Drawing.Point(0, 235)
        Me.pnlProgressBar.Margin = New System.Windows.Forms.Padding(0)
        Me.pnlProgressBar.Name = "pnlProgressBar"
        Me.pnlProgressBar.Size = New System.Drawing.Size(360, 5)
        Me.pnlProgressBar.TabIndex = 0
        '
        'lblPercent
        '
        Me.lblPercent.BackColor = System.Drawing.Color.Transparent
        Me.lblPercent.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblPercent.Font = New System.Drawing.Font("Segoe UI Light", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblPercent.ForeColor = System.Drawing.Color.Gray
        Me.lblPercent.Location = New System.Drawing.Point(0, 215)
        Me.lblPercent.Margin = New System.Windows.Forms.Padding(0)
        Me.lblPercent.Name = "lblPercent"
        Me.lblPercent.Size = New System.Drawing.Size(360, 20)
        Me.lblPercent.TabIndex = 0
        Me.lblPercent.Text = "0%"
        Me.lblPercent.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'lblText
        '
        Me.lblText.BackColor = System.Drawing.Color.Transparent
        Me.lblText.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblText.Font = New System.Drawing.Font("Yu Gothic", 13.0!, System.Drawing.FontStyle.Regular, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblText.ForeColor = System.Drawing.Color.Gray
        Me.lblText.Location = New System.Drawing.Point(0, 170)
        Me.lblText.Margin = New System.Windows.Forms.Padding(0)
        Me.lblText.Name = "lblText"
        Me.lblText.Size = New System.Drawing.Size(360, 45)
        Me.lblText.TabIndex = 0
        Me.lblText.Text = "更新中..."
        Me.lblText.TextAlign = System.Drawing.ContentAlignment.TopCenter
        '
        'picUpdating
        '
        Me.picUpdating.BackColor = System.Drawing.Color.Transparent
        Me.picUpdating.Dock = System.Windows.Forms.DockStyle.Top
        Me.picUpdating.Image = Global.Ibaraki_Pca__BF__Weight.My.Resources.Resources.gUpdating
        Me.picUpdating.Location = New System.Drawing.Point(0, 20)
        Me.picUpdating.Margin = New System.Windows.Forms.Padding(0)
        Me.picUpdating.Name = "picUpdating"
        Me.picUpdating.Size = New System.Drawing.Size(360, 150)
        Me.picUpdating.SizeMode = System.Windows.Forms.PictureBoxSizeMode.Zoom
        Me.picUpdating.TabIndex = 1
        Me.picUpdating.TabStop = False
        '
        'lblCapacity
        '
        Me.lblCapacity.BackColor = System.Drawing.Color.Transparent
        Me.lblCapacity.Dock = System.Windows.Forms.DockStyle.Top
        Me.lblCapacity.Font = New System.Drawing.Font("Segoe UI Light", 10.0!, System.Drawing.FontStyle.Bold, System.Drawing.GraphicsUnit.Point, CType(0, Byte))
        Me.lblCapacity.ForeColor = System.Drawing.Color.Gray
        Me.lblCapacity.Location = New System.Drawing.Point(0, 0)
        Me.lblCapacity.Margin = New System.Windows.Forms.Padding(0)
        Me.lblCapacity.Name = "lblCapacity"
        Me.lblCapacity.Size = New System.Drawing.Size(360, 20)
        Me.lblCapacity.TabIndex = 0
        Me.lblCapacity.Text = "0 MB / 0 MB"
        Me.lblCapacity.TextAlign = System.Drawing.ContentAlignment.BottomCenter
        '
        'FrmUpdate
        '
        Me.AutoScaleMode = System.Windows.Forms.AutoScaleMode.None
        Me.BackColor = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.ClientSize = New System.Drawing.Size(360, 240)
        Me.ControlBox = False
        Me.Controls.Add(Me.pnlMain)
        Me.DoubleBuffered = True
        Me.FormBorderStyle = System.Windows.Forms.FormBorderStyle.None
        Me.Icon = CType(resources.GetObject("$this.Icon"), System.Drawing.Icon)
        Me.MaximizeBox = False
        Me.MinimizeBox = False
        Me.Name = "FrmUpdate"
        Me.Opacity = 0R
        Me.ShowInTaskbar = False
        Me.StartPosition = System.Windows.Forms.FormStartPosition.Manual
        Me.Text = "更新中..."
        Me.TopMost = True
        Me.TransparencyKey = System.Drawing.Color.FromArgb(CType(CType(64, Byte), Integer), CType(CType(0, Byte), Integer), CType(CType(64, Byte), Integer))
        Me.pnlMain.ResumeLayout(False)
        CType(Me.picUpdating, System.ComponentModel.ISupportInitialize).EndInit()
        Me.ResumeLayout(False)

    End Sub

    Friend WithEvents tmrMain As Windows.Forms.Timer
    Private WithEvents pnlMain As Windows.Forms.Panel
    Public WithEvents pnlProgressBar As Windows.Forms.Panel
    Public WithEvents lblPercent As Windows.Forms.Label
    Private WithEvents lblText As Windows.Forms.Label
    Private WithEvents picUpdating As Windows.Forms.PictureBox
    Public WithEvents lblCapacity As Windows.Forms.Label
End Class
