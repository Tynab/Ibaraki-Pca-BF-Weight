Imports System.Diagnostics.Process
Imports System.Math
Imports System.Net
Imports System.Windows.Forms
Imports System.Windows.Forms.Keys

Public Class FrmUpdate
#Region "Fields"
    Private ReadOnly _wc As New WebClient
#End Region

#Region "Overridden"
    ' Hide sub window
    Protected Overrides ReadOnly Property CreateParams() As CreateParams
        Get
            Dim cp As CreateParams = MyBase.CreateParams
            cp.ExStyle = cp.ExStyle Or &H80
            Return cp
        End Get
    End Property

    ' Deny Alt+F4
    Protected Overrides Function ProcessCmdKey(ByRef msg As Message, keyData As Keys) As Boolean
        Return keyData = (Alt Or F4) OrElse MyBase.ProcessCmdKey(msg, keyData)
    End Function
#End Region

#Region "Events"
    ' Load frm
    Private Sub FrmUpdate_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        lblCapacity.Text = ""
        lblPercent.Text = ""
        pnlProgressBar.Width = 1
    End Sub

    ' Shown frm
    Private Sub FrmUpdate_Shown(sender As Object, e As EventArgs) Handles MyBase.Shown
        FIFrm()
        CrtDirAdv(FRNT_PATH)
        DelFileAdv(FILE_SETUP_ADR)
        tmrMain.StrtAdv()
        AddHandler _wc.DownloadProgressChanged, AddressOf Upd_DownloadProgressChanged
        _wc.DownloadFileAsync(New Uri(_wc.DownloadString(My.Resources.link_app)), FILE_SETUP_ADR)
    End Sub

    ' Update download progress
    Private Sub Upd_DownloadProgressChanged(sender As Object, e As DownloadProgressChangedEventArgs)
        lblCapacity.Text = String.Format("{0} MB / {1} MB", (e.BytesReceived / 1024D / 1024D).ToString("0.00"), (e.TotalBytesToReceive / 1024D / 1024D).ToString("0.00"))
        lblPercent.Text = $"{e.ProgressPercentage}%"
        pnlProgressBar.Width = CInt(Ceiling(e.ProgressPercentage * Width / 100D))
    End Sub

    ' tmr main
    Private Sub TmrMain_Tick(sender As Object, e As EventArgs) Handles tmrMain.Tick
        If lblPercent.Text = "100%" Then
            tmrMain.StopAdv()
            Close()
        End If
    End Sub

    ' Closing frm
    Private Sub FrmUpdate_FormClosing(sender As Object, e As FormClosingEventArgs) Handles MyBase.FormClosing
        FOFrm()
        KillPrcs(My.Resources.app_name)
    End Sub

    ' Closed frm
    Private Sub FrmUpdate_FormClosed(sender As Object, e As FormClosedEventArgs) Handles MyBase.FormClosed
        Start(FILE_SETUP_ADR)
    End Sub
#End Region
End Class