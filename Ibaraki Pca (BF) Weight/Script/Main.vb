Imports System.Console
Imports System.Text.Encoding
Imports System.Windows.Forms
Imports System.Windows.Forms.DialogResult
Imports System.Windows.Forms.MessageBox
Imports System.Windows.Forms.MessageBoxButtons

Public Module Main
    ''' <summary>
    ''' Main.
    ''' </summary>
    Public Sub Main()
        OutputEncoding = UTF8
        If Not My.Settings.Chk_Key Then
ChkPt:
            If InputBox("シリアルを入力", "ライセンスキー") = My.Resources.key_ser Then
                UpdVldLic()
                RunApp()
            Else
                If Show("ライセンスが間違っています！", "エラー", RetryCancel, MessageBoxIcon.Error) = Retry Then
                    GoTo ChkPt
                Else
                    ErrSty("終了するには、任意のキーを押してください...")
                    ReadKey()
                End If
            End If
        Else
            RunApp()
        End If
    End Sub
End Module
