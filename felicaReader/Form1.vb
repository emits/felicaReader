Public Class Form1
    '======================
    ' felicalib.dll クラス
    '======================
    Private felicalib As New CFelicaLib

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        '------------------
        ' DLL 存在チェック
        '------------------
        If Not isDLLExists() Then
            MessageBox.Show("felicalib.dll がありません。", Me.Text)
            Application.Exit()
        End If

        '------------------
        ' ボタンの名前変更
        '------------------
        btnFelica.Text = "FeliCa 読み取り"

    End Sub

    '======================
    ' ボタン・クリック処理
    '======================
    Private Sub btnFelica_Click(sender As Object, e As EventArgs) Handles btnFelica.Click
        

        Dim sIDm As String
        Dim sPMm As String
        Dim sMsg As String

        '-------------------
        ' PaSoRi に接続する
        '-------------------
        If Not felicalib.Pasori_Connect() Then
            MessageBox.Show("PaSoRi に接続できませんでした。", Me.Text)
            Return
        End If

        '----------------------------
        ' ポーリング(FeliCa読み取り)
        '----------------------------
        If felicalib.Polling() Then
            '-----------------
            ' IDm, PMm を取得
            '-----------------
            sIDm = felicalib.getIDm()
            sPMm = felicalib.getPMm()

            '------------------
            ' メッセージを表示
            '------------------
            sMsg = "IDm=[" & sIDm & "]" & vbNewLine & _
                   "PMm=[" & sPMm & "]"
            MessageBox.Show(sMsg, Me.Text)
        Else
            '--------------------------
            ' ポーリングに失敗した場合
            '--------------------------
            MessageBox.Show("FeliCa がセットされていません。", Me.Text)
        End If

        '-------------------
        ' PaSoRi を解放する
        '-------------------
        felicalib.Pasori_Free()
    End Sub
End Class
