' ビットマップからWS2812Bに送るバイト列を作成してシリアルポートから送出する。
' 先にSerialPort1を Baud: 3000000, Databits: 7, Parity: None, StopBits: Oneで開いておくこと。
' 16x16px, 1ピクセル送るのに8bytes 2048bytes バッファサイズはデフォルトで2048bytes
' 3M boudで 0.333us * 256 * 24 * 3 = 6.144ms かかる。
Private Sub Ws2812BSendBmp(ByRef sendBmp As Bitmap, Optional ByVal zigzag As Boolean = True)
    If SerialPort1.IsOpen = False Then
        MsgBox("Serialport is not open")
        Exit Sub
    End If

    '送信完了待ちとリセット
    'リセットは50us以上だけど50us作るのは大変そうなので1ms待つ。
    Do
        System.Threading.Thread.Sleep(1)
    Loop While SerialPort1.BytesToWrite > 1

    Try
        Dim colorAry() As Color = ws.ConvertBmpToColorArray(sendBmp, zigzag)
        Dim sendColorByteAry() As Byte = {}
        For Each clr In colorAry
            Dim oneColorByteAry() As Byte =
                ws.ColorToWs2812bByteArray(clr) '電流制限の都合で、dimmerは手前でやることにした。
            sendColorByteAry = ws.MargeByteAry(sendColorByteAry, oneColorByteAry)
        Next
        SerialPort1.Write(sendColorByteAry, 0, sendColorByteAry.Length)
    Catch ex As Exception
        MsgBox(ex.ToString)
    End Try

End Sub
