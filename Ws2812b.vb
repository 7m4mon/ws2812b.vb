''' <summary>
''' WS2812Bに送るバイト配列を作るためのクラス。
''' </summary>


Public Class Ws2812b

    Private Const MAX_CURRENT As Double = 0.0518   '１個のWS2812BでWhiteの電流値（測定結果）
    Private Const ONE_DIGIT_CURRENT As Double = MAX_CURRENT / (255 * 3) ' 1digitあたりの電流値


    '色を受け取ってWS2812B用のバイト配列を作る。
    'シリアルポートは3M boud 7-n-1にする
    'これで、スタートビットとストップビットあわせて9ビットとなる。
    '通常、待機時H、スタートビットはL、ストップビットはHである。
    'よってWS2812Bに送る時は反転する必要がある。
    '普通のUARTは1がHigh,0がLow。つまり、0のときマークする。
    'ONのときは 0,0,1 と送る。 OFFの時は 0,1,1と送る。
    '3MBaudなので、1ビットあたり0.333usとなる。
    'データレート：最大３Ｍｂａｕｄ [AE-FT234X]
    '上記製品は5Vトレラントだけど入力側のしきい値が5V*0.7=3.5Vなので、
    'ダイオードで0.6V電源電圧降下させたTC7S04Fを挟んだ。（Hレベルがギリなので。）
    'デジトラ反転＆レベル変換は220Ωのプルアップ抵抗でも全然だめだった。
    '1バイト送信で3ビット送れるので、24ビットカラーを1つ送るには8バイト必要。
    'SerialPortはLSBファーストらしいので組立順序はLSBから。
    '眩しいのでdimmerLevelをもって暗い方向にシフト出来るように。dimmerLevelは8まで。

    ''' <summary>
    ''' 色を受け取ってWS2812B用のバイト配列を作る。
    ''' </summary>
    ''' <param name="inputColor">色</param>
    ''' <param name="dimmerLevel">ディマーの値(0~8で0が明るい)</param>
    ''' <returns>シリアルポートに送るバイト配列。シリアルポートは3M boud 7-n-1にする。</returns>
    Function ColorToWs2812bByteArray(ByVal inputColor As Color,
                                     Optional ByVal dimmerLevel As Integer = 0) As Byte()
        Dim outAry(7) As Byte       ' 1色送るのに8バイト必要。
        '眩しいのでdimmerLevelをもってシフトする。dimmerLevelは8まで。
        Dim dimmerG As UInteger = CUInt(inputColor.G) >> dimmerLevel
        Dim dimmerR As UInteger = CUInt(inputColor.R) >> dimmerLevel
        Dim dimmerB As UInteger = CUInt(inputColor.B) >> dimmerLevel
        Dim grb24 As Integer = (dimmerG << 16) + (dimmerR << 8) + dimmerB
        If inputColor.A < 128 Then      '透明だったら黒にする。
            grb24 = 0
        End If

        For i = 0 To 7
            Dim tByte As Byte = 0
            If (grb24 And &H800000) <> 0 Then      '24ビット目
                tByte = &H2                     '000 000 10 (7-n-1で、1ビット目はスタートビット0）
            Else
                tByte = &H3                     '000 000 11 (7-n-1で、1ビット目はスタートビット0)
            End If
            If (grb24 And &H400000) <> 0 Then      '23ビット目
                tByte += &H10                   '000 100 00 、LSBファースト
            Else
                tByte += &H18                   '000 110 00 、LSBファースト
            End If
            If (grb24 And &H200000) <> 0 Then      '22ビット目
                tByte += &H0                    '000 000 00 (最後にストップビット1が付く)、LSBファースト
            Else
                tByte += &H40                    '010 000 00 (最後にストップビット1が付く)、LSBファースト
            End If
            outAry(i) = tByte
            grb24 <<= 3                         '3ビット左シフトして次の3ビットをMSBに移動。
        Next
        Return outAry
    End Function


    ''' <summary>
    ''' ビットマップからカラーの配列に変換する。走査方向は横から。
    ''' </summary>
    ''' <param name="inputBmp">入力画像。16x16とか。</param>
    ''' <param name="zigzag">WS2812Bの実装がジグザグになっていたらtrue</param>
    ''' <returns></returns>
    Public Function ConvertBmpToColorArray(ByRef inputBmp As Bitmap, Optional zigzag As Boolean = False) As Color()
        Dim w As Integer = inputBmp.Width
        Dim h As Integer = inputBmp.Height
        Dim colorAry(w * h - 1) As Color

        Dim n As Integer = 0    '個数（配列の位置）
        For y = 0 To h - 1
            Dim startX, stopX, stepX As Integer         'LEDの接続がジグザグになっている。
            Dim odd As Boolean = If((y Mod 2) = 0, False, True)  '奇数列だけ戻る方向
            startX = If(zigzag And odd, w - 1, 0)                'ZIGZAGでかつ奇数列のとき
            stopX = If(zigzag And odd, 0, w - 1)
            stepX = If(zigzag And odd, -1, 1)

            For x = startX To stopX Step stepX
                colorAry(n) = inputBmp.GetPixel(x, y)
                n += 1          '個数（配列の位置）のインクリメント
            Next
        Next
        'Debug.Print(n)
        Return colorAry
    End Function


    ''' <summary>
    ''' バイト配列を結合するファンクション。
    ''' </summary>
    ''' <param name="ary1">バイト配列１</param>
    ''' <param name="ary2">バイト配列２</param>
    ''' <returns>バイト配列１＋バイト配列２</returns>
    Public Function MargeByteAry(ByVal ary1() As Byte, ByVal ary2() As Byte) As Byte()
        Dim mergedArray As Byte() = New Byte(ary1.Length + ary2.Length - 1) {}
        ary1.CopyTo(mergedArray, 0)
        ary2.CopyTo(mergedArray, ary1.Length)
        Return mergedArray
    End Function

    ''' <summary>
    ''' カラー配列を結合するファンクション。
    ''' </summary>
    ''' <param name="ary1">カラー配列１</param>
    ''' <param name="ary2">カラー配列２</param>
    ''' <returns>カラー配列１＋カラー配列２</returns>
    Public Function MargeColorAry(ByVal ary1() As Color, ByVal ary2() As Color) As Color()
        Dim mergedArray As Color() = New Color(ary1.Length + ary2.Length - 1) {}
        ary1.CopyTo(mergedArray, 0)
        ary2.CopyTo(mergedArray, ary1.Length)
        Return mergedArray
    End Function


    ''' <summary>
    ''' WS2812Bの輝度を、最大電流を指定して落とす。maxAmpareは4.8とかを想定。
    ''' </summary>
    ''' <param name="bmp">制限したいビットマップ画像</param>
    ''' <param name="maxAmpare">電流制限値(A)</param>
    Public Sub WS2812_CurrentLimit(ByRef bmp As Bitmap, ByVal maxAmpare As Double)
        Dim oneDigitAmpare As Double = ONE_DIGIT_CURRENT     '最大輝度で51.8mA程度になるらしい。67.7uA程度。
        Dim maxBrightness As UInteger = Math.Floor(maxAmpare / oneDigitAmpare)  '4.8Aの電源の場合、73440
        SatulateBmp(bmp, maxBrightness)

    End Sub

    ''' <summary>
    ''' およその電流値を返す。
    ''' </summary>
    ''' <param name="bmp">電流値を調べたいビットマップ画像</param>
    ''' <returns>電流値(A)</returns>
    Public Function WS2812_CurrentAmount(ByRef bmp As Bitmap) As Double
        Dim clrAry() As Color = ConvertBmpToColorArray(bmp)
        Dim brightness As UInteger = ConvertColorArrayToBrightness(clrAry)
        Dim totalCurrent As Double = 0
        totalCurrent = brightness * ONE_DIGIT_CURRENT
        Return totalCurrent
    End Function

    ''' <summary>
    ''' WS2812Bのだいたいの電流の目安を返す
    ''' </summary>
    ''' <param name="dimmerLevel">ディマーの値(0~8, 0が明るい)</param>
    ''' <param name="numOfUse">LEDの個数</param>
    ''' <returns>目安の電流値(A)</returns>
    Function WS2812_CurrentQuota(ByVal dimmerLevel As Integer, ByVal numOfUse As Integer) As Double
        Dim totalCurrent As Double = 0
        totalCurrent = 2 ^ (-dimmerLevel) * MAX_CURRENT * numOfUse       '
        Return totalCurrent
    End Function

    '色からRGB合計の明るさを返す関数
    Private Function ColorBrightness(ByVal clr As Color) As UInteger
        Dim brightness As UInteger = 0
        brightness += clr.R
        brightness += clr.G
        brightness += clr.B

        Return brightness

    End Function

    'カラー配列からすべての明るさを足して返す。
    '+4,294,967,295まで。
    '256画素の場合、256*(255*3)で195840。
    '色から合計の明るさを返す関数
    Private Function ConvertColorArrayToBrightness(ByVal colorAry() As Color) As UInteger
        Dim brightness As UInteger = 0
        For Each clr In colorAry
            brightness += ColorBrightness(clr)
        Next

        Return brightness

    End Function

    '最大値がある値になるように係数を求め、輝度を調整したbmpにする。
    '主にWS2812_CurrentLimitから呼び出される
    Private Sub SatulateBmp(ByRef bmp As Bitmap, ByVal maxValue As UInteger)
        'bmpをカラー配列にする
        Dim clrAry() As Color = ConvertBmpToColorArray(bmp)
        Dim brightness As UInteger = ConvertColorArrayToBrightness(clrAry)
        'Debug.Print("TotalBrightNess " & brightness)
        If brightness > maxValue Then
            Dim decrease As Double = maxValue / brightness
            'Debug.Print("decrease " & decrease)
            DecreaseBmp(bmp, decrease)
        End If
    End Sub

    'とある係数をかけて輝度を調整したbmpにする。
    '主にSatulateBmpから呼び出される。
    Private Sub DecreaseBmp(ByRef bmp As Bitmap, ByVal decreaseValue As Double)
        For x = 0 To bmp.Width - 1
            For y = 0 To bmp.Height - 1
                Dim clrGet As Color = bmp.GetPixel(x, y)
                Dim r, g, b As UInteger
                r = Math.Floor(clrGet.R * decreaseValue)
                g = Math.Floor(clrGet.G * decreaseValue)
                b = Math.Floor(clrGet.B * decreaseValue)

                Dim clrSet As Color = Color.FromArgb(clrGet.A, r, g, b)

                bmp.SetPixel(x, y, clrSet)
            Next
        Next
    End Sub



End Class
