Option Explicit On
Public Class Form1
    Dim Time(1) As Double

    Private Sub TestButton_Click(sender As Object, e As EventArgs) Handles Test1.Click, Test2.Click, Test3.Click, Test4.Click
        '记录事件开始时间（精确到1x10^(-6)秒）
        Time(0) = My.Computer.Clock.LocalTime.TimeOfDay.TotalSeconds
        Dim TempPicture As Bitmap = CType(sender, Button).Image
        ResPicture.Image = TempPicture
        '调用图像模糊函数
        PictureBlur(TempPicture, ResPicture, SetPixel.Value)
        '记录事件结束时间
        Time(1) = My.Computer.Clock.LocalTime.TimeOfDay.TotalSeconds
        '显示处理事件耗时
        MsgBox("Task finished!" & vbCrLf & "共用时： 【" & Time(1) - Time(0) & "】 秒。", MsgBoxStyle.Information)
    End Sub

    Private Sub PictureBlur(ByVal PictureRes As Bitmap, PictureBoxName As PictureBox, ByVal Pixel As Long)
        Dim GetPicture As Bitmap = PictureRes
        Dim PicWidth As Long = PictureRes.Width, PicHeight As Long = PictureRes.Height
        Dim PositionX, PositionY As Integer, RoundX, RoundY As Integer
        Dim SumA, SumR, SumG, SumB As Long
        Dim Count As Integer, CenterColor As Color

        For PositionY = 0 To PicHeight - 1
            For PositionX = 0 To PicWidth - 1
                Count = 0 : SumA = 0 : SumR = 0 : SumG = 0 : SumB = 0
                For RoundY = PositionY - Pixel To PositionY + Pixel
                    For RoundX = PositionX - Pixel To PositionX + Pixel
                        If (0 <= RoundX And RoundX < PicWidth) And (0 <= RoundY And RoundY < PicHeight) Then
                            Count += 1
                            SumA += PictureRes.GetPixel(RoundX, RoundY).A
                            SumR += PictureRes.GetPixel(RoundX, RoundY).R
                            SumG += PictureRes.GetPixel(RoundX, RoundY).G
                            SumB += PictureRes.GetPixel(RoundX, RoundY).B
                        End If
                    Next
                Next
                CenterColor = Color.FromArgb(SumA \ Count, SumR \ Count, SumG \ Count, SumB \ Count)
                GetPicture.SetPixel(PositionX, PositionY, CenterColor)
            Next
        Next

        PictureBoxName.Image = GetPicture
    End Sub

End Class
