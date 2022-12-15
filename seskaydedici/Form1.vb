Imports System.Runtime.InteropServices
Public Class Form1
    Private Declare Function mciSendString Lib "winmm.dll" Alias "mciSendStringA" (ByVal lpstrCommand As String, ByVal lpstrReturnString As String, ByVal uReturnLength As Integer, ByVal hwndCallback As Integer) As Integer



    Public Sub Kaydedici(ByVal islem As String, ByVal path As String, ByVal isim As String)

        If islem = "baslat" Then

            mciSendString("open new Type waveaudio Alias recsound", "", 0, 0)

            mciSendString("record recsound", "", 0, 0)

        ElseIf islem = "bitir" Then

            mciSendString("save recsound " & path & ".avi", "", 0, 0)

            mciSendString("close recsound ", "", 0, 0)

        End If

    End Sub

    Private Sub Button1_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        Kaydedici("baslat", "", "")
    End Sub



    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles Button2.Click
        SaveFileDialog1.ShowDialog()
        Dim dosyayolu As String

        dosyayolu = SaveFileDialog1.FileName


        Kaydedici("bitir", dosyayolu, "")
    End Sub
End Class
