# PAG
Public Class Form1
    Public Count As Integer

    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        'Button 2,5
        CheckButton(Button1, Button2)
        CheckButton(Button1, Button5)

        CheckSolved()
    End Sub

    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click
        '1,3,6
        CheckButton(Button2, Button1)
        CheckButton(Button2, Button3)
        CheckButton(Button2, Button6)

        CheckSolved()
    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        '2,7,4
        CheckButton(Button3, Button2)
        CheckButton(Button3, Button4)
        CheckButton(Button3, Button7)

        CheckSolved()
    End Sub

    Private Sub Button4_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click
        '3,8
        CheckButton(Button4, Button3)
        CheckButton(Button4, Button8)

        CheckSolved()
    End Sub

    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click
        '1,6,9
        CheckButton(Button5, Button1)
        CheckButton(Button5, Button6)
        CheckButton(Button5, Button9)

        CheckSolved()
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        '2,5,7,10
        CheckButton(Button6, Button2)
        CheckButton(Button6, Button5)
        CheckButton(Button6, Button7)
        CheckButton(Button6, Button10)

        CheckSolved()
    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        '3,6,8,11
        CheckButton(Button7, Button3)
        CheckButton(Button7, Button6)
        CheckButton(Button7, Button8)
        CheckButton(Button7, Button11)

        CheckSolved()
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        '4,7,12
        CheckButton(Button8, Button4)
        CheckButton(Button8, Button7)
        CheckButton(Button8, Button12)

        CheckSolved()
    End Sub

    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        '5,10,13
        CheckButton(Button9, Button5)
        CheckButton(Button9, Button10)
        CheckButton(Button9, Button13)

        CheckSolved()
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        '6,9,11,14
        CheckButton(Button10, Button6)
        CheckButton(Button10, Button9)
        CheckButton(Button10, Button11)
        CheckButton(Button10, Button14)

        CheckSolved()
    End Sub

    Private Sub Button11_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button11.Click
        '7,10,15,12
        CheckButton(Button11, Button7)
        CheckButton(Button11, Button10)
        CheckButton(Button11, Button15)
        CheckButton(Button11, Button12)

        CheckSolved()
    End Sub

    Private Sub Button12_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button12.Click
        '11'8
        CheckButton(Button12, Button11)
        CheckButton(Button12, Button8)

        CheckSolved()
    End Sub

    Private Sub Button13_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button13.Click
        '9,14
        CheckButton(Button13, Button9)
        CheckButton(Button13, Button14)

        CheckSolved()
    End Sub

    Private Sub Button14_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button14.Click
        '10,13,15
        CheckButton(Button14, Button10)
        CheckButton(Button14, Button13)
        CheckButton(Button14, Button15)
        CheckButton(Button14, Button15)

        CheckSolved()
    End Sub

    Private Sub Button15_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button15.Click
        '14,11,16
        CheckButton(Button15, Button14)
        CheckButton(Button15, Button11)
        CheckButton(Button15, Button16)

        CheckSolved()
    End Sub

    Private Sub Button16_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button16.Click
        '15,12
        CheckButton(Button16, Button15)
        CheckButton(Button16, Button12)

        CheckSolved()
    End Sub

    Private Sub ToolStripButton2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton2.Click
        Me.Close()
        MsgBox("Sigur doriti sa iesiti?")
    End Sub

    Private Sub Form1_FormClosing(ByVal sender As Object, ByVal e As System.Windows.Forms.FormClosingEventArgs) Handles Me.FormClosing
        Dim x = MsgBox("Sigur doriti sa iesiti? ", vbYesNo + vbQuestion)
        If (x = Windows.Forms.DialogResult.No) Then
            e.Cancel = True
        End If
    End Sub

    Private Sub Form1_Load(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles MyBase.Load
        Shuffle()

    End Sub

    Private Sub ToolStripButton1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ToolStripButton1.Click
        Shuffle()
        Label2.Text = "0"

    End Sub

    Private Sub ToolStrip1_ItemClicked(ByVal sender As System.Object, ByVal e As System.Windows.Forms.ToolStripItemClickedEventArgs) Handles ToolStrip1.ItemClicked

    End Sub
End Class


Module Module1
    Sub CheckButton(ByRef Butt1 As Button, ByRef Butt2 As Button)
        If Butt2.Text = "" Then
            Butt2.Text = Butt1.Text
            Butt1.Text = ""

        End If
    End Sub

    Sub CheckSolved()
        If Form1.Button1.Text = "1" And Form1.Button2.Text = "2" And Form1.Button3.Text = "3" And
            Form1.Button4.Text = "4" And Form1.Button5.Text = "5" And Form1.Button6.Text = "6" And
            Form1.Button7.Text = "7" And Form1.Button8.Text = "8" And Form1.Button9.Text = "9" And
            Form1.Button10.Text = "10" And Form1.Button11.Text = "11" And Form1.Button12.Text = "12" And
            Form1.Button13.Text = "13" And Form1.Button14.Text = "14" And Form1.Button15.Text = "15" Then
            MsgBox("WOW, ai rezolvat Puzzelul din " & Form1.Count & " mutari", vbInformation, "Sufle")
        End If
        Form1.Count = Form1.Count + 1
        Form1.Text = "Shuffle: " & Form1.Count & " mutari"
        Form1.Label2.Text = Form1.Count
    End Sub


    Sub Shuffle()
        Dim a(15), i, j, RN As Integer
        Dim flag As Boolean



        flag = False
        i = 1
        a(j) = 1
        Do While i <= 15
            Randomize()
            RN = CInt(Int((15 * Rnd()) + 1))
            For j = 1 To i
                If (a(j) = RN) Then
                    flag = True
                    Exit For
                End If
            Next
            If flag = True Then
                flag = False
            Else
                a(i) = RN
                i = i + 1

            End If
        Loop
        Form1.Button1.Text = a(1)
        Form1.Button2.Text = a(2)
        Form1.Button3.Text = a(3)
        Form1.Button4.Text = a(4)
        Form1.Button5.Text = a(5)

        Form1.Button6.Text = a(6)
        Form1.Button7.Text = a(7)
        Form1.Button8.Text = a(8)
        Form1.Button9.Text = a(9)
        Form1.Button10.Text = a(10)

        Form1.Button11.Text = a(11)
        Form1.Button12.Text = a(12)
        Form1.Button13.Text = a(13)
        Form1.Button14.Text = a(14)
        Form1.Button15.Text = a(15)
        Form1.Button16.Text = ""



    End Sub


End Module

