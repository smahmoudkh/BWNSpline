
Public Class Form1
    Dim b As Bitmap = New Bitmap(100, 100)
    Dim g As Graphics = Graphics.FromImage(b)


    Dim moox(2, 1) As Integer
    'Dim mp(1) As PointF
    Dim iclick As Integer
    Dim poi(1) As PointF
    Dim linsiz As Integer
    Dim controlline As Integer

    Dim bitm As Bitmap = New Bitmap(800, 700)
    Dim gr As Graphics = Graphics.FromImage(bitm)

    Dim imin As Double
    Dim jmin As Double
    Dim mindist As Double
    Dim minpoint(2) As Integer


    Dim c1 As Color
    Dim c2 As Color
    Dim c3 As Color

    Dim sty1 As Drawing2D.DashStyle

    ' Dim sb1 As New SolidBrush(c2)
    Dim fill1 As Boolean        'checker
    Dim bwn As Boolean          'checker
    Dim ctrlpoint As Boolean    'checker
    Dim vwpoint As Boolean      'checker
    Dim selectpoint As Boolean
    Dim mouseisdown As Boolean

    Dim v(1), w(1) As Double





    Private Sub Form1_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        c1 = Color.MediumBlue
        c2 = Color.Yellow
        c3 = Color.Black

        linsiz = 1
        controlline = 1
        fill1 = False
        bwn = True
        ctrlpoint = True
        vwpoint = True
        selectpoint = False
        'mouseisdown = False


    End Sub
    Private Sub PictureBox2_MouseMove(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseMove
        Dim dist(UBound(moox, 2)) As Double
        'Dim mindist As Double
        'Dim minpoint(2) As Integer

        If selectpoint = True Then     ''''' Select Mode '''''
            'Dim dist(UBound(moox, 2)) As Double

            '    For i = 0 To UBound(moox, 2)
            '        dist(i) = ((e.X - moox(1, i)) ^ 2 + (e.Y - moox(2, i)) ^ 2) ^ (0.5)
            '    Next
            '    mindist = dist(0)
            '    minpoint(1) = moox(1, 0)
            '    minpoint(2) = moox(2, 0)
            '    For i = 0 To UBound(moox, 2)
            '        If mindist > dist(i) Then
            '            imin = i
            '            mindist = dist(i)
            '            minpoint(1) = moox(1, i)
            '            minpoint(2) = moox(2, i)
            '        End If
            '    Next

            '    If mindist > 10 Then
            '        mindist = 10
            '        minpoint(1) = 0
            '        minpoint(2) = 0
            '        Form1.ActiveForm.Cursor = Cursors.Cross
            '        
            '    Else
            '       
            '        Form1.ActiveForm.Cursor = Cursors.NoMove2D
            '    End If
            'End If



            If mouseisdown = False Then
                'If mouseisdown = True Then
                For i = 0 To UBound(moox, 2)
                    dist(i) = ((e.X - moox(1, i)) ^ 2 + (e.Y - moox(2, i)) ^ 2) ^ (0.5)
                Next
                mindist = dist(0)
                minpoint(1) = moox(1, 0)
                minpoint(2) = moox(2, 0)
                For i = 0 To UBound(moox, 2)
                    If mindist > dist(i) Then
                        imin = i
                        mindist = dist(i)
                        minpoint(1) = moox(1, i)
                        minpoint(2) = moox(2, i)
                    End If
                Next

                If mindist > 10 Then
                    mindist = 10
                    minpoint(1) = 0
                    minpoint(2) = 0

                    Form1.ActiveForm.Cursor = Cursors.Arrow

                Else

                    Form1.ActiveForm.Cursor = Cursors.NoMove2D
                    'moox(1, imin) = e.X 'minpoint(1)
                    'moox(2, imin) = e.Y 'minpoint(2)
                    'lineee(5, 6)
                    ' End If

                    Label2.Text = imin.ToString
                End If

            Else
                If mindist < 10 Then
                    jmin = imin
                    moox(1, imin) = e.X 'minpoint(1)
                    moox(2, imin) = e.Y 'minpoint(2)
                    TrackBar2.Value = Math.Log(v(imin) + 1) * 10
                    TrackBar3.Value = Math.Log(w(imin)) * 10
                    lineee(5, 6)
                End If

            End If

        End If
    End Sub
    Private Sub PictureBox2_MouseUp(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseUp
        If e.Button = MouseButtons.Left Then
            If selectpoint = True Then  ''''' Select Mode '''''

                '    moox(1, imin) = e.X 'minpoint(1)
                '    moox(2, imin) = e.Y 'minpoint(2)
                '    lineee(5, 6)
                '    Label1.Text = "UUP"
                '    Label2.Text = imin.ToString


                mouseisdown = False

            Else       ''''' Design Mode '''''
                iclick = iclick + 1
                ReDim Preserve moox(2, iclick)
                'ReDim Preserve mp(iclick)

                moox(1, iclick) = e.X
                moox(2, iclick) = e.Y
                Label1.Text = moox(1, iclick).ToString
                Label2.Text = moox(2, iclick).ToString
                lineee(5, 6)

            End If


        End If
    End Sub
    Private Sub PictureBox2_MouseDown(ByVal sender As Object, ByVal e As System.Windows.Forms.MouseEventArgs) Handles PictureBox2.MouseDown
        If e.Button = MouseButtons.Left Then


            '    'MsgBox(minpoint(1).ToString)
            '    Label1.Text = "DDo"

            'Else       ''''' Design Mode '''''

            mouseisdown = True
        End If



    End Sub








    Private Sub Button1_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button1.Click
        GroupBox1.Enabled = False
        selectpoint = False
        Form1.ActiveForm.Cursor = Cursors.Cross

        'Dim dist(UBound(moox, 2)) As Double
        'For i = 1 To UBound(moox, 2)

        'Next



    End Sub
    Private Sub Button2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button2.Click

        'gr.Clear()
        gr.Clear(Drawing.Color.Transparent)
        'Dim moox(2, 1) As Integer = 0
        'Dim iclick As Integer = 0
        ' ReDim Preserve mp(1)
        ReDim Preserve moox(2, 1)
        moox(2, 1) = 0

        iclick = 0
        For i = 1 To UBound(v)
            v(i) = 0
            w(i) = 1
        Next
        TrackBar2.Value = 0



    End Sub

    Private Sub Button3_Click(ByVal sender As System.Object, ByVal e As System.EventArgs)
        ColorDialog1.ShowDialog()
        c1 = ColorDialog1.Color
        ' pen10 = New Pen(c1, 3)
        'gr.DrawLines(pen10, poi)

        lineee(5, 7)
    End Sub
    Private Sub Button4_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button4.Click

        ColorDialog1.ShowDialog()
        c1 = ColorDialog1.Color
        lineee(5, 7)
    End Sub
    Private Sub Button5_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button5.Click

        ColorDialog2.ShowDialog()
        c2 = ColorDialog2.Color
        lineee(5, 7)
    End Sub


    Public Sub lineee(ByVal lop As Integer, ByVal shekl As Integer)
        'gr = Form1.ActiveForm.CreateGraphics
        'gr.Clear(Drawing.Color.White)
        gr.Clear(Drawing.Color.Transparent)
        Dim pen1 As Pen


        pen1 = New Pen(c1, linsiz)
        pen1.DashStyle = sty1
        Dim pen2 As Pen
        pen2 = New Pen(Color.Salmon, 2)
        pen2.DashStyle = Drawing2D.DashStyle.DashDot
        Dim pen3 As Pen
        pen3 = New Pen(c3, controlline)
        pen3.DashStyle = Drawing2D.DashStyle.DashDotDot
        Dim pen4 As Pen
        pen4 = New Pen(Color.Black, 2)
        pen4.DashStyle = Drawing2D.DashStyle.Dot
        Dim pen5 As Pen
        pen5 = New Pen(Color.Cyan, 2)
        pen5.DashStyle = Drawing2D.DashStyle.Solid


        Dim sb1 As New SolidBrush(c2)
        Dim sb2 As New SolidBrush(Color.Gold)
        Dim sb3 As New SolidBrush(Color.Cyan)
        Dim sb4 As New SolidBrush(Color.Linen)
        Dim sb5 As New SolidBrush(Color.SeaGreen)





        'Dim v(1), w(1) As Double
        'Dim lop As Integer = 5
        Dim refine As Integer = 31

        Dim n1 As Integer
        n1 = UBound(moox, 2)
        If n1 > 4 Then


            Dim moox2(2, n1) As Integer
            For i = 1 To n1
                moox2(1, i) = moox(1, i)
                moox2(2, i) = moox(2, i)
            Next
            Dim t(n1 + lop) As Double
            ReDim Preserve v(n1 + lop), w(n1 + lop)
            For i = 1 To n1 + lop
                t(i) = i
                'w(i) = 1
                'v(i) = 0
            Next

            For i = 1 To n1 + lop
                If w(i) = 0 Then
                    w(i) = 1
                End If
            Next
            ' For i = 2 To n1
            'v(i) = 0
            ' Next
            For i = n1 + 1 To n1 + lop
                v(i) = v(i - n1)
                w(i) = w(i - n1)
            Next

            ReDim Preserve moox2(2, n1 + lop)
            For i = 1 To lop
                moox2(1, i + n1) = moox2(1, i)
                moox2(2, i + n1) = moox2(2, i)
            Next i

            Dim n As Integer
            n = UBound(moox2, 2)

            Dim a(n), c(n), h(n), la(n), lah(n), mu(n), muh(n), db(n), d(n), f(2, n), vb(2, n), wb(2, n), p2(2, n * refine) As Double

            For i = 1 To n - 1
                h(i) = t(i + 1) - t(i)
            Next

            For i = 1 To n - 1
                c(i) = w(i) / h(i)
                a(i) = 1 / c(i)
            Next

            For i = 2 To n - 1
                d(i) = a(i) * a(i - 1) * v(i) + a(i) + a(i - 1)
            Next

            For i = 3 To n - 1
                db(i) = h(i - 1) * d(i - 1) * d(i) + a(i) * (h(i - 1) + h(i)) * d(i - 1) + a(i - 2) * (h(i - 2) + h(i - 1)) * d(i)
                la(i) = a(i) * h(i) * d(i - 1) / db(i)
                lah(i) = 3 * a(i) * d(i - 1) / db(i)
            Next

            For i = 2 To n - 2
                mu(i) = a(i - 1) * h(i - 1) * d(i + 1) / db(i + 1)
                muh(i) = 3 * a(i - 1) * d(i + 1) / db(i + 1)
            Next

            For i = 2 To n - 2
                f(1, i) = la(i) * moox2(1, i - 1) + (1 - mu(i) - la(i)) * moox2(1, i) + mu(i) * moox2(1, i + 1)
                f(2, i) = la(i) * moox2(2, i - 1) + (1 - mu(i) - la(i)) * moox2(2, i) + mu(i) * moox2(2, i + 1)
                vb(1, i) = (la(i) - lah(i) / 2 * h(i)) * moox2(1, i - 1) + ((1 - mu(i) - la(i)) + (lah(i) - muh(i)) / 2 * h(i)) * moox2(1, i) + (mu(i) + muh(i) / 2 * h(i)) * moox2(1, i + 1)
                vb(2, i) = (la(i) - lah(i) / 2 * h(i)) * moox2(2, i - 1) + ((1 - mu(i) - la(i)) + (lah(i) - muh(i)) / 2 * h(i)) * moox2(2, i) + (mu(i) + muh(i) / 2 * h(i)) * moox2(2, i + 1)
            Next

            For i = 1 To n - 3
                wb(1, i) = (la(i + 1) + lah(i + 1) / 2 * h(i)) * moox2(1, i) + ((1 - mu(i + 1) - la(i + 1)) - (lah(i + 1) - muh(i + 1)) / 2 * h(i)) * moox2(1, i + 1) + (mu(i + 1) - muh(i + 1) / 2 * h(i)) * moox2(1, i + 2)
                wb(2, i) = (la(i + 1) + lah(i + 1) / 2 * h(i)) * moox2(2, i) + ((1 - mu(i + 1) - la(i + 1)) - (lah(i + 1) - muh(i + 1)) / 2 * h(i)) * moox2(2, i + 1) + (mu(i + 1) - muh(i + 1) / 2 * h(i)) * moox2(2, i + 2)
            Next

            Dim te As Double
            Dim oo As Double
            oo = 0
            For i = 3 To n - 3
                For ti = t(i) To t(i + 1) Step h(i) / (refine - 1)
                    te = (ti - t(i)) / h(i)
                    oo = oo + 1
                    p2(1, oo) = (1 - te) ^ 2 * f(1, i) + 2 * te * (1 - te) ^ 2 * vb(1, i) + 2 * te ^ 2 * (1 - te) * wb(1, i) + te ^ 2 * f(1, i + 1)
                    p2(2, oo) = (1 - te) ^ 2 * f(2, i) + 2 * te * (1 - te) ^ 2 * vb(2, i) + 2 * te ^ 2 * (1 - te) * wb(2, i) + te ^ 2 * f(2, i + 1)

                Next ti
            Next i

            'Dim pen3 As Pen
            'pen3 = New Pen(Color.Black, 2)
            'pen3.DashStyle = Drawing2D.DashStyle.DashDotDot






            'Dim poi(oo - 1) As PointF

            ReDim Preserve poi(oo - 1)
            For k = 0 To oo - 1
                poi(k).X = p2(1, k + 1)
                poi(k).Y = p2(2, k + 1)
            Next


            Dim pV(oo - 1) As PointF
            'ReDim Preserve poi(oo - 3)
            For k = 0 To n - 4
                pV(k).X = vb(1, k + 2)
                pV(k).Y = vb(2, k + 2)
            Next


            Dim pW(oo - 1) As PointF
            'ReDim Preserve poi(oo - 3)
            For k = 0 To n - 4
                pW(k).X = wb(1, k + 2)
                pW(k).Y = wb(2, k + 2)
            Next






            'If shekl = 0 Then
            '    gr.DrawLines(pen10, poi)
            'End If
            'If shekl = 1 Then
            '    gr.FillPolygon(sb1, poi)
            'End If
            'If shekl = 2 Then
            '    gr.DrawPolygon(pen10, mp)
            'End If

            'ColorDialog1.ShowDialog()
            'c1 = ColorDialog1.Color





            'gr = Form1.ActiveForm.CreateGraphics


            If bwn = True Then
                gr.DrawLines(pen1, poi)
            End If
            If fill1 = True Then
                gr.FillPolygon(sb1, poi)
            End If
            'If fill1 = 2 Then
            '    gr.DrawLines(pen1, poi)
            '    'gr.FillPolygon(sb1, poi)
            'End If
            'If fill1 = 3 Then
            '    gr.DrawLines(pen1, poi)
            '    'gr.FillPolygon(sb1, poi)

            'End If

            If vwpoint = True And iclick > 1 Then
                For k = 1 To n - 4
                    gr.DrawRectangle(pen2, pV(k).X, pV(k).Y, 4, 4)
                Next
                For k = 0 To n - 4
                    gr.DrawRectangle(pen5, pW(k).X, pW(k).Y, 4, 4)
                Next
            End If






        End If
        If ctrlpoint = True And iclick > 1 Then
            For k = 1 To n1 - 1
                gr.DrawLine(pen3, moox(1, k), moox(2, k), moox(1, k + 1), moox(2, k + 1))
            Next
            gr.DrawLine(pen3, moox(1, n1), moox(2, n1), moox(1, 1), moox(2, 1))
            For k = 1 To n1
                gr.DrawEllipse(pen3, moox(1, k) - 2, moox(2, k) - 2, 4, 4)
            Next
        End If
        PictureBox2.Image = bitm

    End Sub


    Private Sub TrackBar1_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar1.Scroll
        linsiz = TrackBar1.Value
        lineee(5, 0)
    End Sub

    Private Sub TrackBar2_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar2.Scroll
        v(jmin) = Math.Exp(TrackBar2.Value / 10) - 1
        lineee(5, 5)
    End Sub

    Private Sub TrackBar3_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar3.Scroll
        w(jmin) = Math.Exp(TrackBar3.Value / 10)
        lineee(5, 5)
    End Sub

    Private Sub Button6_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button6.Click
        selectpoint = True
        Form1.ActiveForm.Cursor = Cursors.Arrow
        mouseisdown = False
        GroupBox1.Enabled = True
    End Sub


    Private Sub CheckBox1_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox1.Click
        fill1 = CheckBox1.Checked
        lineee(5, 0)
    End Sub

    Private Sub CheckBox2_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox2.Click
        bwn = CheckBox2.Checked
        lineee(5, 0)
    End Sub

    Private Sub CheckBox3_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox3.Click
        ctrlpoint = CheckBox3.Checked
        lineee(5, 0)
    End Sub
    Private Sub CheckBox4_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles CheckBox4.Click
        vwpoint = CheckBox4.Checked
        lineee(5, 0)
    End Sub

    Private Sub Button3_Click_1(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button3.Click
        PictureBox1.SizeMode = PictureBoxSizeMode.Zoom

    End Sub

    Private Sub Button7_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button7.Click
        OpenFileDialog1.ShowDialog()
        If OpenFileDialog1.FileName <> "" Then
            Dim ima As Image
            ima = Image.FromFile(OpenFileDialog1.FileName)
            'gr.FromImage(ima)

            PictureBox1.Image = ima
            PictureBox1.Refresh()
            PictureBox2.Image = bitm
            PictureBox2.Parent = PictureBox1
            PictureBox2.BackColor = Color.Transparent
            'Dim xNew, yNew As Single
            'xNew = PictureBox2.Location.X - PictureBox1.Location.X
            'yNew = PictureBox2.Location.Y - PictureBox1.Location.Y
            'PictureBox2.Location = New System.Drawing.Point(xNew, yNew)

        End If
    End Sub

    Private Sub Button8_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button8.Click
        SaveFileDialog1.Filter = "Jpg Files|*.jpg"
        SaveFileDialog1.FilterIndex = OpenFileDialog1.FilterIndex
        SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.FileName <> "" Then


            'PictureBox1.Image.FromFile(gr)
            'PictureBox1.Image.Save(SaveFileDialog1.FileName)
            'gr.Save(SaveFileDialog1.FileName)
            'gr.DrawImage()



            bitm.Save(SaveFileDialog1.FileName)

        End If


    End Sub



    Private Sub Button9_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button9.Click
        'g.DrawLine(Pens.Black, 5, 5, 50, 5)


        'g.Dispose()
        SaveFileDialog1.Filter = "Jpg Files|*.jpg"
        SaveFileDialog1.FilterIndex = OpenFileDialog1.FilterIndex
        SaveFileDialog1.ShowDialog()
        If SaveFileDialog1.FileName <> "" Then


            'PictureBox1.Image.FromFile(gr)
            'PictureBox1.Image.Save(SaveFileDialog1.FileName)
            'gr.Save(SaveFileDialog1.FileName)
            'gr.DrawImage()

            'gr.Save()

            'bitm.Save(SaveFileDialog1.FileName)

        End If


    End Sub





    Private Sub PictureBox2_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles PictureBox2.Click

    End Sub

    Private Sub TrackBar4_Scroll(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles TrackBar4.Scroll
        controlline = TrackBar4.Value
        lineee(5, 0)
    End Sub

    Private Sub Button10_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Button10.Click
        ColorDialog3.ShowDialog()
        c3 = ColorDialog3.Color
        lineee(5, 7)
    End Sub



 
    Private Sub ComboBox1_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox1.SelectedIndexChanged
        Select Case ComboBox1.SelectedIndex
            Case 0
                sty1 = Drawing2D.DashStyle.Solid
            Case 1
                sty1 = Drawing2D.DashStyle.DashDot
            Case 2
                sty1 = Drawing2D.DashStyle.DashDotDot
            Case 3
                sty1 = Drawing2D.DashStyle.Dot
            Case 4
                sty1 = Drawing2D.DashStyle.Dash
        End Select
        lineee(5, 7)
    End Sub

  Private Sub ComboBox2_SelectedIndexChanged(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles ComboBox2.SelectedIndexChanged

  End Sub
End Class
