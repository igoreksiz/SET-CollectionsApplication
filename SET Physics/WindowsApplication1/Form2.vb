Public Class Form2
    Dim botto As Boolean

    Dim X0, Y0 As New Integer
    Dim Newpoint0 As Drawing.Point

    Dim Newpoint1 As Drawing.Point
    Dim Newpoint4 As Drawing.Point
    Dim Newpoint6 As Drawing.Point

    Dim X2, Y2 As Integer
    Dim NewPoint2 As Drawing.Point

    Dim X5, Y5 As New Integer
    Dim Newpoint5 As Drawing.Point

    Dim Answer As New Integer

    Dim key As String

    Private Sub PictureBoxClose_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxClose.Click
        Me.Close()
    End Sub

    Private Sub PictureBoxClose_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxClose.MouseDown
        PictureBoxClose.Image = My.Resources.controlbox_3
    End Sub

    Private Sub PictureBoxClose_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxClose.MouseLeave
        PictureBoxClose.Image = My.Resources.close_1
    End Sub

    Private Sub PictureBoxClose_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxClose.MouseMove
        PictureBoxClose.Image = My.Resources.close_2
    End Sub

    Private Sub PictureBoxMax_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxMax.Click
        If Me.WindowState = FormWindowState.Maximized Then
            Me.WindowState = FormWindowState.Normal
            PanelPage1.Height = Me.Height - PanelHeader.Height - 24
            PanelPage1.Width = Me.Width
            PanelPage2.Height = Me.Height - PanelHeader.Height - 24
            PanelPage2.Width = Me.Width
            PanelPage3.Height = Me.Height - PanelHeader.Height - 24
            PanelPage3.Width = Me.Width
            PictureBoxPage1TestsScrollBarCover.Height = PanelPage1.Height - 2
            PanelPage1Tests.Height = PanelPage1.Height - 2
            PictureBoxPage1TestsScroll.Height = PanelPage1.Height - 2
            PanelPage1Questions.Width = PanelPage1.Width - 138
            PanelPage1Questions.Height = PanelPage1.Height - 2
            PictureBoxPage1QuestionsLeftSide.Height = PanelPage1Questions.Height
            PictureBoxPage1QuestionsScroll.Height = PanelPage1Questions.Height
            PictureBoxPage1QuestionsRightSide.Height = PanelPage1Questions.Height
            PictureBoxPage1QuestionsScrollBarCover.Height = PanelPage1Questions.Height
        Else
            Me.WindowState = FormWindowState.Maximized
            PanelPage1.Height = Me.Height - PanelHeader.Height - 24
            PanelPage1.Width = Me.Width
            PanelPage2.Height = Me.Height - PanelHeader.Height - 24
            PanelPage2.Width = Me.Width
            PanelPage3.Height = Me.Height - PanelHeader.Height - 24
            PanelPage3.Width = Me.Width
            PictureBoxPage1TestsScrollBarCover.Height = PanelPage1.Height - 2
            PanelPage1Tests.Height = PanelPage1.Height - 2
            PictureBoxPage1TestsScroll.Height = PanelPage1.Height - 2

            PanelPage1Questions.Width = PanelPage1.Width - 138
            PanelPage1Questions.Height = PanelPage1.Height - 2
            PictureBoxPage1QuestionsLeftSide.Height = PanelPage1Questions.Height
            PictureBoxPage1QuestionsScroll.Height = PanelPage1Questions.Height
            PictureBoxPage1QuestionsRightSide.Height = PanelPage1Questions.Height
            PictureBoxPage1QuestionsScrollBarCover.Height = PanelPage1Questions.Height
        End If
        Newpoint1.X = PanelPage1.Width - 28
        Newpoint1.Y = 0
        PictureBoxPage1QuestionsRightSide.Location = Newpoint1
        Newpoint4.X = PanelPage1.Width - 15
        Newpoint4.Y = 0
        PictureBoxPage1QuestionsScroll.Location = Newpoint4
        Newpoint5.X = PanelPage1.Width - 25
        Newpoint5.Y = 0
        PictureBoxPage1QuestionsScroller.Location = Newpoint5
        NewPoint2.X = 109
        NewPoint2.Y = 0
        PictureBoxPage1TestsScroller.Location = NewPoint2
        Newpoint6.X = PanelPage1.Width - 27
        Newpoint6.Y = 0
        PictureBoxPage1QuestionsScrollBarCover.Location = Newpoint6

        PanelPage1Questions.VerticalScroll.Value = Newpoint5.Y * ((9000 - PictureBoxPage1QuestionsScroll.Height) / (PictureBoxPage1QuestionsScroll.Height - PictureBoxPage1QuestionsScroller.Height))

        PanelPage1Tests.VerticalScroll.Value = NewPoint2.Y * ((1125 - PictureBoxPage1TestsScroll.Height) / (PictureBoxPage1TestsScroll.Height - PictureBoxPage1TestsScroller.Height))

        PanelMessage.Location = New Drawing.Point(Me.Width - 230, Me.Height - 152)

        PanelPage1Content.Location = New System.Drawing.Point((PanelPage1.Size.Width - PanelPage1Content.Width) / 2, (PanelPage1.Size.Height - PanelPage1Content.Height) / 2)
        PanelPage3PanelRegistration.Location = New System.Drawing.Point((PanelPage3.Size.Width - PanelPage3PanelRegistration.Width) / 2, (PanelPage3.Size.Height - PanelPage3PanelRegistration.Height) / 2)
        PanelPage3PanelSuccess.Location = New System.Drawing.Point((PanelPage3.Size.Width - PanelPage3PanelSuccess.Width) / 2, (PanelPage3.Size.Height - PanelPage3PanelSuccess.Height) / 2)
    End Sub

    Private Sub PictureBoxMax_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxMax.MouseDown
        PictureBoxMax.Image = My.Resources.controlbox_3
    End Sub

    Private Sub PictureBoxMax_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxMax.MouseLeave
        PictureBoxMax.Image = My.Resources.max_1
    End Sub

    Private Sub PictureBoxMax_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxMax.MouseMove
        PictureBoxMax.Image = My.Resources.max_2
    End Sub

    Private Sub PictureBoxMin_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxMin.Click
        Me.WindowState = FormWindowState.Minimized
    End Sub

    Private Sub PictureBoxMin_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxMin.MouseDown
        PictureBoxMin.Image = My.Resources.controlbox_3
    End Sub

    Private Sub PictureBoxMin_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxMin.MouseLeave
        PictureBoxMin.Image = My.Resources.min_1
    End Sub

    Private Sub PictureBoxMin_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxMin.MouseMove
        PictureBoxMin.Image = My.Resources.min_2
    End Sub

    Private Sub PanelHeader_MouseClick(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PanelHeader.MouseClick
        PanelHeader.Cursor = Cursors.Default
    End Sub

    Private Sub PanelHeader_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PanelHeader.MouseDown
        PanelHeader.Cursor = Cursors.SizeAll
        X0 = Control.MousePosition.X - Me.Location.X
        Y0 = Control.MousePosition.Y - Me.Location.Y
    End Sub

    Private Sub PanelHeader_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PanelHeader.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Newpoint0 = Control.MousePosition
            Newpoint0.X -= X0
            Newpoint0.Y -= Y0
            Me.Location = Newpoint0
        End If
    End Sub

    Private Sub Form2_Load(ByVal sender As Object, ByVal e As EventArgs) Handles MyBase.Load
        Dim cpu_name As String
        Dim cpu_name_length As String

        Dim machine_name As String
        Dim machine_name_length As String

        machine_name = System.Environment.MachineName
        machine_name_length = machine_name.Length

        cpu_name = My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\SYSTEM\Centralprocessor\0", "processornamestring", Nothing)
        cpu_name_length = cpu_name.Length

        TextBoxPage3RegistrationNumber.Text = ((12123484482274 + ((cpu_name_length + machine_name_length) * 33344432) + ((My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\SYSTEM\Centralprocessor\0", "~MHz", Nothing)) * (3 + System.Environment.ProcessorCount)) + ((My.Computer.Registry.GetValue("HKEY_LOCAL_MACHINE\HARDWARE\DESCRIPTION\SYSTEM\Centralprocessor\0", "~MHz", Nothing)) * ((3 ^ System.Environment.ProcessorCount) + System.Environment.ProcessorCount))) + System.Environment.ProcessorCount ^ 10)
        key = ((TextBoxPage3RegistrationNumber.Text * 4) - 19933131993313)
        If My.Settings.php_design = "php_design" Then
            PanelPage3PanelRegistration.Visible = False
        Else
            PanelPage3PanelSuccess.Visible = False
            PictureBoxTest02_1.Visible = False
            PictureBoxTest03_1.Visible = False
            PictureBoxTest04_1.Visible = False
            PictureBoxTest05_1.Visible = False
            PictureBoxTest06_1.Visible = False
            PictureBoxTest07_1.Visible = False
            PictureBoxTest08_1.Visible = False
            PictureBoxTest09_1.Visible = False
            PictureBoxTest10_1.Visible = False
            PictureBoxTest11_1.Visible = False
            PictureBoxTest12_1.Visible = False
            PictureBoxTest13_1.Visible = False
            PictureBoxTest14_1.Visible = False
            PictureBoxTest15_1.Visible = False
            PictureBoxTest16_1.Visible = False
            PictureBoxTest17_1.Visible = False
            PictureBoxTest18_1.Visible = False
            PictureBoxTest19_1.Visible = False
            PictureBoxTest20_1.Visible = False
            PictureBoxTest21_1.Visible = False
            PictureBoxTest22_1.Visible = False
            PictureBoxTest23_1.Visible = False
            PictureBoxTest24_1.Visible = False
            PictureBoxTest25_1.Visible = False
        End If

        PanelPage1.Height = Me.Height - PanelHeader.Height - 24
        PanelPage1.Width = Me.Width
        PanelPage2.Height = Me.Height - PanelHeader.Height - 24
        PanelPage2.Width = Me.Width
        PanelPage3.Height = Me.Height - PanelHeader.Height - 24
        PanelPage3.Width = Me.Width
        PictureBoxPage1TestsScrollBarCover.Height = PanelPage1.Height - 2
        PanelPage1Tests.Height = PanelPage1.Height - 2
        PictureBoxPage1TestsScroll.Height = PanelPage1.Height - 2
        NewPoint2.X = 109

        PanelPage1Questions.Width = PanelPage1.Width - 138
        PanelPage1Questions.Height = PanelPage1.Height - 2
        PictureBoxPage1QuestionsLeftSide.Height = PanelPage1Questions.Height
        PictureBoxPage1QuestionsScroll.Height = PanelPage1Questions.Height
        PictureBoxPage1QuestionsRightSide.Height = PanelPage1Questions.Height
        PictureBoxPage1QuestionsScrollBarCover.Height = PanelPage1Questions.Height

        PanelMessage.Location = New Drawing.Point(Me.Width - 230, Me.Height - 152)

        Newpoint1.X = PanelPage1.Width - 28
        Newpoint1.Y = 0
        PictureBoxPage1QuestionsRightSide.Location = Newpoint1
        Newpoint4.X = PanelPage1.Width - 15
        Newpoint4.Y = 0
        PictureBoxPage1QuestionsScroll.Location = Newpoint4
        Newpoint5.X = PanelPage1.Width - 25
        Newpoint5.Y = 0
        PictureBoxPage1QuestionsScroller.Location = Newpoint5
        Newpoint6.X = PanelPage1.Width - 27
        Newpoint6.Y = 0
        PictureBoxPage1QuestionsScrollBarCover.Location = Newpoint6
        PanelPage1.Location = New System.Drawing.Point(0, 24)
        PanelPage2.Location = New System.Drawing.Point(0, 24)
        PanelPage3.Location = New System.Drawing.Point(0, 24)
        PanelPage1Content.Location = New System.Drawing.Point((PanelPage1.Size.Width - PanelPage1Content.Width) / 2, (PanelPage1.Size.Height - PanelPage1Content.Height) / 2)
        PanelPage3PanelRegistration.Location = New System.Drawing.Point((PanelPage3.Size.Width - PanelPage3PanelRegistration.Width) / 2, (PanelPage3.Size.Height - PanelPage3PanelRegistration.Height) / 2)
        PanelPage3PanelSuccess.Location = New System.Drawing.Point((PanelPage3.Size.Width - PanelPage3PanelSuccess.Width) / 2, (PanelPage3.Size.Height - PanelPage3PanelSuccess.Height) / 2)
    End Sub

    Private Sub PictureBoxPage1TestsScroller_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxPage1TestsScroller.MouseDown
        PictureBoxPage1TestsScroller.Image = My.Resources.Scroll_3

        Y2 = Control.MousePosition.Y - PictureBoxPage1TestsScroller.Location.Y
        PanelPage1.Select()
    End Sub

    Private Sub PictureBoxPage1TestsScroller_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxPage1TestsScroller.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            NewPoint2.Y = Control.MousePosition.Y - Y2
            If NewPoint2.Y <= 0 Then
                NewPoint2.Y = 0
                PictureBoxPage1TestsScroller.Location = NewPoint2
            Else
                If NewPoint2.Y >= (PictureBoxPage1TestsScroll.Height - PictureBoxPage1TestsScroller.Height) Then
                    NewPoint2.Y = (PictureBoxPage1TestsScroll.Height - PictureBoxPage1TestsScroller.Height)
                    PictureBoxPage1TestsScroller.Location = NewPoint2
                Else
                    PictureBoxPage1TestsScroller.Location = NewPoint2
                End If
            End If
        End If
        PanelPage1Tests.VerticalScroll.Value = NewPoint2.Y * ((1125 - PictureBoxPage1TestsScroll.Height) / (PictureBoxPage1TestsScroll.Height - PictureBoxPage1TestsScroller.Height))
    End Sub

    Private Sub PictureBoxPage1TestsScroller_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxPage1TestsScroller.MouseLeave
        PictureBoxPage1TestsScroller.Image = My.Resources.Scroll_1
    End Sub

    Private Sub PictureBoxPage1TestsScroller_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxPage1TestsScroller.Click
        PictureBoxPage1TestsScroller.Image = My.Resources.Scroll_1
    End Sub

    Private Sub PictureBoxPage1QuestionsScroller_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxPage1QuestionsScroller.Click
        PictureBoxPage1QuestionsScroller.Image = My.Resources.Scroll_1
    End Sub

    Private Sub PictureBoxPage1QuestionsScroller_MouseDown(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxPage1QuestionsScroller.MouseDown
        PictureBoxPage1QuestionsScroller.Image = My.Resources.Scroll_3

        Y5 = Control.MousePosition.Y - PictureBoxPage1QuestionsScroller.Location.Y
        PanelPage1.Select()
    End Sub

    Private Sub PictureBoxPage1QuestionsScroller_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxPage1QuestionsScroller.MouseLeave
        PictureBoxPage1QuestionsScroller.Image = My.Resources.Scroll_1
    End Sub

    Private Sub PictureBoxPage1QuestionsScroller_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxPage1QuestionsScroller.MouseMove
        If e.Button = Windows.Forms.MouseButtons.Left Then
            Newpoint5.Y = Control.MousePosition.Y - Y5
            If Newpoint5.Y <= 0 Then
                Newpoint5.Y = 0
                PictureBoxPage1QuestionsScroller.Location = Newpoint5
            Else
                If Newpoint5.Y >= (PictureBoxPage1QuestionsScroll.Height - PictureBoxPage1QuestionsScroller.Height) Then
                    Newpoint5.Y = (PictureBoxPage1QuestionsScroll.Height - PictureBoxPage1QuestionsScroller.Height)
                    PictureBoxPage1QuestionsScroller.Location = Newpoint5
                Else
                    PictureBoxPage1QuestionsScroller.Location = Newpoint5
                End If
            End If
        End If
        PanelPage1Questions.VerticalScroll.Value = Newpoint5.Y * ((PanelFinish.Size.Height + 1820 + PictureBoxQuestion01.Size.Height + PictureBoxQuestion02.Size.Height + PictureBoxQuestion03.Size.Height + PictureBoxQuestion04.Size.Height + PictureBoxQuestion05.Size.Height + PictureBoxQuestion06.Size.Height + PictureBoxQuestion07.Size.Height + PictureBoxQuestion08.Size.Height + PictureBoxQuestion09.Size.Height + PictureBoxQuestion10.Size.Height + PictureBoxQuestion11.Size.Height + PictureBoxQuestion12.Size.Height + PictureBoxQuestion13.Size.Height + PictureBoxQuestion14.Size.Height + PictureBoxQuestion15.Size.Height + PictureBoxQuestion16.Size.Height + PictureBoxQuestion17.Size.Height + PictureBoxQuestion18.Size.Height + PictureBoxQuestion19.Size.Height + PictureBoxQuestion20.Size.Height + PictureBoxQuestion21.Size.Height + PictureBoxQuestion22.Size.Height + PictureBoxQuestion23.Size.Height + PictureBoxQuestion24.Size.Height + PictureBoxQuestion25.Size.Height + PictureBoxQuestion26.Size.Height + PictureBoxQuestion27.Size.Height + PictureBoxQuestion28.Size.Height + PictureBoxQuestion29.Size.Height + PictureBoxQuestion30.Size.Height - PictureBoxPage1QuestionsScroll.Height) / (PictureBoxPage1QuestionsScroll.Height - PictureBoxPage1QuestionsScroller.Height))
    End Sub

    Private Sub PanelPage1Questions_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PanelPage1Questions.MouseMove
        PanelPage1Questions.Select()
    End Sub

    Private Sub PanelPage1Questions_MouseWheel(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PanelPage1Questions.MouseWheel
        Newpoint5.Y = (PanelPage1Questions.VerticalScroll.Value) * (PictureBoxPage1QuestionsScroll.Size.Height - 21) / (PanelFinish.Size.Height + 1820 + PictureBoxQuestion01.Size.Height + PictureBoxQuestion02.Size.Height + PictureBoxQuestion03.Size.Height + PictureBoxQuestion04.Size.Height + PictureBoxQuestion05.Size.Height + PictureBoxQuestion06.Size.Height + PictureBoxQuestion07.Size.Height + PictureBoxQuestion08.Size.Height + PictureBoxQuestion09.Size.Height + PictureBoxQuestion10.Size.Height + PictureBoxQuestion11.Size.Height + PictureBoxQuestion12.Size.Height + PictureBoxQuestion13.Size.Height + PictureBoxQuestion14.Size.Height + PictureBoxQuestion15.Size.Height + PictureBoxQuestion16.Size.Height + PictureBoxQuestion17.Size.Height + PictureBoxQuestion18.Size.Height + PictureBoxQuestion19.Size.Height + PictureBoxQuestion20.Size.Height + PictureBoxQuestion21.Size.Height + PictureBoxQuestion22.Size.Height + PictureBoxQuestion23.Size.Height + PictureBoxQuestion24.Size.Height + PictureBoxQuestion25.Size.Height + PictureBoxQuestion26.Size.Height + PictureBoxQuestion27.Size.Height + PictureBoxQuestion28.Size.Height + PictureBoxQuestion29.Size.Height + PictureBoxQuestion30.Size.Height - PictureBoxPage1QuestionsScroll.Height)
        PictureBoxPage1QuestionsScroller.Location = Newpoint5
    End Sub

    Private Sub PanelPage1Questions_Scroll(ByVal sender As Object, ByVal e As ScrollEventArgs) Handles PanelPage1Questions.Scroll
        Newpoint5.Y = PanelPage1Questions.VerticalScroll.Value * (PanelFinish.Size.Height + 1820 + PictureBoxQuestion01.Size.Height + PictureBoxQuestion02.Size.Height + PictureBoxQuestion03.Size.Height + PictureBoxQuestion04.Size.Height + PictureBoxQuestion05.Size.Height + PictureBoxQuestion06.Size.Height + PictureBoxQuestion07.Size.Height + PictureBoxQuestion08.Size.Height + PictureBoxQuestion09.Size.Height + PictureBoxQuestion10.Size.Height + PictureBoxQuestion11.Size.Height + PictureBoxQuestion12.Size.Height + PictureBoxQuestion13.Size.Height + PictureBoxQuestion14.Size.Height + PictureBoxQuestion15.Size.Height + PictureBoxQuestion16.Size.Height + PictureBoxQuestion17.Size.Height + PictureBoxQuestion18.Size.Height + PictureBoxQuestion19.Size.Height + PictureBoxQuestion20.Size.Height + PictureBoxQuestion21.Size.Height + PictureBoxQuestion22.Size.Height + PictureBoxQuestion23.Size.Height + PictureBoxQuestion24.Size.Height + PictureBoxQuestion25.Size.Height + PictureBoxQuestion26.Size.Height + PictureBoxQuestion27.Size.Height + PictureBoxQuestion28.Size.Height + PictureBoxQuestion29.Size.Height + PictureBoxQuestion30.Size.Height - PictureBoxPage1QuestionsScroll.Height)
        PictureBoxPage1QuestionsScroller.Location = Newpoint5
    End Sub

    Private Sub PictureBoxTest01_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest01_1.MouseLeave
        PictureBoxTest01_1.Image = My.Resources.Test_01_1
    End Sub

    Private Sub PictureBoxTest01_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest01_1.MouseMove
        PictureBoxTest01_1.Image = My.Resources.Test_01_2
    End Sub

    Private Sub PictureBoxTest02_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest02_1.MouseLeave
        PictureBoxTest02_1.Image = My.Resources.Test_02_1
    End Sub

    Private Sub PictureBoxTest02_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest02_1.MouseMove
        PictureBoxTest02_1.Image = My.Resources.Test_02_2
    End Sub

    Private Sub PictureBoxTest03_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest03_1.MouseLeave
        PictureBoxTest03_1.Image = My.Resources.Test_03_1
    End Sub

    Private Sub PictureBoxTest03_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest03_1.MouseMove
        PictureBoxTest03_1.Image = My.Resources.Test_03_2
    End Sub

    Private Sub PictureBoxTest04_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest04_1.MouseLeave
        PictureBoxTest04_1.Image = My.Resources.Test_04_1
    End Sub

    Private Sub PictureBoxTest04_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest04_1.MouseMove
        PictureBoxTest04_1.Image = My.Resources.Test_04_2
    End Sub

    Private Sub PictureBoxTest05_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest05_1.MouseLeave
        PictureBoxTest05_1.Image = My.Resources.Test_05_1
    End Sub

    Private Sub PictureBoxTest05_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest05_1.MouseMove
        PictureBoxTest05_1.Image = My.Resources.Test_05_2
    End Sub

    Private Sub PictureBoxTest06_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest06_1.MouseLeave
        PictureBoxTest06_1.Image = My.Resources.Test_06_1
    End Sub

    Private Sub PictureBoxTest06_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest06_1.MouseMove
        PictureBoxTest06_1.Image = My.Resources.Test_06_2
    End Sub

    Private Sub PictureBoxTest07_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest07_1.MouseLeave
        PictureBoxTest07_1.Image = My.Resources.Test_07_1
    End Sub

    Private Sub PictureBoxTest07_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest07_1.MouseMove
        PictureBoxTest07_1.Image = My.Resources.Test_07_2
    End Sub

    Private Sub PictureBoxTest08_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest08_1.MouseLeave
        PictureBoxTest08_1.Image = My.Resources.Test_08_1
    End Sub

    Private Sub PictureBoxTest08_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest08_1.MouseMove
        PictureBoxTest08_1.Image = My.Resources.Test_08_2
    End Sub

    Private Sub PictureBoxTest09_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest09_1.MouseLeave
        PictureBoxTest09_1.Image = My.Resources.Test_09_1
    End Sub

    Private Sub PictureBoxTest09_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest09_1.MouseMove
        PictureBoxTest09_1.Image = My.Resources.Test_09_2
    End Sub

    Private Sub PictureBoxTest10_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest10_1.MouseLeave
        PictureBoxTest10_1.Image = My.Resources.Test_10_1
    End Sub

    Private Sub PictureBoxTest10_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest10_1.MouseMove
        PictureBoxTest10_1.Image = My.Resources.Test_10_2
    End Sub

    Private Sub PictureBoxTest11_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest11_1.MouseLeave
        PictureBoxTest11_1.Image = My.Resources.Test_11_1
    End Sub

    Private Sub PictureBoxTest11_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest11_1.MouseMove
        PictureBoxTest11_1.Image = My.Resources.Test_11_2
    End Sub

    Private Sub PictureBoxTest12_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest12_1.MouseLeave
        PictureBoxTest12_1.Image = My.Resources.Test_12_1
    End Sub

    Private Sub PictureBoxTest12_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest12_1.MouseMove
        PictureBoxTest12_1.Image = My.Resources.Test_12_2
    End Sub

    Private Sub PictureBoxTest13_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest13_1.MouseLeave
        PictureBoxTest13_1.Image = My.Resources.Test_13_1
    End Sub

    Private Sub PictureBoxTest13_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest13_1.MouseMove
        PictureBoxTest13_1.Image = My.Resources.Test_13_2
    End Sub

    Private Sub PictureBoxTest14_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest14_1.MouseLeave
        PictureBoxTest14_1.Image = My.Resources.Test_14_1
    End Sub

    Private Sub PictureBoxTest14_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest14_1.MouseMove
        PictureBoxTest14_1.Image = My.Resources.Test_14_2
    End Sub

    Private Sub PictureBoxTest15_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest15_1.MouseLeave
        PictureBoxTest15_1.Image = My.Resources.Test_15_1
    End Sub

    Private Sub PictureBoxTest15_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest15_1.MouseMove
        PictureBoxTest15_1.Image = My.Resources.Test_15_2
    End Sub

    Private Sub PictureBoxTest16_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest16_1.MouseLeave
        PictureBoxTest16_1.Image = My.Resources.Test_16_1
    End Sub

    Private Sub PictureBoxTest16_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest16_1.MouseMove
        PictureBoxTest16_1.Image = My.Resources.Test_16_2
    End Sub

    Private Sub PictureBoxTest17_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest17_1.MouseLeave
        PictureBoxTest17_1.Image = My.Resources.Test_17_1
    End Sub

    Private Sub PictureBoxTest17_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest17_1.MouseMove
        PictureBoxTest17_1.Image = My.Resources.Test_17_2
    End Sub

    Private Sub PictureBoxTest18_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest18_1.MouseLeave
        PictureBoxTest18_1.Image = My.Resources.Test_18_1
    End Sub

    Private Sub PictureBoxTest18_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest18_1.MouseMove
        PictureBoxTest18_1.Image = My.Resources.Test_18_2
    End Sub

    Private Sub PictureBoxTest19_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest19_1.MouseLeave
        PictureBoxTest19_1.Image = My.Resources.Test_19_1
    End Sub

    Private Sub PictureBoxTest19_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest19_1.MouseMove
        PictureBoxTest19_1.Image = My.Resources.Test_19_2
    End Sub

    Private Sub PictureBoxTest20_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest20_1.MouseLeave
        PictureBoxTest20_1.Image = My.Resources.Test_20_1
    End Sub

    Private Sub PictureBoxTest20_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest20_1.MouseMove
        PictureBoxTest20_1.Image = My.Resources.Test_20_2
    End Sub

    Private Sub PictureBoxTest21_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest21_1.MouseLeave
        PictureBoxTest21_1.Image = My.Resources.Test_21_1
    End Sub

    Private Sub PictureBoxTest21_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest21_1.MouseMove
        PictureBoxTest21_1.Image = My.Resources.Test_21_2
    End Sub

    Private Sub PictureBoxTest22_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest22_1.MouseLeave
        PictureBoxTest22_1.Image = My.Resources.Test_22_1
    End Sub

    Private Sub PictureBoxTest22_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest22_1.MouseMove
        PictureBoxTest22_1.Image = My.Resources.Test_22_2
    End Sub

    Private Sub PictureBoxTest23_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest23_1.MouseLeave
        PictureBoxTest23_1.Image = My.Resources.Test_23_1
    End Sub

    Private Sub PictureBoxTest23_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest23_1.MouseMove
        PictureBoxTest23_1.Image = My.Resources.Test_23_2
    End Sub

    Private Sub PictureBoxTest24_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest24_1.MouseLeave
        PictureBoxTest24_1.Image = My.Resources.Test_24_1
    End Sub

    Private Sub PictureBoxTest24_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest24_1.MouseMove
        PictureBoxTest24_1.Image = My.Resources.Test_24_2
    End Sub

    Private Sub PictureBoxTest25_1_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTest25_1.MouseLeave
        PictureBoxTest25_1.Image = My.Resources.Test_25_1
    End Sub

    Private Sub PictureBoxTest25_1_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxTest25_1.MouseMove
        PictureBoxTest25_1.Image = My.Resources.Test_25_2
    End Sub

    Private Sub A01_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A01_1.Click
        A01_3.Visible = True
        B01_3.Visible = False
        C01_3.Visible = False
        D01_3.Visible = False
        E01_3.Visible = False
    End Sub

    Private Sub B01_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B01_1.Click
        B01_3.Visible = True
        A01_3.Visible = False
        C01_3.Visible = False
        D01_3.Visible = False
        E01_3.Visible = False
    End Sub

    Private Sub C01_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C01_1.Click
        C01_3.Visible = True
        A01_3.Visible = False
        B01_3.Visible = False
        D01_3.Visible = False
        E01_3.Visible = False
    End Sub

    Private Sub D01_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D01_1.Click
        D01_3.Visible = True
        A01_3.Visible = False
        B01_3.Visible = False
        C01_3.Visible = False
        E01_3.Visible = False
    End Sub

    Private Sub E01_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E01_1.Click
        E01_3.Visible = True
        A01_3.Visible = False
        B01_3.Visible = False
        C01_3.Visible = False
        D01_3.Visible = False
    End Sub

    Private Sub A02_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A02_1.Click
        A02_3.Visible = True
        B02_3.Visible = False
        C02_3.Visible = False
        D02_3.Visible = False
        E02_3.Visible = False
    End Sub

    Private Sub B02_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B02_1.Click
        B02_3.Visible = True
        A02_3.Visible = False
        C02_3.Visible = False
        D02_3.Visible = False
        E02_3.Visible = False
    End Sub

    Private Sub C02_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C02_1.Click
        C02_3.Visible = True
        A02_3.Visible = False
        B02_3.Visible = False
        D02_3.Visible = False
        E02_3.Visible = False
    End Sub

    Private Sub D02_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D02_1.Click
        D02_3.Visible = True
        A02_3.Visible = False
        B02_3.Visible = False
        C02_3.Visible = False
        E02_3.Visible = False
    End Sub

    Private Sub E02_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E02_1.Click
        E02_3.Visible = True
        A02_3.Visible = False
        B02_3.Visible = False
        C02_3.Visible = False
        D02_3.Visible = False
    End Sub

    Private Sub A03_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A03_1.Click
        A03_3.Visible = True
        B03_3.Visible = False
        C03_3.Visible = False
        D03_3.Visible = False
        E03_3.Visible = False
    End Sub

    Private Sub B03_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B03_1.Click
        B03_3.Visible = True
        A03_3.Visible = False
        C03_3.Visible = False
        D03_3.Visible = False
        E03_3.Visible = False
    End Sub

    Private Sub C03_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C03_1.Click
        C03_3.Visible = True
        A03_3.Visible = False
        B03_3.Visible = False
        D03_3.Visible = False
        E03_3.Visible = False
    End Sub

    Private Sub D03_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D03_1.Click
        D03_3.Visible = True
        A03_3.Visible = False
        B03_3.Visible = False
        C03_3.Visible = False
        E03_3.Visible = False
    End Sub

    Private Sub E03_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E03_1.Click
        E03_3.Visible = True
        A03_3.Visible = False
        B03_3.Visible = False
        C03_3.Visible = False
        D03_3.Visible = False
    End Sub

    Private Sub A04_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A04_1.Click
        A04_3.Visible = True
        B04_3.Visible = False
        C04_3.Visible = False
        D04_3.Visible = False
        E04_3.Visible = False
    End Sub

    Private Sub B04_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B04_1.Click
        B04_3.Visible = True
        A04_3.Visible = False
        C04_3.Visible = False
        D04_3.Visible = False
        E04_3.Visible = False
    End Sub

    Private Sub C04_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C04_1.Click
        C04_3.Visible = True
        A04_3.Visible = False
        B04_3.Visible = False
        D04_3.Visible = False
        E04_3.Visible = False
    End Sub

    Private Sub D04_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D04_1.Click
        D04_3.Visible = True
        A04_3.Visible = False
        B04_3.Visible = False
        C04_3.Visible = False
        E04_3.Visible = False
    End Sub

    Private Sub E04_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E04_1.Click
        E04_3.Visible = True
        A04_3.Visible = False
        B04_3.Visible = False
        C04_3.Visible = False
        D04_3.Visible = False
    End Sub

    Private Sub A05_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A05_1.Click
        A05_3.Visible = True
        B05_3.Visible = False
        C05_3.Visible = False
        D05_3.Visible = False
        E05_3.Visible = False
    End Sub

    Private Sub B05_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B05_1.Click
        B05_3.Visible = True
        A05_3.Visible = False
        C05_3.Visible = False
        D05_3.Visible = False
        E05_3.Visible = False
    End Sub

    Private Sub C05_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C05_1.Click
        C05_3.Visible = True
        A05_3.Visible = False
        B05_3.Visible = False
        D05_3.Visible = False
        E05_3.Visible = False
    End Sub

    Private Sub D05_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D05_1.Click
        D05_3.Visible = True
        A05_3.Visible = False
        B05_3.Visible = False
        C05_3.Visible = False
        E05_3.Visible = False
    End Sub

    Private Sub E05_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E05_1.Click
        E05_3.Visible = True
        A05_3.Visible = False
        B05_3.Visible = False
        C05_3.Visible = False
        D05_3.Visible = False
    End Sub

    Private Sub A06_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A06_1.Click
        A06_3.Visible = True
        B06_3.Visible = False
        C06_3.Visible = False
        D06_3.Visible = False
        E06_3.Visible = False
    End Sub

    Private Sub B06_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B06_1.Click
        B06_3.Visible = True
        A06_3.Visible = False
        C06_3.Visible = False
        D06_3.Visible = False
        E06_3.Visible = False
    End Sub

    Private Sub C06_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C06_1.Click
        C06_3.Visible = True
        A06_3.Visible = False
        B06_3.Visible = False
        D06_3.Visible = False
        E06_3.Visible = False
    End Sub

    Private Sub D06_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D06_1.Click
        D06_3.Visible = True
        A06_3.Visible = False
        B06_3.Visible = False
        C06_3.Visible = False
        E06_3.Visible = False
    End Sub

    Private Sub E06_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E06_1.Click
        E06_3.Visible = True
        A06_3.Visible = False
        B06_3.Visible = False
        C06_3.Visible = False
        D06_3.Visible = False
    End Sub

    Private Sub A07_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A07_1.Click
        A07_3.Visible = True
        B07_3.Visible = False
        C07_3.Visible = False
        D07_3.Visible = False
        E07_3.Visible = False
    End Sub

    Private Sub B07_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B07_1.Click
        B07_3.Visible = True
        A07_3.Visible = False
        C07_3.Visible = False
        D07_3.Visible = False
        E07_3.Visible = False
    End Sub

    Private Sub C07_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C07_1.Click
        C07_3.Visible = True
        A07_3.Visible = False
        B07_3.Visible = False
        D07_3.Visible = False
        E07_3.Visible = False
    End Sub

    Private Sub D07_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D07_1.Click
        D07_3.Visible = True
        A07_3.Visible = False
        B07_3.Visible = False
        C07_3.Visible = False
        E07_3.Visible = False
    End Sub

    Private Sub E07_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E07_1.Click
        E07_3.Visible = True
        A07_3.Visible = False
        B07_3.Visible = False
        C07_3.Visible = False
        D07_3.Visible = False
    End Sub

    Private Sub A08_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A08_1.Click
        A08_3.Visible = True
        B08_3.Visible = False
        C08_3.Visible = False
        D08_3.Visible = False
        E08_3.Visible = False
    End Sub

    Private Sub B08_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B08_1.Click
        B08_3.Visible = True
        A08_3.Visible = False
        C08_3.Visible = False
        D08_3.Visible = False
        E08_3.Visible = False
    End Sub

    Private Sub C08_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C08_1.Click
        C08_3.Visible = True
        A08_3.Visible = False
        B08_3.Visible = False
        D08_3.Visible = False
        E08_3.Visible = False
    End Sub

    Private Sub D08_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D08_1.Click
        D08_3.Visible = True
        A08_3.Visible = False
        B08_3.Visible = False
        C08_3.Visible = False
        E08_3.Visible = False
    End Sub

    Private Sub E08_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E08_1.Click
        E08_3.Visible = True
        A08_3.Visible = False
        B08_3.Visible = False
        C08_3.Visible = False
        D08_3.Visible = False
    End Sub

    Private Sub A09_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A09_1.Click
        A09_3.Visible = True
        B09_3.Visible = False
        C09_3.Visible = False
        D09_3.Visible = False
        E09_3.Visible = False
    End Sub

    Private Sub B09_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B09_1.Click
        B09_3.Visible = True
        A09_3.Visible = False
        C09_3.Visible = False
        D09_3.Visible = False
        E09_3.Visible = False
    End Sub

    Private Sub C09_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C09_1.Click
        C09_3.Visible = True
        A09_3.Visible = False
        B09_3.Visible = False
        D09_3.Visible = False
        E09_3.Visible = False
    End Sub

    Private Sub D09_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D09_1.Click
        D09_3.Visible = True
        A09_3.Visible = False
        B09_3.Visible = False
        C09_3.Visible = False
        E09_3.Visible = False
    End Sub

    Private Sub E09_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E09_1.Click
        E09_3.Visible = True
        A09_3.Visible = False
        B09_3.Visible = False
        C09_3.Visible = False
        D09_3.Visible = False
    End Sub

    Private Sub A10_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A10_1.Click
        A10_3.Visible = True
        B10_3.Visible = False
        C10_3.Visible = False
        D10_3.Visible = False
        E10_3.Visible = False
    End Sub

    Private Sub B10_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B10_1.Click
        B10_3.Visible = True
        A10_3.Visible = False
        C10_3.Visible = False
        D10_3.Visible = False
        E10_3.Visible = False
    End Sub

    Private Sub C10_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C10_1.Click
        C10_3.Visible = True
        A10_3.Visible = False
        B10_3.Visible = False
        D10_3.Visible = False
        E10_3.Visible = False
    End Sub

    Private Sub D10_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D10_1.Click
        D10_3.Visible = True
        A10_3.Visible = False
        B10_3.Visible = False
        C10_3.Visible = False
        E10_3.Visible = False
    End Sub

    Private Sub E10_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E10_1.Click
        E10_3.Visible = True
        A10_3.Visible = False
        B10_3.Visible = False
        C10_3.Visible = False
        D10_3.Visible = False
    End Sub

    Private Sub A11_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A11_1.Click
        A11_3.Visible = True
        B11_3.Visible = False
        C11_3.Visible = False
        D11_3.Visible = False
        E11_3.Visible = False
    End Sub

    Private Sub B11_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B11_1.Click
        B11_3.Visible = True
        A11_3.Visible = False
        C11_3.Visible = False
        D11_3.Visible = False
        E11_3.Visible = False
    End Sub

    Private Sub C11_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C11_1.Click
        C11_3.Visible = True
        A11_3.Visible = False
        B11_3.Visible = False
        D11_3.Visible = False
        E11_3.Visible = False
    End Sub

    Private Sub D11_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D11_1.Click
        D11_3.Visible = True
        A11_3.Visible = False
        B11_3.Visible = False
        C11_3.Visible = False
        E11_3.Visible = False
    End Sub

    Private Sub E11_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E11_1.Click
        E11_3.Visible = True
        A11_3.Visible = False
        B11_3.Visible = False
        C11_3.Visible = False
        D11_3.Visible = False
    End Sub

    Private Sub A12_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A12_1.Click
        A12_3.Visible = True
        B12_3.Visible = False
        C12_3.Visible = False
        D12_3.Visible = False
        E12_3.Visible = False
    End Sub

    Private Sub B12_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B12_1.Click
        B12_3.Visible = True
        A12_3.Visible = False
        C12_3.Visible = False
        D12_3.Visible = False
        E12_3.Visible = False
    End Sub

    Private Sub C12_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C12_1.Click
        C12_3.Visible = True
        A12_3.Visible = False
        B12_3.Visible = False
        D12_3.Visible = False
        E12_3.Visible = False
    End Sub

    Private Sub D12_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D12_1.Click
        D12_3.Visible = True
        A12_3.Visible = False
        B12_3.Visible = False
        C12_3.Visible = False
        E12_3.Visible = False
    End Sub

    Private Sub E12_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E12_1.Click
        E12_3.Visible = True
        A12_3.Visible = False
        B12_3.Visible = False
        C12_3.Visible = False
        D12_3.Visible = False
    End Sub

    Private Sub A13_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A13_1.Click
        A13_3.Visible = True
        B13_3.Visible = False
        C13_3.Visible = False
        D13_3.Visible = False
        E13_3.Visible = False
    End Sub

    Private Sub B13_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B13_1.Click
        B13_3.Visible = True
        A13_3.Visible = False
        C13_3.Visible = False
        D13_3.Visible = False
        E13_3.Visible = False
    End Sub

    Private Sub C13_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C13_1.Click
        C13_3.Visible = True
        A13_3.Visible = False
        B13_3.Visible = False
        D13_3.Visible = False
        E13_3.Visible = False
    End Sub

    Private Sub D13_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D13_1.Click
        D13_3.Visible = True
        A13_3.Visible = False
        B13_3.Visible = False
        C13_3.Visible = False
        E13_3.Visible = False
    End Sub

    Private Sub E13_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E13_1.Click
        E13_3.Visible = True
        A13_3.Visible = False
        B13_3.Visible = False
        C13_3.Visible = False
        D13_3.Visible = False
    End Sub

    Private Sub A14_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A14_1.Click
        A14_3.Visible = True
        B14_3.Visible = False
        C14_3.Visible = False
        D14_3.Visible = False
        E14_3.Visible = False
    End Sub

    Private Sub B14_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B14_1.Click
        B14_3.Visible = True
        A14_3.Visible = False
        C14_3.Visible = False
        D14_3.Visible = False
        E14_3.Visible = False
    End Sub

    Private Sub C14_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C14_1.Click
        C14_3.Visible = True
        A14_3.Visible = False
        B14_3.Visible = False
        D14_3.Visible = False
        E14_3.Visible = False
    End Sub

    Private Sub D14_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D14_1.Click
        D14_3.Visible = True
        A14_3.Visible = False
        B14_3.Visible = False
        C14_3.Visible = False
        E14_3.Visible = False
    End Sub

    Private Sub E14_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E14_1.Click
        E14_3.Visible = True
        A14_3.Visible = False
        B14_3.Visible = False
        C14_3.Visible = False
        D14_3.Visible = False
    End Sub

    Private Sub A15_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A15_1.Click
        A15_3.Visible = True
        B15_3.Visible = False
        C15_3.Visible = False
        D15_3.Visible = False
        E15_3.Visible = False
    End Sub

    Private Sub B15_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B15_1.Click
        B15_3.Visible = True
        A15_3.Visible = False
        C15_3.Visible = False
        D15_3.Visible = False
        E15_3.Visible = False
    End Sub

    Private Sub C15_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C15_1.Click
        C15_3.Visible = True
        A15_3.Visible = False
        B15_3.Visible = False
        D15_3.Visible = False
        E15_3.Visible = False
    End Sub

    Private Sub D15_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D15_1.Click
        D15_3.Visible = True
        A15_3.Visible = False
        B15_3.Visible = False
        C15_3.Visible = False
        E15_3.Visible = False
    End Sub

    Private Sub E15_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E15_1.Click
        E15_3.Visible = True
        A15_3.Visible = False
        B15_3.Visible = False
        C15_3.Visible = False
        D15_3.Visible = False
    End Sub

    Private Sub A16_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A16_1.Click
        A16_3.Visible = True
        B16_3.Visible = False
        C16_3.Visible = False
        D16_3.Visible = False
        E16_3.Visible = False
    End Sub

    Private Sub B16_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B16_1.Click
        B16_3.Visible = True
        A16_3.Visible = False
        C16_3.Visible = False
        D16_3.Visible = False
        E16_3.Visible = False
    End Sub

    Private Sub C16_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C16_1.Click
        C16_3.Visible = True
        A16_3.Visible = False
        B16_3.Visible = False
        D16_3.Visible = False
        E16_3.Visible = False
    End Sub

    Private Sub D16_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D16_1.Click
        D16_3.Visible = True
        A16_3.Visible = False
        B16_3.Visible = False
        C16_3.Visible = False
        E16_3.Visible = False
    End Sub

    Private Sub E16_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E16_1.Click
        E16_3.Visible = True
        A16_3.Visible = False
        B16_3.Visible = False
        C16_3.Visible = False
        D16_3.Visible = False
    End Sub

    Private Sub A17_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A17_1.Click
        A17_3.Visible = True
        B17_3.Visible = False
        C17_3.Visible = False
        D17_3.Visible = False
        E17_3.Visible = False
    End Sub

    Private Sub B17_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B17_1.Click
        B17_3.Visible = True
        A17_3.Visible = False
        C17_3.Visible = False
        D17_3.Visible = False
        E17_3.Visible = False
    End Sub

    Private Sub C17_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C17_1.Click
        C17_3.Visible = True
        A17_3.Visible = False
        B17_3.Visible = False
        D17_3.Visible = False
        E17_3.Visible = False
    End Sub

    Private Sub D17_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D17_1.Click
        D17_3.Visible = True
        A17_3.Visible = False
        B17_3.Visible = False
        C17_3.Visible = False
        E17_3.Visible = False
    End Sub

    Private Sub E17_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E17_1.Click
        E17_3.Visible = True
        A17_3.Visible = False
        B17_3.Visible = False
        C17_3.Visible = False
        D17_3.Visible = False
    End Sub

    Private Sub A18_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A18_1.Click
        A18_3.Visible = True
        B18_3.Visible = False
        C18_3.Visible = False
        D18_3.Visible = False
        E18_3.Visible = False
    End Sub

    Private Sub B18_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B18_1.Click
        B18_3.Visible = True
        A18_3.Visible = False
        C18_3.Visible = False
        D18_3.Visible = False
        E18_3.Visible = False
    End Sub

    Private Sub C18_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C18_1.Click
        C18_3.Visible = True
        A18_3.Visible = False
        B18_3.Visible = False
        D18_3.Visible = False
        E18_3.Visible = False
    End Sub

    Private Sub D18_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D18_1.Click
        D18_3.Visible = True
        A18_3.Visible = False
        B18_3.Visible = False
        C18_3.Visible = False
        E18_3.Visible = False
    End Sub

    Private Sub E18_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E18_1.Click
        E18_3.Visible = True
        A18_3.Visible = False
        B18_3.Visible = False
        C18_3.Visible = False
        D18_3.Visible = False
    End Sub

    Private Sub A19_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A19_1.Click
        A19_3.Visible = True
        B19_3.Visible = False
        C19_3.Visible = False
        D19_3.Visible = False
        E19_3.Visible = False
    End Sub

    Private Sub B19_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B19_1.Click
        B19_3.Visible = True
        A19_3.Visible = False
        C19_3.Visible = False
        D19_3.Visible = False
        E19_3.Visible = False
    End Sub

    Private Sub C19_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C19_1.Click
        C19_3.Visible = True
        A19_3.Visible = False
        B19_3.Visible = False
        D19_3.Visible = False
        E19_3.Visible = False
    End Sub

    Private Sub D19_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D19_1.Click
        D19_3.Visible = True
        A19_3.Visible = False
        B19_3.Visible = False
        C19_3.Visible = False
        E19_3.Visible = False
    End Sub

    Private Sub E19_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E19_1.Click
        E19_3.Visible = True
        A19_3.Visible = False
        B19_3.Visible = False
        C19_3.Visible = False
        D19_3.Visible = False
    End Sub

    Private Sub A20_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A20_1.Click
        A20_3.Visible = True
        B20_3.Visible = False
        C20_3.Visible = False
        D20_3.Visible = False
        E20_3.Visible = False
    End Sub

    Private Sub B20_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B20_1.Click
        B20_3.Visible = True
        A20_3.Visible = False
        C20_3.Visible = False
        D20_3.Visible = False
        E20_3.Visible = False
    End Sub

    Private Sub C20_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C20_1.Click
        C20_3.Visible = True
        A20_3.Visible = False
        B20_3.Visible = False
        D20_3.Visible = False
        E20_3.Visible = False
    End Sub

    Private Sub D20_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D20_1.Click
        D20_3.Visible = True
        A20_3.Visible = False
        B20_3.Visible = False
        C20_3.Visible = False
        E20_3.Visible = False
    End Sub

    Private Sub E20_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E20_1.Click
        E20_3.Visible = True
        A20_3.Visible = False
        B20_3.Visible = False
        C20_3.Visible = False
        D20_3.Visible = False
    End Sub

    Private Sub A21_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A21_1.Click
        A21_3.Visible = True
        B21_3.Visible = False
        C21_3.Visible = False
        D21_3.Visible = False
        E21_3.Visible = False
    End Sub

    Private Sub B21_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B21_1.Click
        B21_3.Visible = True
        A21_3.Visible = False
        C21_3.Visible = False
        D21_3.Visible = False
        E21_3.Visible = False
    End Sub

    Private Sub C21_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C21_1.Click
        C21_3.Visible = True
        A21_3.Visible = False
        B21_3.Visible = False
        D21_3.Visible = False
        E21_3.Visible = False
    End Sub

    Private Sub D21_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D21_1.Click
        D21_3.Visible = True
        A21_3.Visible = False
        B21_3.Visible = False
        C21_3.Visible = False
        E21_3.Visible = False
    End Sub

    Private Sub E21_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E21_1.Click
        E21_3.Visible = True
        A21_3.Visible = False
        B21_3.Visible = False
        C21_3.Visible = False
        D21_3.Visible = False
    End Sub

    Private Sub A22_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A22_1.Click
        A22_3.Visible = True
        B22_3.Visible = False
        C22_3.Visible = False
        D22_3.Visible = False
        E22_3.Visible = False
    End Sub

    Private Sub B22_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B22_1.Click
        B22_3.Visible = True
        A22_3.Visible = False
        C22_3.Visible = False
        D22_3.Visible = False
        E22_3.Visible = False
    End Sub

    Private Sub C22_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C22_1.Click
        C22_3.Visible = True
        A22_3.Visible = False
        B22_3.Visible = False
        D22_3.Visible = False
        E22_3.Visible = False
    End Sub

    Private Sub D22_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D22_1.Click
        D22_3.Visible = True
        A22_3.Visible = False
        B22_3.Visible = False
        C22_3.Visible = False
        E22_3.Visible = False
    End Sub

    Private Sub E22_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E22_1.Click
        E22_3.Visible = True
        A22_3.Visible = False
        B22_3.Visible = False
        C22_3.Visible = False
        D22_3.Visible = False
    End Sub

    Private Sub A23_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A23_1.Click
        A23_3.Visible = True
        B23_3.Visible = False
        C23_3.Visible = False
        D23_3.Visible = False
        E23_3.Visible = False
    End Sub

    Private Sub B23_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B23_1.Click
        B23_3.Visible = True
        A23_3.Visible = False
        C23_3.Visible = False
        D23_3.Visible = False
        E23_3.Visible = False
    End Sub

    Private Sub C23_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C23_1.Click
        C23_3.Visible = True
        A23_3.Visible = False
        B23_3.Visible = False
        D23_3.Visible = False
        E23_3.Visible = False
    End Sub

    Private Sub D23_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D23_1.Click
        D23_3.Visible = True
        A23_3.Visible = False
        B23_3.Visible = False
        C23_3.Visible = False
        E23_3.Visible = False
    End Sub

    Private Sub E23_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E23_1.Click
        E23_3.Visible = True
        A23_3.Visible = False
        B23_3.Visible = False
        C23_3.Visible = False
        D23_3.Visible = False
    End Sub

    Private Sub A24_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A24_1.Click
        A24_3.Visible = True
        B24_3.Visible = False
        C24_3.Visible = False
        D24_3.Visible = False
        E24_3.Visible = False
    End Sub

    Private Sub B24_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B24_1.Click
        B24_3.Visible = True
        A24_3.Visible = False
        C24_3.Visible = False
        D24_3.Visible = False
        E24_3.Visible = False
    End Sub

    Private Sub C24_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C24_1.Click
        C24_3.Visible = True
        A24_3.Visible = False
        B24_3.Visible = False
        D24_3.Visible = False
        E24_3.Visible = False
    End Sub

    Private Sub D24_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D24_1.Click
        D24_3.Visible = True
        A24_3.Visible = False
        B24_3.Visible = False
        C24_3.Visible = False
        E24_3.Visible = False
    End Sub

    Private Sub E24_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E24_1.Click
        E24_3.Visible = True
        A24_3.Visible = False
        B24_3.Visible = False
        C24_3.Visible = False
        D24_3.Visible = False
    End Sub

    Private Sub A25_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A25_1.Click
        A25_3.Visible = True
        B25_3.Visible = False
        C25_3.Visible = False
        D25_3.Visible = False
        E25_3.Visible = False
    End Sub

    Private Sub B25_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B25_1.Click
        B25_3.Visible = True
        A25_3.Visible = False
        C25_3.Visible = False
        D25_3.Visible = False
        E25_3.Visible = False
    End Sub

    Private Sub C25_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C25_1.Click
        C25_3.Visible = True
        A25_3.Visible = False
        B25_3.Visible = False
        D25_3.Visible = False
        E25_3.Visible = False
    End Sub

    Private Sub D25_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D25_1.Click
        D25_3.Visible = True
        A25_3.Visible = False
        B25_3.Visible = False
        C25_3.Visible = False
        E25_3.Visible = False
    End Sub

    Private Sub E25_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E25_1.Click
        E25_3.Visible = True
        A25_3.Visible = False
        B25_3.Visible = False
        C25_3.Visible = False
        D25_3.Visible = False
    End Sub

    Private Sub A26_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A26_1.Click
        A26_3.Visible = True
        B26_3.Visible = False
        C26_3.Visible = False
        D26_3.Visible = False
        E26_3.Visible = False
    End Sub

    Private Sub B26_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B26_1.Click
        B26_3.Visible = True
        A26_3.Visible = False
        C26_3.Visible = False
        D26_3.Visible = False
        E26_3.Visible = False
    End Sub

    Private Sub C26_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C26_1.Click
        C26_3.Visible = True
        A26_3.Visible = False
        B26_3.Visible = False
        D26_3.Visible = False
        E26_3.Visible = False
    End Sub

    Private Sub D26_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D26_1.Click
        D26_3.Visible = True
        A26_3.Visible = False
        B26_3.Visible = False
        C26_3.Visible = False
        E26_3.Visible = False
    End Sub

    Private Sub E26_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E26_1.Click
        E26_3.Visible = True
        A26_3.Visible = False
        B26_3.Visible = False
        C26_3.Visible = False
        D26_3.Visible = False
    End Sub

    Private Sub A27_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A27_1.Click
        A27_3.Visible = True
        B27_3.Visible = False
        C27_3.Visible = False
        D27_3.Visible = False
        E27_3.Visible = False
    End Sub

    Private Sub B27_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B27_1.Click
        B27_3.Visible = True
        A27_3.Visible = False
        C27_3.Visible = False
        D27_3.Visible = False
        E27_3.Visible = False
    End Sub

    Private Sub C27_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C27_1.Click
        C27_3.Visible = True
        A27_3.Visible = False
        B27_3.Visible = False
        D27_3.Visible = False
        E27_3.Visible = False
    End Sub

    Private Sub D27_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D27_1.Click
        D27_3.Visible = True
        A27_3.Visible = False
        B27_3.Visible = False
        C27_3.Visible = False
        E27_3.Visible = False
    End Sub

    Private Sub E27_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E27_1.Click
        E27_3.Visible = True
        A27_3.Visible = False
        B27_3.Visible = False
        C27_3.Visible = False
        D27_3.Visible = False
    End Sub

    Private Sub A28_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A28_1.Click
        A28_3.Visible = True
        B28_3.Visible = False
        C28_3.Visible = False
        D28_3.Visible = False
        E28_3.Visible = False
    End Sub

    Private Sub B28_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B28_1.Click
        B28_3.Visible = True
        A28_3.Visible = False
        C28_3.Visible = False
        D28_3.Visible = False
        E28_3.Visible = False
    End Sub

    Private Sub C28_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C28_1.Click
        C28_3.Visible = True
        A28_3.Visible = False
        B28_3.Visible = False
        D28_3.Visible = False
        E28_3.Visible = False
    End Sub

    Private Sub D28_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D28_1.Click
        D28_3.Visible = True
        A28_3.Visible = False
        B28_3.Visible = False
        C28_3.Visible = False
        E28_3.Visible = False
    End Sub

    Private Sub E28_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E28_1.Click
        E28_3.Visible = True
        A28_3.Visible = False
        B28_3.Visible = False
        C28_3.Visible = False
        D28_3.Visible = False
    End Sub

    Private Sub A29_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A29_1.Click
        A29_3.Visible = True
        B29_3.Visible = False
        C29_3.Visible = False
        D29_3.Visible = False
        E29_3.Visible = False
    End Sub

    Private Sub B29_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B29_1.Click
        B29_3.Visible = True
        A29_3.Visible = False
        C29_3.Visible = False
        D29_3.Visible = False
        E29_3.Visible = False
    End Sub

    Private Sub C29_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C29_1.Click
        C29_3.Visible = True
        A29_3.Visible = False
        B29_3.Visible = False
        D29_3.Visible = False
        E29_3.Visible = False
    End Sub

    Private Sub D29_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D29_1.Click
        D29_3.Visible = True
        A29_3.Visible = False
        B29_3.Visible = False
        C29_3.Visible = False
        E29_3.Visible = False
    End Sub

    Private Sub E29_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E29_1.Click
        E29_3.Visible = True
        A29_3.Visible = False
        B29_3.Visible = False
        C29_3.Visible = False
        D29_3.Visible = False
    End Sub

    Private Sub A30_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles A30_1.Click
        A30_3.Visible = True
        B30_3.Visible = False
        C30_3.Visible = False
        D30_3.Visible = False
        E30_3.Visible = False
    End Sub

    Private Sub B30_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles B30_1.Click
        B30_3.Visible = True
        A30_3.Visible = False
        C30_3.Visible = False
        D30_3.Visible = False
        E30_3.Visible = False
    End Sub

    Private Sub C30_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles C30_1.Click
        C30_3.Visible = True
        A30_3.Visible = False
        B30_3.Visible = False
        D30_3.Visible = False
        E30_3.Visible = False
    End Sub

    Private Sub D30_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles D30_1.Click
        D30_3.Visible = True
        A30_3.Visible = False
        B30_3.Visible = False
        C30_3.Visible = False
        E30_3.Visible = False
    End Sub

    Private Sub E30_1_Click(ByVal sender As Object, ByVal e As EventArgs) Handles E30_1.Click
        E30_3.Visible = True
        A30_3.Visible = False
        B30_3.Visible = False
        C30_3.Visible = False
        D30_3.Visible = False
    End Sub

    Private Sub PictureBoxFinish_MouseLeave(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxFinish.MouseLeave
        PictureBoxFinish.Image = My.Resources.finish_1
    End Sub

    Private Sub PictureBoxFinish_MouseMove(ByVal sender As Object, ByVal e As MouseEventArgs) Handles PictureBoxFinish.MouseMove
        PictureBoxFinish.Image = My.Resources.finish_2
    End Sub

    Private Function FunctionTests()
        PanelPage1Questions.Visible = True
        PictureBoxPage1QuestionsScroll.Visible = True
        PictureBoxPage1QuestionsScroller.Visible = True

        PanelPage1Questions.VerticalScroll.Value = 0
        Newpoint5.X = PanelPage1.Width - 25
        Newpoint5.Y = 0
        PictureBoxPage1QuestionsScroller.Location = Newpoint5

        PanelMessage.Visible = False

        PictureBoxQuestion01.Location = New System.Drawing.Point(10, 20)
        PanelQuestion01.Location = New System.Drawing.Point(10, PictureBoxQuestion01.Location.Y + PictureBoxQuestion01.Size.Height)
        PictureBoxQuestion02.Location = New System.Drawing.Point(10, 20 + PanelQuestion01.Location.Y + PanelQuestion01.Size.Height)
        PanelQuestion02.Location = New System.Drawing.Point(10, PictureBoxQuestion02.Location.Y + PictureBoxQuestion02.Size.Height)
        PictureBoxQuestion03.Location = New System.Drawing.Point(10, 20 + PanelQuestion02.Location.Y + PanelQuestion02.Size.Height)
        PanelQuestion03.Location = New System.Drawing.Point(10, PictureBoxQuestion03.Location.Y + PictureBoxQuestion03.Size.Height)
        PictureBoxQuestion04.Location = New System.Drawing.Point(10, 20 + PanelQuestion03.Location.Y + PanelQuestion03.Size.Height)
        PanelQuestion04.Location = New System.Drawing.Point(10, PictureBoxQuestion04.Location.Y + PictureBoxQuestion04.Size.Height)
        PictureBoxQuestion05.Location = New System.Drawing.Point(10, 20 + PanelQuestion04.Location.Y + PanelQuestion04.Size.Height)
        PanelQuestion05.Location = New System.Drawing.Point(10, PictureBoxQuestion05.Location.Y + PictureBoxQuestion05.Size.Height)
        PictureBoxQuestion06.Location = New System.Drawing.Point(10, 20 + PanelQuestion05.Location.Y + PanelQuestion05.Size.Height)
        PanelQuestion06.Location = New System.Drawing.Point(10, PictureBoxQuestion06.Location.Y + PictureBoxQuestion06.Size.Height)
        PictureBoxQuestion07.Location = New System.Drawing.Point(10, 20 + PanelQuestion06.Location.Y + PanelQuestion06.Size.Height)
        PanelQuestion07.Location = New System.Drawing.Point(10, PictureBoxQuestion07.Location.Y + PictureBoxQuestion07.Size.Height)
        PictureBoxQuestion08.Location = New System.Drawing.Point(10, 20 + PanelQuestion07.Location.Y + PanelQuestion07.Size.Height)
        PanelQuestion08.Location = New System.Drawing.Point(10, PictureBoxQuestion08.Location.Y + PictureBoxQuestion08.Size.Height)
        PictureBoxQuestion09.Location = New System.Drawing.Point(10, 20 + PanelQuestion08.Location.Y + PanelQuestion08.Size.Height)
        PanelQuestion09.Location = New System.Drawing.Point(10, PictureBoxQuestion09.Location.Y + PictureBoxQuestion09.Size.Height)
        PictureBoxQuestion10.Location = New System.Drawing.Point(10, 20 + PanelQuestion09.Location.Y + PanelQuestion09.Size.Height)
        PanelQuestion10.Location = New System.Drawing.Point(10, PictureBoxQuestion10.Location.Y + PictureBoxQuestion10.Size.Height)
        PictureBoxQuestion11.Location = New System.Drawing.Point(10, 20 + PanelQuestion10.Location.Y + PanelQuestion10.Size.Height)
        PanelQuestion11.Location = New System.Drawing.Point(10, PictureBoxQuestion11.Location.Y + PictureBoxQuestion11.Size.Height)
        PictureBoxQuestion12.Location = New System.Drawing.Point(10, 20 + PanelQuestion11.Location.Y + PanelQuestion11.Size.Height)
        PanelQuestion12.Location = New System.Drawing.Point(10, PictureBoxQuestion12.Location.Y + PictureBoxQuestion12.Size.Height)
        PictureBoxQuestion13.Location = New System.Drawing.Point(10, 20 + PanelQuestion12.Location.Y + PanelQuestion12.Size.Height)
        PanelQuestion13.Location = New System.Drawing.Point(10, PictureBoxQuestion13.Location.Y + PictureBoxQuestion13.Size.Height)
        PictureBoxQuestion14.Location = New System.Drawing.Point(10, 20 + PanelQuestion13.Location.Y + PanelQuestion13.Size.Height)
        PanelQuestion14.Location = New System.Drawing.Point(10, PictureBoxQuestion14.Location.Y + PictureBoxQuestion14.Size.Height)
        PictureBoxQuestion15.Location = New System.Drawing.Point(10, 20 + PanelQuestion14.Location.Y + PanelQuestion14.Size.Height)
        PanelQuestion15.Location = New System.Drawing.Point(10, PictureBoxQuestion15.Location.Y + PictureBoxQuestion15.Size.Height)
        PictureBoxQuestion16.Location = New System.Drawing.Point(10, 20 + PanelQuestion15.Location.Y + PanelQuestion15.Size.Height)
        PanelQuestion16.Location = New System.Drawing.Point(10, PictureBoxQuestion16.Location.Y + PictureBoxQuestion16.Size.Height)
        PictureBoxQuestion17.Location = New System.Drawing.Point(10, 20 + PanelQuestion16.Location.Y + PanelQuestion16.Size.Height)
        PanelQuestion17.Location = New System.Drawing.Point(10, PictureBoxQuestion17.Location.Y + PictureBoxQuestion17.Size.Height)
        PictureBoxQuestion18.Location = New System.Drawing.Point(10, 20 + PanelQuestion17.Location.Y + PanelQuestion17.Size.Height)
        PanelQuestion18.Location = New System.Drawing.Point(10, PictureBoxQuestion18.Location.Y + PictureBoxQuestion18.Size.Height)
        PictureBoxQuestion19.Location = New System.Drawing.Point(10, 20 + PanelQuestion18.Location.Y + PanelQuestion18.Size.Height)
        PanelQuestion19.Location = New System.Drawing.Point(10, PictureBoxQuestion19.Location.Y + PictureBoxQuestion19.Size.Height)
        PictureBoxQuestion20.Location = New System.Drawing.Point(10, 20 + PanelQuestion19.Location.Y + PanelQuestion19.Size.Height)
        PanelQuestion20.Location = New System.Drawing.Point(10, PictureBoxQuestion20.Location.Y + PictureBoxQuestion20.Size.Height)
        PictureBoxQuestion21.Location = New System.Drawing.Point(10, 20 + PanelQuestion20.Location.Y + PanelQuestion20.Size.Height)
        PanelQuestion21.Location = New System.Drawing.Point(10, PictureBoxQuestion21.Location.Y + PictureBoxQuestion21.Size.Height)
        PictureBoxQuestion22.Location = New System.Drawing.Point(10, 20 + PanelQuestion21.Location.Y + PanelQuestion21.Size.Height)
        PanelQuestion22.Location = New System.Drawing.Point(10, PictureBoxQuestion22.Location.Y + PictureBoxQuestion22.Size.Height)
        PictureBoxQuestion23.Location = New System.Drawing.Point(10, 20 + PanelQuestion22.Location.Y + PanelQuestion22.Size.Height)
        PanelQuestion23.Location = New System.Drawing.Point(10, PictureBoxQuestion23.Location.Y + PictureBoxQuestion23.Size.Height)
        PictureBoxQuestion24.Location = New System.Drawing.Point(10, 20 + PanelQuestion23.Location.Y + PanelQuestion23.Size.Height)
        PanelQuestion24.Location = New System.Drawing.Point(10, PictureBoxQuestion24.Location.Y + PictureBoxQuestion24.Size.Height)
        PictureBoxQuestion25.Location = New System.Drawing.Point(10, 20 + PanelQuestion24.Location.Y + PanelQuestion24.Size.Height)
        PanelQuestion25.Location = New System.Drawing.Point(10, PictureBoxQuestion25.Location.Y + PictureBoxQuestion25.Size.Height)
        PictureBoxQuestion26.Location = New System.Drawing.Point(10, 20 + PanelQuestion25.Location.Y + PanelQuestion25.Size.Height)
        PanelQuestion26.Location = New System.Drawing.Point(10, PictureBoxQuestion26.Location.Y + PictureBoxQuestion26.Size.Height)
        PictureBoxQuestion27.Location = New System.Drawing.Point(10, 20 + PanelQuestion26.Location.Y + PanelQuestion26.Size.Height)
        PanelQuestion27.Location = New System.Drawing.Point(10, PictureBoxQuestion27.Location.Y + PictureBoxQuestion27.Size.Height)
        PictureBoxQuestion28.Location = New System.Drawing.Point(10, 20 + PanelQuestion27.Location.Y + PanelQuestion27.Size.Height)
        PanelQuestion28.Location = New System.Drawing.Point(10, PictureBoxQuestion28.Location.Y + PictureBoxQuestion28.Size.Height)
        PictureBoxQuestion29.Location = New System.Drawing.Point(10, 20 + PanelQuestion28.Location.Y + PanelQuestion28.Size.Height)
        PanelQuestion29.Location = New System.Drawing.Point(10, PictureBoxQuestion29.Location.Y + PictureBoxQuestion29.Size.Height)
        PictureBoxQuestion30.Location = New System.Drawing.Point(10, 20 + PanelQuestion29.Location.Y + PanelQuestion29.Size.Height)
        PanelQuestion30.Location = New System.Drawing.Point(10, PictureBoxQuestion30.Location.Y + PictureBoxQuestion30.Size.Height)

        PanelFinish.Location = New System.Drawing.Point(10, 20 + PanelQuestion30.Location.Y + PanelQuestion30.Size.Height)

        PanelPage1Questions.Select()
    End Function

    Function FunctionHider()
        A01_3.Visible = False
        B01_3.Visible = False
        C01_3.Visible = False
        D01_3.Visible = False
        E01_3.Visible = False
        W01.Visible = False
        R01.Visible = False
        A02_3.Visible = False
        B02_3.Visible = False
        C02_3.Visible = False
        D02_3.Visible = False
        E02_3.Visible = False
        W02.Visible = False
        R02.Visible = False
        A03_3.Visible = False
        B03_3.Visible = False
        C03_3.Visible = False
        D03_3.Visible = False
        E03_3.Visible = False
        W03.Visible = False
        R03.Visible = False
        A04_3.Visible = False
        B04_3.Visible = False
        C04_3.Visible = False
        D04_3.Visible = False
        E04_3.Visible = False
        W04.Visible = False
        R04.Visible = False
        A05_3.Visible = False
        B05_3.Visible = False
        C05_3.Visible = False
        D05_3.Visible = False
        E05_3.Visible = False
        W05.Visible = False
        R05.Visible = False
        A06_3.Visible = False
        B06_3.Visible = False
        C06_3.Visible = False
        D06_3.Visible = False
        E06_3.Visible = False
        W06.Visible = False
        R06.Visible = False
        A07_3.Visible = False
        B07_3.Visible = False
        C07_3.Visible = False
        D07_3.Visible = False
        E07_3.Visible = False
        W07.Visible = False
        R07.Visible = False
        A08_3.Visible = False
        B08_3.Visible = False
        C08_3.Visible = False
        D08_3.Visible = False
        E08_3.Visible = False
        W08.Visible = False
        R08.Visible = False
        A09_3.Visible = False
        B09_3.Visible = False
        C09_3.Visible = False
        D09_3.Visible = False
        E09_3.Visible = False
        W09.Visible = False
        R09.Visible = False
        A10_3.Visible = False
        B10_3.Visible = False
        C10_3.Visible = False
        D10_3.Visible = False
        E10_3.Visible = False
        W10.Visible = False
        R10.Visible = False
        A11_3.Visible = False
        B11_3.Visible = False
        C11_3.Visible = False
        D11_3.Visible = False
        E11_3.Visible = False
        W11.Visible = False
        R11.Visible = False
        A12_3.Visible = False
        B12_3.Visible = False
        C12_3.Visible = False
        D12_3.Visible = False
        E12_3.Visible = False
        W12.Visible = False
        R12.Visible = False
        A13_3.Visible = False
        B13_3.Visible = False
        C13_3.Visible = False
        D13_3.Visible = False
        E13_3.Visible = False
        W13.Visible = False
        R13.Visible = False
        A14_3.Visible = False
        B14_3.Visible = False
        C14_3.Visible = False
        D14_3.Visible = False
        E14_3.Visible = False
        W14.Visible = False
        R14.Visible = False
        A15_3.Visible = False
        B15_3.Visible = False
        C15_3.Visible = False
        D15_3.Visible = False
        E15_3.Visible = False
        W15.Visible = False
        R15.Visible = False
        A16_3.Visible = False
        B16_3.Visible = False
        C16_3.Visible = False
        D16_3.Visible = False
        E16_3.Visible = False
        W16.Visible = False
        R16.Visible = False
        A17_3.Visible = False
        B17_3.Visible = False
        C17_3.Visible = False
        D17_3.Visible = False
        E17_3.Visible = False
        W17.Visible = False
        R17.Visible = False
        A18_3.Visible = False
        B18_3.Visible = False
        C18_3.Visible = False
        D18_3.Visible = False
        E18_3.Visible = False
        W18.Visible = False
        R18.Visible = False
        A19_3.Visible = False
        B19_3.Visible = False
        C19_3.Visible = False
        D19_3.Visible = False
        E19_3.Visible = False
        W19.Visible = False
        R19.Visible = False
        A20_3.Visible = False
        B20_3.Visible = False
        C20_3.Visible = False
        D20_3.Visible = False
        E20_3.Visible = False
        W20.Visible = False
        R20.Visible = False
        A21_3.Visible = False
        B21_3.Visible = False
        C21_3.Visible = False
        D21_3.Visible = False
        E21_3.Visible = False
        W21.Visible = False
        R21.Visible = False
        A22_3.Visible = False
        B22_3.Visible = False
        C22_3.Visible = False
        D22_3.Visible = False
        E22_3.Visible = False
        W22.Visible = False
        R22.Visible = False
        A23_3.Visible = False
        B23_3.Visible = False
        C23_3.Visible = False
        D23_3.Visible = False
        E23_3.Visible = False
        W23.Visible = False
        R23.Visible = False
        A24_3.Visible = False
        B24_3.Visible = False
        C24_3.Visible = False
        D24_3.Visible = False
        E24_3.Visible = False
        W24.Visible = False
        R24.Visible = False
        A25_3.Visible = False
        B25_3.Visible = False
        C25_3.Visible = False
        D25_3.Visible = False
        E25_3.Visible = False
        W25.Visible = False
        R25.Visible = False
        A26_3.Visible = False
        B26_3.Visible = False
        C26_3.Visible = False
        D26_3.Visible = False
        E26_3.Visible = False
        W26.Visible = False
        R26.Visible = False
        A27_3.Visible = False
        B27_3.Visible = False
        C27_3.Visible = False
        D27_3.Visible = False
        E27_3.Visible = False
        W27.Visible = False
        R27.Visible = False
        A28_3.Visible = False
        B28_3.Visible = False
        C28_3.Visible = False
        D28_3.Visible = False
        E28_3.Visible = False
        W28.Visible = False
        R28.Visible = False
        A29_3.Visible = False
        B29_3.Visible = False
        C29_3.Visible = False
        D29_3.Visible = False
        E29_3.Visible = False
        W29.Visible = False
        R29.Visible = False
        A30_3.Visible = False
        B30_3.Visible = False
        C30_3.Visible = False
        D30_3.Visible = False
        E30_3.Visible = False
        W30.Visible = False
        R30.Visible = False
    End Function

    Private Sub PictureBoxTest01_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest01_1.Click
        PictureBoxTest01_3.Visible = True
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT01Q01
        PictureBoxQuestion02.Image = My.Resources.PT01Q02
        PictureBoxQuestion03.Image = My.Resources.PT01Q03
        PictureBoxQuestion04.Image = My.Resources.PT01Q04
        PictureBoxQuestion05.Image = My.Resources.PT01Q05
        PictureBoxQuestion06.Image = My.Resources.PT01Q06
        PictureBoxQuestion07.Image = My.Resources.PT01Q07
        PictureBoxQuestion08.Image = My.Resources.PT01Q08
        PictureBoxQuestion09.Image = My.Resources.PT01Q09
        PictureBoxQuestion10.Image = My.Resources.PT01Q10
        PictureBoxQuestion11.Image = My.Resources.PT01Q11
        PictureBoxQuestion12.Image = My.Resources.PT01Q12
        PictureBoxQuestion13.Image = My.Resources.PT01Q13
        PictureBoxQuestion14.Image = My.Resources.PT01Q14
        PictureBoxQuestion15.Image = My.Resources.PT01Q15
        PictureBoxQuestion16.Image = My.Resources.PT01Q16
        PictureBoxQuestion17.Image = My.Resources.PT01Q17
        PictureBoxQuestion18.Image = My.Resources.PT01Q18
        PictureBoxQuestion19.Image = My.Resources.PT01Q19
        PictureBoxQuestion20.Image = My.Resources.PT01Q20
        PictureBoxQuestion21.Image = My.Resources.PT01Q21
        PictureBoxQuestion22.Image = My.Resources.PT01Q22
        PictureBoxQuestion23.Image = My.Resources.PT01Q23
        PictureBoxQuestion24.Image = My.Resources.PT01Q24
        PictureBoxQuestion25.Image = My.Resources.PT01Q25
        PictureBoxQuestion26.Image = My.Resources.PT01Q26
        PictureBoxQuestion27.Image = My.Resources.PT01Q27
        PictureBoxQuestion28.Image = My.Resources.PT01Q28
        PictureBoxQuestion29.Image = My.Resources.PT01Q29
        PictureBoxQuestion30.Image = My.Resources.PT01Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest02_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest02_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = True
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT02Q01
        PictureBoxQuestion02.Image = My.Resources.PT02Q02
        PictureBoxQuestion03.Image = My.Resources.PT02Q03
        PictureBoxQuestion04.Image = My.Resources.PT02Q04
        PictureBoxQuestion05.Image = My.Resources.PT02Q05
        PictureBoxQuestion06.Image = My.Resources.PT02Q06
        PictureBoxQuestion07.Image = My.Resources.PT02Q07
        PictureBoxQuestion08.Image = My.Resources.PT02Q08
        PictureBoxQuestion09.Image = My.Resources.PT02Q09
        PictureBoxQuestion10.Image = My.Resources.PT02Q10
        PictureBoxQuestion11.Image = My.Resources.PT02Q11
        PictureBoxQuestion12.Image = My.Resources.PT02Q12
        PictureBoxQuestion13.Image = My.Resources.PT02Q13
        PictureBoxQuestion14.Image = My.Resources.PT02Q14
        PictureBoxQuestion15.Image = My.Resources.PT02Q15
        PictureBoxQuestion16.Image = My.Resources.PT02Q16
        PictureBoxQuestion17.Image = My.Resources.PT02Q17
        PictureBoxQuestion18.Image = My.Resources.PT02Q18
        PictureBoxQuestion19.Image = My.Resources.PT02Q19
        PictureBoxQuestion20.Image = My.Resources.PT02Q20
        PictureBoxQuestion21.Image = My.Resources.PT02Q21
        PictureBoxQuestion22.Image = My.Resources.PT02Q22
        PictureBoxQuestion23.Image = My.Resources.PT02Q23
        PictureBoxQuestion24.Image = My.Resources.PT02Q24
        PictureBoxQuestion25.Image = My.Resources.PT02Q25
        PictureBoxQuestion26.Image = My.Resources.PT02Q26
        PictureBoxQuestion27.Image = My.Resources.PT02Q27
        PictureBoxQuestion28.Image = My.Resources.PT02Q28
        PictureBoxQuestion29.Image = My.Resources.PT02Q29
        PictureBoxQuestion30.Image = My.Resources.PT02Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest03_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest03_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = True
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT03Q01
        PictureBoxQuestion02.Image = My.Resources.PT03Q02
        PictureBoxQuestion03.Image = My.Resources.PT03Q03
        PictureBoxQuestion04.Image = My.Resources.PT03Q04
        PictureBoxQuestion05.Image = My.Resources.PT03Q05
        PictureBoxQuestion06.Image = My.Resources.PT03Q06
        PictureBoxQuestion07.Image = My.Resources.PT03Q07
        PictureBoxQuestion08.Image = My.Resources.PT03Q08
        PictureBoxQuestion09.Image = My.Resources.PT03Q09
        PictureBoxQuestion10.Image = My.Resources.PT03Q10
        PictureBoxQuestion11.Image = My.Resources.PT03Q11
        PictureBoxQuestion12.Image = My.Resources.PT03Q12
        PictureBoxQuestion13.Image = My.Resources.PT03Q13
        PictureBoxQuestion14.Image = My.Resources.PT03Q14
        PictureBoxQuestion15.Image = My.Resources.PT03Q15
        PictureBoxQuestion16.Image = My.Resources.PT03Q16
        PictureBoxQuestion17.Image = My.Resources.PT03Q17
        PictureBoxQuestion18.Image = My.Resources.PT03Q18
        PictureBoxQuestion19.Image = My.Resources.PT03Q19
        PictureBoxQuestion20.Image = My.Resources.PT03Q20
        PictureBoxQuestion21.Image = My.Resources.PT03Q21
        PictureBoxQuestion22.Image = My.Resources.PT03Q22
        PictureBoxQuestion23.Image = My.Resources.PT03Q23
        PictureBoxQuestion24.Image = My.Resources.PT03Q24
        PictureBoxQuestion25.Image = My.Resources.PT03Q25
        PictureBoxQuestion26.Image = My.Resources.PT03Q26
        PictureBoxQuestion27.Image = My.Resources.PT03Q27
        PictureBoxQuestion28.Image = My.Resources.PT03Q28
        PictureBoxQuestion29.Image = My.Resources.PT03Q29
        PictureBoxQuestion30.Image = My.Resources.PT03Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest04_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest04_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = True
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT04Q01
        PictureBoxQuestion02.Image = My.Resources.PT04Q02
        PictureBoxQuestion03.Image = My.Resources.PT04Q03
        PictureBoxQuestion04.Image = My.Resources.PT04Q04
        PictureBoxQuestion05.Image = My.Resources.PT04Q05
        PictureBoxQuestion06.Image = My.Resources.PT04Q06
        PictureBoxQuestion07.Image = My.Resources.PT04Q07
        PictureBoxQuestion08.Image = My.Resources.PT04Q08
        PictureBoxQuestion09.Image = My.Resources.PT04Q09
        PictureBoxQuestion10.Image = My.Resources.PT04Q10
        PictureBoxQuestion11.Image = My.Resources.PT04Q11
        PictureBoxQuestion12.Image = My.Resources.PT04Q12
        PictureBoxQuestion13.Image = My.Resources.PT04Q13
        PictureBoxQuestion14.Image = My.Resources.PT04Q14
        PictureBoxQuestion15.Image = My.Resources.PT04Q15
        PictureBoxQuestion16.Image = My.Resources.PT04Q16
        PictureBoxQuestion17.Image = My.Resources.PT04Q17
        PictureBoxQuestion18.Image = My.Resources.PT04Q18
        PictureBoxQuestion19.Image = My.Resources.PT04Q19
        PictureBoxQuestion20.Image = My.Resources.PT04Q20
        PictureBoxQuestion21.Image = My.Resources.PT04Q21
        PictureBoxQuestion22.Image = My.Resources.PT04Q22
        PictureBoxQuestion23.Image = My.Resources.PT04Q23
        PictureBoxQuestion24.Image = My.Resources.PT04Q24
        PictureBoxQuestion25.Image = My.Resources.PT04Q25
        PictureBoxQuestion26.Image = My.Resources.PT04Q26
        PictureBoxQuestion27.Image = My.Resources.PT04Q27
        PictureBoxQuestion28.Image = My.Resources.PT04Q28
        PictureBoxQuestion29.Image = My.Resources.PT04Q29
        PictureBoxQuestion30.Image = My.Resources.PT04Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest05_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest05_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = True
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT05Q01
        PictureBoxQuestion02.Image = My.Resources.PT05Q02
        PictureBoxQuestion03.Image = My.Resources.PT05Q03
        PictureBoxQuestion04.Image = My.Resources.PT05Q04
        PictureBoxQuestion05.Image = My.Resources.PT05Q05
        PictureBoxQuestion06.Image = My.Resources.PT05Q06
        PictureBoxQuestion07.Image = My.Resources.PT05Q07
        PictureBoxQuestion08.Image = My.Resources.PT05Q08
        PictureBoxQuestion09.Image = My.Resources.PT05Q09
        PictureBoxQuestion10.Image = My.Resources.PT05Q10
        PictureBoxQuestion11.Image = My.Resources.PT05Q11
        PictureBoxQuestion12.Image = My.Resources.PT05Q12
        PictureBoxQuestion13.Image = My.Resources.PT05Q13
        PictureBoxQuestion14.Image = My.Resources.PT05Q14
        PictureBoxQuestion15.Image = My.Resources.PT05Q15
        PictureBoxQuestion16.Image = My.Resources.PT05Q16
        PictureBoxQuestion17.Image = My.Resources.PT05Q17
        PictureBoxQuestion18.Image = My.Resources.PT05Q18
        PictureBoxQuestion19.Image = My.Resources.PT05Q19
        PictureBoxQuestion20.Image = My.Resources.PT05Q20
        PictureBoxQuestion21.Image = My.Resources.PT05Q21
        PictureBoxQuestion22.Image = My.Resources.PT05Q22
        PictureBoxQuestion23.Image = My.Resources.PT05Q23
        PictureBoxQuestion24.Image = My.Resources.PT05Q24
        PictureBoxQuestion25.Image = My.Resources.PT05Q25
        PictureBoxQuestion26.Image = My.Resources.PT05Q26
        PictureBoxQuestion27.Image = My.Resources.PT05Q27
        PictureBoxQuestion28.Image = My.Resources.PT05Q28
        PictureBoxQuestion29.Image = My.Resources.PT05Q29
        PictureBoxQuestion30.Image = My.Resources.PT05Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest06_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest06_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = True
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT06Q01
        PictureBoxQuestion02.Image = My.Resources.PT06Q02
        PictureBoxQuestion03.Image = My.Resources.PT06Q03
        PictureBoxQuestion04.Image = My.Resources.PT06Q04
        PictureBoxQuestion05.Image = My.Resources.PT06Q05
        PictureBoxQuestion06.Image = My.Resources.PT06Q06
        PictureBoxQuestion07.Image = My.Resources.PT06Q07
        PictureBoxQuestion08.Image = My.Resources.PT06Q08
        PictureBoxQuestion09.Image = My.Resources.PT06Q09
        PictureBoxQuestion10.Image = My.Resources.PT06Q10
        PictureBoxQuestion11.Image = My.Resources.PT06Q11
        PictureBoxQuestion12.Image = My.Resources.PT06Q12
        PictureBoxQuestion13.Image = My.Resources.PT06Q13
        PictureBoxQuestion14.Image = My.Resources.PT06Q14
        PictureBoxQuestion15.Image = My.Resources.PT06Q15
        PictureBoxQuestion16.Image = My.Resources.PT06Q16
        PictureBoxQuestion17.Image = My.Resources.PT06Q17
        PictureBoxQuestion18.Image = My.Resources.PT06Q18
        PictureBoxQuestion19.Image = My.Resources.PT06Q19
        PictureBoxQuestion20.Image = My.Resources.PT06Q20
        PictureBoxQuestion21.Image = My.Resources.PT06Q21
        PictureBoxQuestion22.Image = My.Resources.PT06Q22
        PictureBoxQuestion23.Image = My.Resources.PT06Q23
        PictureBoxQuestion24.Image = My.Resources.PT06Q24
        PictureBoxQuestion25.Image = My.Resources.PT06Q25
        PictureBoxQuestion26.Image = My.Resources.PT06Q26
        PictureBoxQuestion27.Image = My.Resources.PT06Q27
        PictureBoxQuestion28.Image = My.Resources.PT06Q28
        PictureBoxQuestion29.Image = My.Resources.PT06Q29
        PictureBoxQuestion30.Image = My.Resources.PT06Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest07_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest07_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = True
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT07Q01
        PictureBoxQuestion02.Image = My.Resources.PT07Q02
        PictureBoxQuestion03.Image = My.Resources.PT07Q03
        PictureBoxQuestion04.Image = My.Resources.PT07Q04
        PictureBoxQuestion05.Image = My.Resources.PT07Q05
        PictureBoxQuestion06.Image = My.Resources.PT07Q06
        PictureBoxQuestion07.Image = My.Resources.PT07Q07
        PictureBoxQuestion08.Image = My.Resources.PT07Q08
        PictureBoxQuestion09.Image = My.Resources.PT07Q09
        PictureBoxQuestion10.Image = My.Resources.PT07Q10
        PictureBoxQuestion11.Image = My.Resources.PT07Q11
        PictureBoxQuestion12.Image = My.Resources.PT07Q12
        PictureBoxQuestion13.Image = My.Resources.PT07Q13
        PictureBoxQuestion14.Image = My.Resources.PT07Q14
        PictureBoxQuestion15.Image = My.Resources.PT07Q15
        PictureBoxQuestion16.Image = My.Resources.PT07Q16
        PictureBoxQuestion17.Image = My.Resources.PT07Q17
        PictureBoxQuestion18.Image = My.Resources.PT07Q18
        PictureBoxQuestion19.Image = My.Resources.PT07Q19
        PictureBoxQuestion20.Image = My.Resources.PT07Q20
        PictureBoxQuestion21.Image = My.Resources.PT07Q21
        PictureBoxQuestion22.Image = My.Resources.PT07Q22
        PictureBoxQuestion23.Image = My.Resources.PT07Q23
        PictureBoxQuestion24.Image = My.Resources.PT07Q24
        PictureBoxQuestion25.Image = My.Resources.PT07Q25
        PictureBoxQuestion26.Image = My.Resources.PT07Q26
        PictureBoxQuestion27.Image = My.Resources.PT07Q27
        PictureBoxQuestion28.Image = My.Resources.PT07Q28
        PictureBoxQuestion29.Image = My.Resources.PT07Q29
        PictureBoxQuestion30.Image = My.Resources.PT07Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest08_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest08_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = True
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT08Q01
        PictureBoxQuestion02.Image = My.Resources.PT08Q02
        PictureBoxQuestion03.Image = My.Resources.PT08Q03
        PictureBoxQuestion04.Image = My.Resources.PT08Q04
        PictureBoxQuestion05.Image = My.Resources.PT08Q05
        PictureBoxQuestion06.Image = My.Resources.PT08Q06
        PictureBoxQuestion07.Image = My.Resources.PT08Q07
        PictureBoxQuestion08.Image = My.Resources.PT08Q08
        PictureBoxQuestion09.Image = My.Resources.PT08Q09
        PictureBoxQuestion10.Image = My.Resources.PT08Q10
        PictureBoxQuestion11.Image = My.Resources.PT08Q11
        PictureBoxQuestion12.Image = My.Resources.PT08Q12
        PictureBoxQuestion13.Image = My.Resources.PT08Q13
        PictureBoxQuestion14.Image = My.Resources.PT08Q14
        PictureBoxQuestion15.Image = My.Resources.PT08Q15
        PictureBoxQuestion16.Image = My.Resources.PT08Q16
        PictureBoxQuestion17.Image = My.Resources.PT08Q17
        PictureBoxQuestion18.Image = My.Resources.PT08Q18
        PictureBoxQuestion19.Image = My.Resources.PT08Q19
        PictureBoxQuestion20.Image = My.Resources.PT08Q20
        PictureBoxQuestion21.Image = My.Resources.PT08Q21
        PictureBoxQuestion22.Image = My.Resources.PT08Q22
        PictureBoxQuestion23.Image = My.Resources.PT08Q23
        PictureBoxQuestion24.Image = My.Resources.PT08Q24
        PictureBoxQuestion25.Image = My.Resources.PT08Q25
        PictureBoxQuestion26.Image = My.Resources.PT08Q26
        PictureBoxQuestion27.Image = My.Resources.PT08Q27
        PictureBoxQuestion28.Image = My.Resources.PT08Q28
        PictureBoxQuestion29.Image = My.Resources.PT08Q29
        PictureBoxQuestion30.Image = My.Resources.PT08Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest09_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest09_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = True
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT09Q01
        PictureBoxQuestion02.Image = My.Resources.PT09Q02
        PictureBoxQuestion03.Image = My.Resources.PT09Q03
        PictureBoxQuestion04.Image = My.Resources.PT09Q04
        PictureBoxQuestion05.Image = My.Resources.PT09Q05
        PictureBoxQuestion06.Image = My.Resources.PT09Q06
        PictureBoxQuestion07.Image = My.Resources.PT09Q07
        PictureBoxQuestion08.Image = My.Resources.PT09Q08
        PictureBoxQuestion09.Image = My.Resources.PT09Q09
        PictureBoxQuestion10.Image = My.Resources.PT09Q10
        PictureBoxQuestion11.Image = My.Resources.PT09Q11
        PictureBoxQuestion12.Image = My.Resources.PT09Q12
        PictureBoxQuestion13.Image = My.Resources.PT09Q13
        PictureBoxQuestion14.Image = My.Resources.PT09Q14
        PictureBoxQuestion15.Image = My.Resources.PT09Q15
        PictureBoxQuestion16.Image = My.Resources.PT09Q16
        PictureBoxQuestion17.Image = My.Resources.PT09Q17
        PictureBoxQuestion18.Image = My.Resources.PT09Q18
        PictureBoxQuestion19.Image = My.Resources.PT09Q19
        PictureBoxQuestion20.Image = My.Resources.PT09Q20
        PictureBoxQuestion21.Image = My.Resources.PT09Q21
        PictureBoxQuestion22.Image = My.Resources.PT09Q22
        PictureBoxQuestion23.Image = My.Resources.PT09Q23
        PictureBoxQuestion24.Image = My.Resources.PT09Q24
        PictureBoxQuestion25.Image = My.Resources.PT09Q25
        PictureBoxQuestion26.Image = My.Resources.PT09Q26
        PictureBoxQuestion27.Image = My.Resources.PT09Q27
        PictureBoxQuestion28.Image = My.Resources.PT09Q28
        PictureBoxQuestion29.Image = My.Resources.PT09Q29
        PictureBoxQuestion30.Image = My.Resources.PT09Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest10_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest10_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = True
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT10Q01
        PictureBoxQuestion02.Image = My.Resources.PT10Q02
        PictureBoxQuestion03.Image = My.Resources.PT10Q03
        PictureBoxQuestion04.Image = My.Resources.PT10Q04
        PictureBoxQuestion05.Image = My.Resources.PT10Q05
        PictureBoxQuestion06.Image = My.Resources.PT10Q06
        PictureBoxQuestion07.Image = My.Resources.PT10Q07
        PictureBoxQuestion08.Image = My.Resources.PT10Q08
        PictureBoxQuestion09.Image = My.Resources.PT10Q09
        PictureBoxQuestion10.Image = My.Resources.PT10Q10
        PictureBoxQuestion11.Image = My.Resources.PT10Q11
        PictureBoxQuestion12.Image = My.Resources.PT10Q12
        PictureBoxQuestion13.Image = My.Resources.PT10Q13
        PictureBoxQuestion14.Image = My.Resources.PT10Q14
        PictureBoxQuestion15.Image = My.Resources.PT10Q15
        PictureBoxQuestion16.Image = My.Resources.PT10Q16
        PictureBoxQuestion17.Image = My.Resources.PT10Q17
        PictureBoxQuestion18.Image = My.Resources.PT10Q18
        PictureBoxQuestion19.Image = My.Resources.PT10Q19
        PictureBoxQuestion20.Image = My.Resources.PT10Q20
        PictureBoxQuestion21.Image = My.Resources.PT10Q21
        PictureBoxQuestion22.Image = My.Resources.PT10Q22
        PictureBoxQuestion23.Image = My.Resources.PT10Q23
        PictureBoxQuestion24.Image = My.Resources.PT10Q24
        PictureBoxQuestion25.Image = My.Resources.PT10Q25
        PictureBoxQuestion26.Image = My.Resources.PT10Q26
        PictureBoxQuestion27.Image = My.Resources.PT10Q27
        PictureBoxQuestion28.Image = My.Resources.PT10Q28
        PictureBoxQuestion29.Image = My.Resources.PT10Q29
        PictureBoxQuestion30.Image = My.Resources.PT10Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest11_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest11_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = True
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT11Q01
        PictureBoxQuestion02.Image = My.Resources.PT11Q02
        PictureBoxQuestion03.Image = My.Resources.PT11Q03
        PictureBoxQuestion04.Image = My.Resources.PT11Q04
        PictureBoxQuestion05.Image = My.Resources.PT11Q05
        PictureBoxQuestion06.Image = My.Resources.PT11Q06
        PictureBoxQuestion07.Image = My.Resources.PT11Q07
        PictureBoxQuestion08.Image = My.Resources.PT11Q08
        PictureBoxQuestion09.Image = My.Resources.PT11Q09
        PictureBoxQuestion10.Image = My.Resources.PT11Q10
        PictureBoxQuestion11.Image = My.Resources.PT11Q11
        PictureBoxQuestion12.Image = My.Resources.PT11Q12
        PictureBoxQuestion13.Image = My.Resources.PT11Q13
        PictureBoxQuestion14.Image = My.Resources.PT11Q14
        PictureBoxQuestion15.Image = My.Resources.PT11Q15
        PictureBoxQuestion16.Image = My.Resources.PT11Q16
        PictureBoxQuestion17.Image = My.Resources.PT11Q17
        PictureBoxQuestion18.Image = My.Resources.PT11Q18
        PictureBoxQuestion19.Image = My.Resources.PT11Q19
        PictureBoxQuestion20.Image = My.Resources.PT11Q20
        PictureBoxQuestion21.Image = My.Resources.PT11Q21
        PictureBoxQuestion22.Image = My.Resources.PT11Q22
        PictureBoxQuestion23.Image = My.Resources.PT11Q23
        PictureBoxQuestion24.Image = My.Resources.PT11Q24
        PictureBoxQuestion25.Image = My.Resources.PT11Q25
        PictureBoxQuestion26.Image = My.Resources.PT11Q26
        PictureBoxQuestion27.Image = My.Resources.PT11Q27
        PictureBoxQuestion28.Image = My.Resources.PT11Q28
        PictureBoxQuestion29.Image = My.Resources.PT11Q29
        PictureBoxQuestion30.Image = My.Resources.PT11Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest12_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest12_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = True
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT12Q01
        PictureBoxQuestion02.Image = My.Resources.PT12Q02
        PictureBoxQuestion03.Image = My.Resources.PT12Q03
        PictureBoxQuestion04.Image = My.Resources.PT12Q04
        PictureBoxQuestion05.Image = My.Resources.PT12Q05
        PictureBoxQuestion06.Image = My.Resources.PT12Q06
        PictureBoxQuestion07.Image = My.Resources.PT12Q07
        PictureBoxQuestion08.Image = My.Resources.PT12Q08
        PictureBoxQuestion09.Image = My.Resources.PT12Q09
        PictureBoxQuestion10.Image = My.Resources.PT12Q10
        PictureBoxQuestion11.Image = My.Resources.PT12Q11
        PictureBoxQuestion12.Image = My.Resources.PT12Q12
        PictureBoxQuestion13.Image = My.Resources.PT12Q13
        PictureBoxQuestion14.Image = My.Resources.PT12Q14
        PictureBoxQuestion15.Image = My.Resources.PT12Q15
        PictureBoxQuestion16.Image = My.Resources.PT12Q16
        PictureBoxQuestion17.Image = My.Resources.PT12Q17
        PictureBoxQuestion18.Image = My.Resources.PT12Q18
        PictureBoxQuestion19.Image = My.Resources.PT12Q19
        PictureBoxQuestion20.Image = My.Resources.PT12Q20
        PictureBoxQuestion21.Image = My.Resources.PT12Q21
        PictureBoxQuestion22.Image = My.Resources.PT12Q22
        PictureBoxQuestion23.Image = My.Resources.PT12Q23
        PictureBoxQuestion24.Image = My.Resources.PT12Q24
        PictureBoxQuestion25.Image = My.Resources.PT12Q25
        PictureBoxQuestion26.Image = My.Resources.PT12Q26
        PictureBoxQuestion27.Image = My.Resources.PT12Q27
        PictureBoxQuestion28.Image = My.Resources.PT12Q28
        PictureBoxQuestion29.Image = My.Resources.PT12Q29
        PictureBoxQuestion30.Image = My.Resources.PT12Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest13_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest13_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = True
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT13Q01
        PictureBoxQuestion02.Image = My.Resources.PT13Q02
        PictureBoxQuestion03.Image = My.Resources.PT13Q03
        PictureBoxQuestion04.Image = My.Resources.PT13Q04
        PictureBoxQuestion05.Image = My.Resources.PT13Q05
        PictureBoxQuestion06.Image = My.Resources.PT13Q06
        PictureBoxQuestion07.Image = My.Resources.PT13Q07
        PictureBoxQuestion08.Image = My.Resources.PT13Q08
        PictureBoxQuestion09.Image = My.Resources.PT13Q09
        PictureBoxQuestion10.Image = My.Resources.PT13Q10
        PictureBoxQuestion11.Image = My.Resources.PT13Q11
        PictureBoxQuestion12.Image = My.Resources.PT13Q12
        PictureBoxQuestion13.Image = My.Resources.PT13Q13
        PictureBoxQuestion14.Image = My.Resources.PT13Q14
        PictureBoxQuestion15.Image = My.Resources.PT13Q15
        PictureBoxQuestion16.Image = My.Resources.PT13Q16
        PictureBoxQuestion17.Image = My.Resources.PT13Q17
        PictureBoxQuestion18.Image = My.Resources.PT13Q18
        PictureBoxQuestion19.Image = My.Resources.PT13Q19
        PictureBoxQuestion20.Image = My.Resources.PT13Q20
        PictureBoxQuestion21.Image = My.Resources.PT13Q21
        PictureBoxQuestion22.Image = My.Resources.PT13Q22
        PictureBoxQuestion23.Image = My.Resources.PT13Q23
        PictureBoxQuestion24.Image = My.Resources.PT13Q24
        PictureBoxQuestion25.Image = My.Resources.PT13Q25
        PictureBoxQuestion26.Image = My.Resources.PT13Q26
        PictureBoxQuestion27.Image = My.Resources.PT13Q27
        PictureBoxQuestion28.Image = My.Resources.PT13Q28
        PictureBoxQuestion29.Image = My.Resources.PT13Q29
        PictureBoxQuestion30.Image = My.Resources.PT13Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest14_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest14_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = True
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT14Q01
        PictureBoxQuestion02.Image = My.Resources.PT14Q02
        PictureBoxQuestion03.Image = My.Resources.PT14Q03
        PictureBoxQuestion04.Image = My.Resources.PT14Q04
        PictureBoxQuestion05.Image = My.Resources.PT14Q05
        PictureBoxQuestion06.Image = My.Resources.PT14Q06
        PictureBoxQuestion07.Image = My.Resources.PT14Q07
        PictureBoxQuestion08.Image = My.Resources.PT14Q08
        PictureBoxQuestion09.Image = My.Resources.PT14Q09
        PictureBoxQuestion10.Image = My.Resources.PT14Q10
        PictureBoxQuestion11.Image = My.Resources.PT14Q11
        PictureBoxQuestion12.Image = My.Resources.PT14Q12
        PictureBoxQuestion13.Image = My.Resources.PT14Q13
        PictureBoxQuestion14.Image = My.Resources.PT14Q14
        PictureBoxQuestion15.Image = My.Resources.PT14Q15
        PictureBoxQuestion16.Image = My.Resources.PT14Q16
        PictureBoxQuestion17.Image = My.Resources.PT14Q17
        PictureBoxQuestion18.Image = My.Resources.PT14Q18
        PictureBoxQuestion19.Image = My.Resources.PT14Q19
        PictureBoxQuestion20.Image = My.Resources.PT14Q20
        PictureBoxQuestion21.Image = My.Resources.PT14Q21
        PictureBoxQuestion22.Image = My.Resources.PT14Q22
        PictureBoxQuestion23.Image = My.Resources.PT14Q23
        PictureBoxQuestion24.Image = My.Resources.PT14Q24
        PictureBoxQuestion25.Image = My.Resources.PT14Q25
        PictureBoxQuestion26.Image = My.Resources.PT14Q26
        PictureBoxQuestion27.Image = My.Resources.PT14Q27
        PictureBoxQuestion28.Image = My.Resources.PT14Q28
        PictureBoxQuestion29.Image = My.Resources.PT14Q29
        PictureBoxQuestion30.Image = My.Resources.PT14Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest15_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest15_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = True
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT15Q01
        PictureBoxQuestion02.Image = My.Resources.PT15Q02
        PictureBoxQuestion03.Image = My.Resources.PT15Q03
        PictureBoxQuestion04.Image = My.Resources.PT15Q04
        PictureBoxQuestion05.Image = My.Resources.PT15Q05
        PictureBoxQuestion06.Image = My.Resources.PT15Q06
        PictureBoxQuestion07.Image = My.Resources.PT15Q07
        PictureBoxQuestion08.Image = My.Resources.PT15Q08
        PictureBoxQuestion09.Image = My.Resources.PT15Q09
        PictureBoxQuestion10.Image = My.Resources.PT15Q10
        PictureBoxQuestion11.Image = My.Resources.PT15Q11
        PictureBoxQuestion12.Image = My.Resources.PT15Q12
        PictureBoxQuestion13.Image = My.Resources.PT15Q13
        PictureBoxQuestion14.Image = My.Resources.PT15Q14
        PictureBoxQuestion15.Image = My.Resources.PT15Q15
        PictureBoxQuestion16.Image = My.Resources.PT15Q16
        PictureBoxQuestion17.Image = My.Resources.PT15Q17
        PictureBoxQuestion18.Image = My.Resources.PT15Q18
        PictureBoxQuestion19.Image = My.Resources.PT15Q19
        PictureBoxQuestion20.Image = My.Resources.PT15Q20
        PictureBoxQuestion21.Image = My.Resources.PT15Q21
        PictureBoxQuestion22.Image = My.Resources.PT15Q22
        PictureBoxQuestion23.Image = My.Resources.PT15Q23
        PictureBoxQuestion24.Image = My.Resources.PT15Q24
        PictureBoxQuestion25.Image = My.Resources.PT15Q25
        PictureBoxQuestion26.Image = My.Resources.PT15Q26
        PictureBoxQuestion27.Image = My.Resources.PT15Q27
        PictureBoxQuestion28.Image = My.Resources.PT15Q28
        PictureBoxQuestion29.Image = My.Resources.PT15Q29
        PictureBoxQuestion30.Image = My.Resources.PT15Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest16_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest16_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = True
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT16Q01
        PictureBoxQuestion02.Image = My.Resources.PT16Q02
        PictureBoxQuestion03.Image = My.Resources.PT16Q03
        PictureBoxQuestion04.Image = My.Resources.PT16Q04
        PictureBoxQuestion05.Image = My.Resources.PT16Q05
        PictureBoxQuestion06.Image = My.Resources.PT16Q06
        PictureBoxQuestion07.Image = My.Resources.PT16Q07
        PictureBoxQuestion08.Image = My.Resources.PT16Q08
        PictureBoxQuestion09.Image = My.Resources.PT16Q09
        PictureBoxQuestion10.Image = My.Resources.PT16Q10
        PictureBoxQuestion11.Image = My.Resources.PT16Q11
        PictureBoxQuestion12.Image = My.Resources.PT16Q12
        PictureBoxQuestion13.Image = My.Resources.PT16Q13
        PictureBoxQuestion14.Image = My.Resources.PT16Q14
        PictureBoxQuestion15.Image = My.Resources.PT16Q15
        PictureBoxQuestion16.Image = My.Resources.PT16Q16
        PictureBoxQuestion17.Image = My.Resources.PT16Q17
        PictureBoxQuestion18.Image = My.Resources.PT16Q18
        PictureBoxQuestion19.Image = My.Resources.PT16Q19
        PictureBoxQuestion20.Image = My.Resources.PT16Q20
        PictureBoxQuestion21.Image = My.Resources.PT16Q21
        PictureBoxQuestion22.Image = My.Resources.PT16Q22
        PictureBoxQuestion23.Image = My.Resources.PT16Q23
        PictureBoxQuestion24.Image = My.Resources.PT16Q24
        PictureBoxQuestion25.Image = My.Resources.PT16Q25
        PictureBoxQuestion26.Image = My.Resources.PT16Q26
        PictureBoxQuestion27.Image = My.Resources.PT16Q27
        PictureBoxQuestion28.Image = My.Resources.PT16Q28
        PictureBoxQuestion29.Image = My.Resources.PT16Q29
        PictureBoxQuestion30.Image = My.Resources.PT16Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest17_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest17_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = True
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT17Q01
        PictureBoxQuestion02.Image = My.Resources.PT17Q02
        PictureBoxQuestion03.Image = My.Resources.PT17Q03
        PictureBoxQuestion04.Image = My.Resources.PT17Q04
        PictureBoxQuestion05.Image = My.Resources.PT17Q05
        PictureBoxQuestion06.Image = My.Resources.PT17Q06
        PictureBoxQuestion07.Image = My.Resources.PT17Q07
        PictureBoxQuestion08.Image = My.Resources.PT17Q08
        PictureBoxQuestion09.Image = My.Resources.PT17Q09
        PictureBoxQuestion10.Image = My.Resources.PT17Q10
        PictureBoxQuestion11.Image = My.Resources.PT17Q11
        PictureBoxQuestion12.Image = My.Resources.PT17Q12
        PictureBoxQuestion13.Image = My.Resources.PT17Q13
        PictureBoxQuestion14.Image = My.Resources.PT17Q14
        PictureBoxQuestion15.Image = My.Resources.PT17Q15
        PictureBoxQuestion16.Image = My.Resources.PT17Q16
        PictureBoxQuestion17.Image = My.Resources.PT17Q17
        PictureBoxQuestion18.Image = My.Resources.PT17Q18
        PictureBoxQuestion19.Image = My.Resources.PT17Q19
        PictureBoxQuestion20.Image = My.Resources.PT17Q20
        PictureBoxQuestion21.Image = My.Resources.PT17Q21
        PictureBoxQuestion22.Image = My.Resources.PT17Q22
        PictureBoxQuestion23.Image = My.Resources.PT17Q23
        PictureBoxQuestion24.Image = My.Resources.PT17Q24
        PictureBoxQuestion25.Image = My.Resources.PT17Q25
        PictureBoxQuestion26.Image = My.Resources.PT17Q26
        PictureBoxQuestion27.Image = My.Resources.PT17Q27
        PictureBoxQuestion28.Image = My.Resources.PT17Q28
        PictureBoxQuestion29.Image = My.Resources.PT17Q29
        PictureBoxQuestion30.Image = My.Resources.PT17Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest18_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest18_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = True
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT18Q01
        PictureBoxQuestion02.Image = My.Resources.PT18Q02
        PictureBoxQuestion03.Image = My.Resources.PT18Q03
        PictureBoxQuestion04.Image = My.Resources.PT18Q04
        PictureBoxQuestion05.Image = My.Resources.PT18Q05
        PictureBoxQuestion06.Image = My.Resources.PT18Q06
        PictureBoxQuestion07.Image = My.Resources.PT18Q07
        PictureBoxQuestion08.Image = My.Resources.PT18Q08
        PictureBoxQuestion09.Image = My.Resources.PT18Q09
        PictureBoxQuestion10.Image = My.Resources.PT18Q10
        PictureBoxQuestion11.Image = My.Resources.PT18Q11
        PictureBoxQuestion12.Image = My.Resources.PT18Q12
        PictureBoxQuestion13.Image = My.Resources.PT18Q13
        PictureBoxQuestion14.Image = My.Resources.PT18Q14
        PictureBoxQuestion15.Image = My.Resources.PT18Q15
        PictureBoxQuestion16.Image = My.Resources.PT18Q16
        PictureBoxQuestion17.Image = My.Resources.PT18Q17
        PictureBoxQuestion18.Image = My.Resources.PT18Q18
        PictureBoxQuestion19.Image = My.Resources.PT18Q19
        PictureBoxQuestion20.Image = My.Resources.PT18Q20
        PictureBoxQuestion21.Image = My.Resources.PT18Q21
        PictureBoxQuestion22.Image = My.Resources.PT18Q22
        PictureBoxQuestion23.Image = My.Resources.PT18Q23
        PictureBoxQuestion24.Image = My.Resources.PT18Q24
        PictureBoxQuestion25.Image = My.Resources.PT18Q25
        PictureBoxQuestion26.Image = My.Resources.PT18Q26
        PictureBoxQuestion27.Image = My.Resources.PT18Q27
        PictureBoxQuestion28.Image = My.Resources.PT18Q28
        PictureBoxQuestion29.Image = My.Resources.PT18Q29
        PictureBoxQuestion30.Image = My.Resources.PT18Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest19_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest19_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = True
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT19Q01
        PictureBoxQuestion02.Image = My.Resources.PT19Q02
        PictureBoxQuestion03.Image = My.Resources.PT19Q03
        PictureBoxQuestion04.Image = My.Resources.PT19Q04
        PictureBoxQuestion05.Image = My.Resources.PT19Q05
        PictureBoxQuestion06.Image = My.Resources.PT19Q06
        PictureBoxQuestion07.Image = My.Resources.PT19Q07
        PictureBoxQuestion08.Image = My.Resources.PT19Q08
        PictureBoxQuestion09.Image = My.Resources.PT19Q09
        PictureBoxQuestion10.Image = My.Resources.PT19Q10
        PictureBoxQuestion11.Image = My.Resources.PT19Q11
        PictureBoxQuestion12.Image = My.Resources.PT19Q12
        PictureBoxQuestion13.Image = My.Resources.PT19Q13
        PictureBoxQuestion14.Image = My.Resources.PT19Q14
        PictureBoxQuestion15.Image = My.Resources.PT19Q15
        PictureBoxQuestion16.Image = My.Resources.PT19Q16
        PictureBoxQuestion17.Image = My.Resources.PT19Q17
        PictureBoxQuestion18.Image = My.Resources.PT19Q18
        PictureBoxQuestion19.Image = My.Resources.PT19Q19
        PictureBoxQuestion20.Image = My.Resources.PT19Q20
        PictureBoxQuestion21.Image = My.Resources.PT19Q21
        PictureBoxQuestion22.Image = My.Resources.PT19Q22
        PictureBoxQuestion23.Image = My.Resources.PT19Q23
        PictureBoxQuestion24.Image = My.Resources.PT19Q24
        PictureBoxQuestion25.Image = My.Resources.PT19Q25
        PictureBoxQuestion26.Image = My.Resources.PT19Q26
        PictureBoxQuestion27.Image = My.Resources.PT19Q27
        PictureBoxQuestion28.Image = My.Resources.PT19Q28
        PictureBoxQuestion29.Image = My.Resources.PT19Q29
        PictureBoxQuestion30.Image = My.Resources.PT19Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest20_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest20_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = True
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT20Q01
        PictureBoxQuestion02.Image = My.Resources.PT20Q02
        PictureBoxQuestion03.Image = My.Resources.PT20Q03
        PictureBoxQuestion04.Image = My.Resources.PT20Q04
        PictureBoxQuestion05.Image = My.Resources.PT20Q05
        PictureBoxQuestion06.Image = My.Resources.PT20Q06
        PictureBoxQuestion07.Image = My.Resources.PT20Q07
        PictureBoxQuestion08.Image = My.Resources.PT20Q08
        PictureBoxQuestion09.Image = My.Resources.PT20Q09
        PictureBoxQuestion10.Image = My.Resources.PT20Q10
        PictureBoxQuestion11.Image = My.Resources.PT20Q11
        PictureBoxQuestion12.Image = My.Resources.PT20Q12
        PictureBoxQuestion13.Image = My.Resources.PT20Q13
        PictureBoxQuestion14.Image = My.Resources.PT20Q14
        PictureBoxQuestion15.Image = My.Resources.PT20Q15
        PictureBoxQuestion16.Image = My.Resources.PT20Q16
        PictureBoxQuestion17.Image = My.Resources.PT20Q17
        PictureBoxQuestion18.Image = My.Resources.PT20Q18
        PictureBoxQuestion19.Image = My.Resources.PT20Q19
        PictureBoxQuestion20.Image = My.Resources.PT20Q20
        PictureBoxQuestion21.Image = My.Resources.PT20Q21
        PictureBoxQuestion22.Image = My.Resources.PT20Q22
        PictureBoxQuestion23.Image = My.Resources.PT20Q23
        PictureBoxQuestion24.Image = My.Resources.PT20Q24
        PictureBoxQuestion25.Image = My.Resources.PT20Q25
        PictureBoxQuestion26.Image = My.Resources.PT20Q26
        PictureBoxQuestion27.Image = My.Resources.PT20Q27
        PictureBoxQuestion28.Image = My.Resources.PT20Q28
        PictureBoxQuestion29.Image = My.Resources.PT20Q29
        PictureBoxQuestion30.Image = My.Resources.PT20Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest21_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest21_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = True
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT21Q01
        PictureBoxQuestion02.Image = My.Resources.PT21Q02
        PictureBoxQuestion03.Image = My.Resources.PT21Q03
        PictureBoxQuestion04.Image = My.Resources.PT21Q04
        PictureBoxQuestion05.Image = My.Resources.PT21Q05
        PictureBoxQuestion06.Image = My.Resources.PT21Q06
        PictureBoxQuestion07.Image = My.Resources.PT21Q07
        PictureBoxQuestion08.Image = My.Resources.PT21Q08
        PictureBoxQuestion09.Image = My.Resources.PT21Q09
        PictureBoxQuestion10.Image = My.Resources.PT21Q10
        PictureBoxQuestion11.Image = My.Resources.PT21Q11
        PictureBoxQuestion12.Image = My.Resources.PT21Q12
        PictureBoxQuestion13.Image = My.Resources.PT21Q13
        PictureBoxQuestion14.Image = My.Resources.PT21Q14
        PictureBoxQuestion15.Image = My.Resources.PT21Q15
        PictureBoxQuestion16.Image = My.Resources.PT21Q16
        PictureBoxQuestion17.Image = My.Resources.PT21Q17
        PictureBoxQuestion18.Image = My.Resources.PT21Q18
        PictureBoxQuestion19.Image = My.Resources.PT21Q19
        PictureBoxQuestion20.Image = My.Resources.PT21Q20
        PictureBoxQuestion21.Image = My.Resources.PT21Q21
        PictureBoxQuestion22.Image = My.Resources.PT21Q22
        PictureBoxQuestion23.Image = My.Resources.PT21Q23
        PictureBoxQuestion24.Image = My.Resources.PT21Q24
        PictureBoxQuestion25.Image = My.Resources.PT21Q25
        PictureBoxQuestion26.Image = My.Resources.PT21Q26
        PictureBoxQuestion27.Image = My.Resources.PT21Q27
        PictureBoxQuestion28.Image = My.Resources.PT21Q28
        PictureBoxQuestion29.Image = My.Resources.PT21Q29
        PictureBoxQuestion30.Image = My.Resources.PT21Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest22_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest22_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = True
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT22Q01
        PictureBoxQuestion02.Image = My.Resources.PT22Q02
        PictureBoxQuestion03.Image = My.Resources.PT22Q03
        PictureBoxQuestion04.Image = My.Resources.PT22Q04
        PictureBoxQuestion05.Image = My.Resources.PT22Q05
        PictureBoxQuestion06.Image = My.Resources.PT22Q06
        PictureBoxQuestion07.Image = My.Resources.PT22Q07
        PictureBoxQuestion08.Image = My.Resources.PT22Q08
        PictureBoxQuestion09.Image = My.Resources.PT22Q09
        PictureBoxQuestion10.Image = My.Resources.PT22Q10
        PictureBoxQuestion11.Image = My.Resources.PT22Q11
        PictureBoxQuestion12.Image = My.Resources.PT22Q12
        PictureBoxQuestion13.Image = My.Resources.PT22Q13
        PictureBoxQuestion14.Image = My.Resources.PT22Q14
        PictureBoxQuestion15.Image = My.Resources.PT22Q15
        PictureBoxQuestion16.Image = My.Resources.PT22Q16
        PictureBoxQuestion17.Image = My.Resources.PT22Q17
        PictureBoxQuestion18.Image = My.Resources.PT22Q18
        PictureBoxQuestion19.Image = My.Resources.PT22Q19
        PictureBoxQuestion20.Image = My.Resources.PT22Q20
        PictureBoxQuestion21.Image = My.Resources.PT22Q21
        PictureBoxQuestion22.Image = My.Resources.PT22Q22
        PictureBoxQuestion23.Image = My.Resources.PT22Q23
        PictureBoxQuestion24.Image = My.Resources.PT22Q24
        PictureBoxQuestion25.Image = My.Resources.PT22Q25
        PictureBoxQuestion26.Image = My.Resources.PT22Q26
        PictureBoxQuestion27.Image = My.Resources.PT22Q27
        PictureBoxQuestion28.Image = My.Resources.PT22Q28
        PictureBoxQuestion29.Image = My.Resources.PT22Q29
        PictureBoxQuestion30.Image = My.Resources.PT22Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest23_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest23_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = True
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT23Q01
        PictureBoxQuestion02.Image = My.Resources.PT23Q02
        PictureBoxQuestion03.Image = My.Resources.PT23Q03
        PictureBoxQuestion04.Image = My.Resources.PT23Q04
        PictureBoxQuestion05.Image = My.Resources.PT23Q05
        PictureBoxQuestion06.Image = My.Resources.PT23Q06
        PictureBoxQuestion07.Image = My.Resources.PT23Q07
        PictureBoxQuestion08.Image = My.Resources.PT23Q08
        PictureBoxQuestion09.Image = My.Resources.PT23Q09
        PictureBoxQuestion10.Image = My.Resources.PT23Q10
        PictureBoxQuestion11.Image = My.Resources.PT23Q11
        PictureBoxQuestion12.Image = My.Resources.PT23Q12
        PictureBoxQuestion13.Image = My.Resources.PT23Q13
        PictureBoxQuestion14.Image = My.Resources.PT23Q14
        PictureBoxQuestion15.Image = My.Resources.PT23Q15
        PictureBoxQuestion16.Image = My.Resources.PT23Q16
        PictureBoxQuestion17.Image = My.Resources.PT23Q17
        PictureBoxQuestion18.Image = My.Resources.PT23Q18
        PictureBoxQuestion19.Image = My.Resources.PT23Q19
        PictureBoxQuestion20.Image = My.Resources.PT23Q20
        PictureBoxQuestion21.Image = My.Resources.PT23Q21
        PictureBoxQuestion22.Image = My.Resources.PT23Q22
        PictureBoxQuestion23.Image = My.Resources.PT23Q23
        PictureBoxQuestion24.Image = My.Resources.PT23Q24
        PictureBoxQuestion25.Image = My.Resources.PT23Q25
        PictureBoxQuestion26.Image = My.Resources.PT23Q26
        PictureBoxQuestion27.Image = My.Resources.PT23Q27
        PictureBoxQuestion28.Image = My.Resources.PT23Q28
        PictureBoxQuestion29.Image = My.Resources.PT23Q29
        PictureBoxQuestion30.Image = My.Resources.PT23Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest24_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest24_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = True
        PictureBoxTest25_3.Visible = False

        PictureBoxQuestion01.Image = My.Resources.PT24Q01
        PictureBoxQuestion02.Image = My.Resources.PT24Q02
        PictureBoxQuestion03.Image = My.Resources.PT24Q03
        PictureBoxQuestion04.Image = My.Resources.PT24Q04
        PictureBoxQuestion05.Image = My.Resources.PT24Q05
        PictureBoxQuestion06.Image = My.Resources.PT24Q06
        PictureBoxQuestion07.Image = My.Resources.PT24Q07
        PictureBoxQuestion08.Image = My.Resources.PT24Q08
        PictureBoxQuestion09.Image = My.Resources.PT24Q09
        PictureBoxQuestion10.Image = My.Resources.PT24Q10
        PictureBoxQuestion11.Image = My.Resources.PT24Q11
        PictureBoxQuestion12.Image = My.Resources.PT24Q12
        PictureBoxQuestion13.Image = My.Resources.PT24Q13
        PictureBoxQuestion14.Image = My.Resources.PT24Q14
        PictureBoxQuestion15.Image = My.Resources.PT24Q15
        PictureBoxQuestion16.Image = My.Resources.PT24Q16
        PictureBoxQuestion17.Image = My.Resources.PT24Q17
        PictureBoxQuestion18.Image = My.Resources.PT24Q18
        PictureBoxQuestion19.Image = My.Resources.PT24Q19
        PictureBoxQuestion20.Image = My.Resources.PT24Q20
        PictureBoxQuestion21.Image = My.Resources.PT24Q21
        PictureBoxQuestion22.Image = My.Resources.PT24Q22
        PictureBoxQuestion23.Image = My.Resources.PT24Q23
        PictureBoxQuestion24.Image = My.Resources.PT24Q24
        PictureBoxQuestion25.Image = My.Resources.PT24Q25
        PictureBoxQuestion26.Image = My.Resources.PT24Q26
        PictureBoxQuestion27.Image = My.Resources.PT24Q27
        PictureBoxQuestion28.Image = My.Resources.PT24Q28
        PictureBoxQuestion29.Image = My.Resources.PT24Q29
        PictureBoxQuestion30.Image = My.Resources.PT24Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxTest25_1_Click(sender As Object, e As EventArgs) Handles PictureBoxTest25_1.Click
        PictureBoxTest01_3.Visible = False
        PictureBoxTest02_3.Visible = False
        PictureBoxTest03_3.Visible = False
        PictureBoxTest04_3.Visible = False
        PictureBoxTest05_3.Visible = False
        PictureBoxTest06_3.Visible = False
        PictureBoxTest07_3.Visible = False
        PictureBoxTest08_3.Visible = False
        PictureBoxTest09_3.Visible = False
        PictureBoxTest10_3.Visible = False
        PictureBoxTest11_3.Visible = False
        PictureBoxTest12_3.Visible = False
        PictureBoxTest13_3.Visible = False
        PictureBoxTest14_3.Visible = False
        PictureBoxTest15_3.Visible = False
        PictureBoxTest16_3.Visible = False
        PictureBoxTest17_3.Visible = False
        PictureBoxTest18_3.Visible = False
        PictureBoxTest19_3.Visible = False
        PictureBoxTest20_3.Visible = False
        PictureBoxTest21_3.Visible = False
        PictureBoxTest22_3.Visible = False
        PictureBoxTest23_3.Visible = False
        PictureBoxTest24_3.Visible = False
        PictureBoxTest25_3.Visible = True

        PictureBoxQuestion01.Image = My.Resources.PT25Q01
        PictureBoxQuestion02.Image = My.Resources.PT25Q02
        PictureBoxQuestion03.Image = My.Resources.PT25Q03
        PictureBoxQuestion04.Image = My.Resources.PT25Q04
        PictureBoxQuestion05.Image = My.Resources.PT25Q05
        PictureBoxQuestion06.Image = My.Resources.PT25Q06
        PictureBoxQuestion07.Image = My.Resources.PT25Q07
        PictureBoxQuestion08.Image = My.Resources.PT25Q08
        PictureBoxQuestion09.Image = My.Resources.PT25Q09
        PictureBoxQuestion10.Image = My.Resources.PT25Q10
        PictureBoxQuestion11.Image = My.Resources.PT25Q11
        PictureBoxQuestion12.Image = My.Resources.PT25Q12
        PictureBoxQuestion13.Image = My.Resources.PT25Q13
        PictureBoxQuestion14.Image = My.Resources.PT25Q14
        PictureBoxQuestion15.Image = My.Resources.PT25Q15
        PictureBoxQuestion16.Image = My.Resources.PT25Q16
        PictureBoxQuestion17.Image = My.Resources.PT25Q17
        PictureBoxQuestion18.Image = My.Resources.PT25Q18
        PictureBoxQuestion19.Image = My.Resources.PT25Q19
        PictureBoxQuestion20.Image = My.Resources.PT25Q20
        PictureBoxQuestion21.Image = My.Resources.PT25Q21
        PictureBoxQuestion22.Image = My.Resources.PT25Q22
        PictureBoxQuestion23.Image = My.Resources.PT25Q23
        PictureBoxQuestion24.Image = My.Resources.PT25Q24
        PictureBoxQuestion25.Image = My.Resources.PT25Q25
        PictureBoxQuestion26.Image = My.Resources.PT25Q26
        PictureBoxQuestion27.Image = My.Resources.PT25Q27
        PictureBoxQuestion28.Image = My.Resources.PT25Q28
        PictureBoxQuestion29.Image = My.Resources.PT25Q29
        PictureBoxQuestion30.Image = My.Resources.PT25Q30

        FunctionTests()
        FunctionHider()
    End Sub

    Private Sub PictureBoxFinish_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxFinish.Click
        Answer = 0
        PanelMessage.Visible = True

        If PictureBoxTest01_3.Visible = True Then
            If A01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If A02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If A04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If B05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If B06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If C07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If A08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If A09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If C10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If C11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If B14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If A16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If C17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If D18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If B19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If A20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If C21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If C22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If A23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If D24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If C25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If B26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If A27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If B28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If B29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If C30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest02_3.Visible = True Then
            If B01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If B02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If C04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If C05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If C07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If B09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If C10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If B11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If C12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If A13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If B14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If D16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If C17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If D18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If E19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If D20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If A21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If B22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If B23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If D24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If B25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If C26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If B27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If B28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If B29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If D30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest03_3.Visible = True Then
            If B01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If C02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If B04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If B05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If C06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If A07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If D09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If C10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If C11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If A12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If D13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If D14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If C15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If D16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If C17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If C19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If A20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If C21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If D22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If A23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If C24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If B25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If B26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If C27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If A28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If C29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If C30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest04_3.Visible = True Then
            If C01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If A02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If B04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If B05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If A06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If B07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If D08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If D09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If B10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If C11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If B14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If E16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If B17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If A19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If C20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If B21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If C22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If C23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If C24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If B25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If A26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If C27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If A28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If B29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If C30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest05_3.Visible = True Then
            If B01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If B02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If C03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If C04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If D05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If D06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If A07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If C10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If B11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If D13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If C14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If E16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If C17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If B18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If E19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If C20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If B21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If A22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If D23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If A24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If B25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If B26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If B27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If D28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If B29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If D30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest06_3.Visible = True Then
            If B01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If A02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If B04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If D05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If B07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If D10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If B11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If D14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If E16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If B17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If D19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If B20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If A21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If B22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If B23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If C24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If A25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If C26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If B27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If C28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If B29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If C30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest07_3.Visible = True Then
            If D01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If B02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If B04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If D05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If C07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If A08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If B10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If A11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If B14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If D15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If A16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If A17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If E19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If B20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If B21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If B22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If B23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If E24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If B25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If C26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If A27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If D28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If C29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If E30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest08_3.Visible = True Then
            If C01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If C02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If C03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If B04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If D05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If A07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If D09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If A10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If B11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If B13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If C14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If A15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If B16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If C17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If C19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If B20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If C21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If A22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If D23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If D24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If C25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If E26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If C27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If D28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If C29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If D30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest09_3.Visible = True Then
            If B01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If B02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If A04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If B05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If A06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If C07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If A10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If A11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If B13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If C14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If C16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If B17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If D18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If E19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If C20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If E21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If C22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If A23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If C24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If C25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If A26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If A27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If C28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If D29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If C30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest10_3.Visible = True Then
            If C01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If B02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If B04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If B05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If B07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If A10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If D11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If D12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If C14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If A16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If D17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If E19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If D20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If E21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If A22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If B23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If C24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If B25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If C26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If A27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If A28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If D29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If C30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest11_3.Visible = True Then
            If C01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If B02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If C04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If E05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If C07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If C10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If D11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If B12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If B13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If C14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If D15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If C16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If C17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If E19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If A20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If E21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If B22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If B23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If D24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If B25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If D26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If B27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If D28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If B29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If A30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest12_3.Visible = True Then
            If B01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If D02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If C04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If E05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If B06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If B07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If D10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If A11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If C12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If A14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If A15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If A16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If D17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If A18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If E19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If B20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If C21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If C22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If C23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If C24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If C25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If B26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If B27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If D28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If C29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If C30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest13_3.Visible = True Then
            If D01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If D02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If C04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If E05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If D07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If C10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If E11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If D12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If A14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If A15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If B16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If C17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If E19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If D20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If E21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If C22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If B23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If B24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If D25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If B26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If A27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If B28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If D29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If D30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest14_3.Visible = True Then
            If B01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If A02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If C04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If E05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If A06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If B07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If B10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If E11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If C12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If A14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If A15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If B16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If B17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If A18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If D19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If B20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If B21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If A22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If C23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If B24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If B25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If D26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If C27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If E28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If B29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If D30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest15_3.Visible = True Then
            If D01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If B02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If B04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If E05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If D07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If C10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If E11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If D12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If B13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If C14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If A15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If A16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If D17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If E19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If C20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If E21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If D22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If E23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If D24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If A25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If C26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If C27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If D28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If A29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If E30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest16_3.Visible = True Then
            If B01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If C02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If A04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If E05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If E07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If D08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If C10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If E11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If D14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If D15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If A16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If A17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If D19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If B20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If E21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If D22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If C23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If C24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If A25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If D26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If D27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If B28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If E29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If B30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest17_3.Visible = True Then
            If D01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If A02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If B04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If E05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If C07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If B09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If D10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If E11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If C12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If A13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If D14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If A15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If A16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If C17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If A18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If E19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If A20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If E21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If A22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If A23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If A24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If E25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If A26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If E27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If D28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If A29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If C30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest18_3.Visible = True Then
            If C01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If C02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If A04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If E05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If E07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If B10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If E11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If D13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If C14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If B16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If C17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If B19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If D20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If A21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If D22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If A23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If C24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If D25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If C26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If C27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If B28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If B29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If A30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest19_3.Visible = True Then
            If C01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If D02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If A04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If E05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If A06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If E07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If E08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If D09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If B10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If E11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If C14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If C16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If B17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If C19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If D20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If C21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If B22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If B23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If D24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If B25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If C26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If A27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If A28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If C29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If D30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest20_3.Visible = True Then
            If B01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If B02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If A03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If C04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If C05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If E06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If E07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If D08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If D10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If E11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If D13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If D14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If B16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If B17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If E18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If B19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If D20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If B21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If C22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If E23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If E24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If C25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If C26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If D27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If D28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If D29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If C30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest21_3.Visible = True Then
            If E01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If D02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If C03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If E04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If B05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If D06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If B07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If A08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If B10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If A11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If C12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If B13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If A14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If D15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If D16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If B17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If B18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If D19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If D20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If B21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If D22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If A23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If A24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If A25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If B26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If C27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If C28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If D29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If A30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest22_3.Visible = True Then
            If E01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If E02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If C03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If E04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If B05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If D06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If E07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If A08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If B09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If B10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If D11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If A12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If B13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If E14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If C15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If A16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If B17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If B18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If A19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If C20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If E21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If D22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If B23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If A24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If D25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If C26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If B27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If C28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If E29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If C30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest23_3.Visible = True Then
            If A01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If B02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If B03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If E04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If C05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If D06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If B07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If B08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If A09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If B10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If B11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If E12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If D13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If D14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If E16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If E17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If C18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If C19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If E20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If D21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If B22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If B23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If B24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If D25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If B26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If C27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If A28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If C29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If A30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest24_3.Visible = True Then
            If A01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If B02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If B03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If E04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If B05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If C06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If D07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If B08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If C09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If A10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If C11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If A12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If A14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If E15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If E16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If E17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If C18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If B19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If E20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If A21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If C22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If A23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If A24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If A25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If C26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If D27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If D28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If D29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If B30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        If PictureBoxTest25_3.Visible = True Then
            If D01_3.Visible = True Then : R01.Visible = True : W01.Visible = False : Answer = Answer + 1 : Else : W01.Visible = True : R01.Visible = False : End If
            If E02_3.Visible = True Then : R02.Visible = True : W02.Visible = False : Answer = Answer + 1 : Else : W02.Visible = True : R02.Visible = False : End If
            If C03_3.Visible = True Then : R03.Visible = True : W03.Visible = False : Answer = Answer + 1 : Else : W03.Visible = True : R03.Visible = False : End If
            If D04_3.Visible = True Then : R04.Visible = True : W04.Visible = False : Answer = Answer + 1 : Else : W04.Visible = True : R04.Visible = False : End If
            If A05_3.Visible = True Then : R05.Visible = True : W05.Visible = False : Answer = Answer + 1 : Else : W05.Visible = True : R05.Visible = False : End If
            If D06_3.Visible = True Then : R06.Visible = True : W06.Visible = False : Answer = Answer + 1 : Else : W06.Visible = True : R06.Visible = False : End If
            If D07_3.Visible = True Then : R07.Visible = True : W07.Visible = False : Answer = Answer + 1 : Else : W07.Visible = True : R07.Visible = False : End If
            If B08_3.Visible = True Then : R08.Visible = True : W08.Visible = False : Answer = Answer + 1 : Else : W08.Visible = True : R08.Visible = False : End If
            If E09_3.Visible = True Then : R09.Visible = True : W09.Visible = False : Answer = Answer + 1 : Else : W09.Visible = True : R09.Visible = False : End If
            If E10_3.Visible = True Then : R10.Visible = True : W10.Visible = False : Answer = Answer + 1 : Else : W10.Visible = True : R10.Visible = False : End If
            If A11_3.Visible = True Then : R11.Visible = True : W11.Visible = False : Answer = Answer + 1 : Else : W11.Visible = True : R11.Visible = False : End If
            If D12_3.Visible = True Then : R12.Visible = True : W12.Visible = False : Answer = Answer + 1 : Else : W12.Visible = True : R12.Visible = False : End If
            If C13_3.Visible = True Then : R13.Visible = True : W13.Visible = False : Answer = Answer + 1 : Else : W13.Visible = True : R13.Visible = False : End If
            If B14_3.Visible = True Then : R14.Visible = True : W14.Visible = False : Answer = Answer + 1 : Else : W14.Visible = True : R14.Visible = False : End If
            If B15_3.Visible = True Then : R15.Visible = True : W15.Visible = False : Answer = Answer + 1 : Else : W15.Visible = True : R15.Visible = False : End If
            If A16_3.Visible = True Then : R16.Visible = True : W16.Visible = False : Answer = Answer + 1 : Else : W16.Visible = True : R16.Visible = False : End If
            If C17_3.Visible = True Then : R17.Visible = True : W17.Visible = False : Answer = Answer + 1 : Else : W17.Visible = True : R17.Visible = False : End If
            If A18_3.Visible = True Then : R18.Visible = True : W18.Visible = False : Answer = Answer + 1 : Else : W18.Visible = True : R18.Visible = False : End If
            If C19_3.Visible = True Then : R19.Visible = True : W19.Visible = False : Answer = Answer + 1 : Else : W19.Visible = True : R19.Visible = False : End If
            If A20_3.Visible = True Then : R20.Visible = True : W20.Visible = False : Answer = Answer + 1 : Else : W20.Visible = True : R20.Visible = False : End If
            If D21_3.Visible = True Then : R21.Visible = True : W21.Visible = False : Answer = Answer + 1 : Else : W21.Visible = True : R21.Visible = False : End If
            If E22_3.Visible = True Then : R22.Visible = True : W22.Visible = False : Answer = Answer + 1 : Else : W22.Visible = True : R22.Visible = False : End If
            If B23_3.Visible = True Then : R23.Visible = True : W23.Visible = False : Answer = Answer + 1 : Else : W23.Visible = True : R23.Visible = False : End If
            If E24_3.Visible = True Then : R24.Visible = True : W24.Visible = False : Answer = Answer + 1 : Else : W24.Visible = True : R24.Visible = False : End If
            If C25_3.Visible = True Then : R25.Visible = True : W25.Visible = False : Answer = Answer + 1 : Else : W25.Visible = True : R25.Visible = False : End If
            If E26_3.Visible = True Then : R26.Visible = True : W26.Visible = False : Answer = Answer + 1 : Else : W26.Visible = True : R26.Visible = False : End If
            If C27_3.Visible = True Then : R27.Visible = True : W27.Visible = False : Answer = Answer + 1 : Else : W27.Visible = True : R27.Visible = False : End If
            If E28_3.Visible = True Then : R28.Visible = True : W28.Visible = False : Answer = Answer + 1 : Else : W28.Visible = True : R28.Visible = False : End If
            If A29_3.Visible = True Then : R29.Visible = True : W29.Visible = False : Answer = Answer + 1 : Else : W29.Visible = True : R29.Visible = False : End If
            If A30_3.Visible = True Then : R30.Visible = True : W30.Visible = False : Answer = Answer + 1 : Else : W30.Visible = True : R30.Visible = False : End If
        End If
        LabelAnswer.Text = Answer
        If Answer >= 15 Then
            PanelMessage.BackColor = Color.Green
        Else
            If Answer <= 12 Then
                PanelMessage.BackColor = Color.Red
            Else
                PanelMessage.BackColor = Color.Yellow
            End If
        End If
        LabelAnswer.Location = New Drawing.Point((PanelMessage.Size.Width - LabelAnswer.Size.Width) / 2, 10)
    End Sub

    Private Sub ButtonLogin_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ButtonLogin.Click
        FunctionLogin()
    End Sub

    Private Sub WebBrowserLogin_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowserLogin.DocumentCompleted
        If WebBrowserLogin.Url.ToString.Contains("myportal") Then
            PanelLogin.Visible = False
            WebBrowserResults.Navigate("http://admissions.nu.edu.kz/wps/myportal/!ut/p/b1/04_Sj7SwNLQ0tTA1N9CP0I_KSyzLTE8syczPS8wB8aPM4p0sPX2dnAwdDSzMfA0NPJ0sjVwtA10MLSzMgQoigQoMcABHA0L6w_WjwEqc3R09TMx9DAwsfNxNDTwdPUKDLAONjQ0cjaEK8Fjh55Gfm6qfG5Vj6anrqAgAfco5bg!!/dl4/d5/L2dJQSEvUUt3QS80SmtFL1o2X0I5SU1CQjFBMDhONUEwSUI5Q0QwTEYzMEww/")
            WebBrowserProfile.Navigate("http://admissions.nu.edu.kz/wps/myportal/!ut/p/b1/04_SjzQxNzY0NDQzMNaP0I_KSyzLTE8syczPS8wB8aPM4p0sPX2dnAwdDSzMfA0NPJ0sjVwtA10MLSzMgQoigQoMcABHA0L6w_WjwEqc3R09TMx9DAwsfNxNDTwdPUKDLAONjQ0cjaEK8Fjh55Gfm6qfG5Vj6anrqAgAe0PFOg!!/dl4/d5/L2dJQSEvUUt3QS80SmtFL1o2X0I5SU1CQjFBME9NVEEwSVROVlBCVDYxMDQ1/")
        Else
            If WebBrowserLogin.Url.ToString.Contains("wps/portal") Then
                PanelLogin.Visible = True
                PanelLogin.BringToFront()
                PanelLogOut.Visible = False
            End If
        End If
    End Sub

    Private Sub PictureBoxPersonalCabinet_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxPersonalCabinet.Click
        If System.Net.NetworkInformation.NetworkInterface.GetIsNetworkAvailable = True Then
            LabelPage2PersonalCabinet.Visible = False
            PanelLogOut.Visible = False
            PanelLogin.Visible = False
            WebBrowserResults.Visible = False
            WebBrowserLogin.Navigate("http://admissions.nu.edu.kz")
            WebBrowserResults.Navigate("")
            WebBrowserProfile.Navigate("")
            TextBoxPassword.Text = ""
            PanelPage1.Visible = False
            PanelPage2.Visible = True
            PanelPage3.Visible = False
            PictureBoxTests.Image = My.Resources.Tests_1
            PictureBoxPersonalCabinet.Image = My.Resources.PersonalCabinet_2
            PictureBoxRegistration.Image = My.Resources.Registration_1

            PanelPage1Content.Location = New System.Drawing.Point((PanelPage1.Size.Width - PanelPage1Content.Width) / 2, (PanelPage1.Size.Height - PanelPage1Content.Height) / 2)
        Else
            MsgBox("Sorry, You have no internet connection!")
        End If
    End Sub

    Private Sub PictureBoxTests_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxTests.Click
        PanelPage1.Visible = True
        PanelPage2.Visible = False
        PanelPage3.Visible = False
        PictureBoxTests.Image = My.Resources.Tests_2
        PictureBoxPersonalCabinet.Image = My.Resources.PersonalCabinet_1
        PictureBoxRegistration.Image = My.Resources.Registration_1
    End Sub

    Private Sub PictureBoxRegistration_Click(ByVal sender As Object, ByVal e As EventArgs) Handles PictureBoxRegistration.Click
        PanelPage1.Visible = False
        PanelPage2.Visible = False
        PanelPage3.Visible = True
        PictureBoxTests.Image = My.Resources.Tests_1
        PictureBoxPersonalCabinet.Image = My.Resources.PersonalCabinet_1
        PictureBoxRegistration.Image = My.Resources.Registration_2
    End Sub

    Private Sub ButtonLogOut_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ButtonLogOut.Click
        PanelLogin.Visible = False
        LabelPage2PersonalCabinet.Visible = False
        PanelLogOut.Visible = False
        WebBrowserResults.Visible = False
        WebBrowserLogin.Navigate("http://admissions.nu.edu.kz")
        WebBrowserResults.Navigate("")
        WebBrowserProfile.Navigate("")
        TextBoxPassword.Text = ""
    End Sub

    Private Sub WebBrowserResults_DocumentCompleted(ByVal sender As Object, ByVal e As WebBrowserDocumentCompletedEventArgs) Handles WebBrowserResults.DocumentCompleted
        If WebBrowserResults.Url.ToString.Contains("myportal") Then
            WebBrowserResults.Visible = True
        End If
    End Sub

    Private Sub WebBrowserProfile_DocumentCompleted(ByVal sender As System.Object, ByVal e As System.Windows.Forms.WebBrowserDocumentCompletedEventArgs) Handles WebBrowserProfile.DocumentCompleted
        If WebBrowserProfile.Url.ToString.Contains("myportal") Then
            PanelLogOut.Visible = True
            PanelLogOut.BringToFront()
            LabelPage2PersonalCabinet.Visible = True
            LabelPage2PersonalCabinet.BringToFront()
            LabelFirstName.Text = WebBrowserProfile.Document.GetElementById("firstName").GetAttribute("value").ToString
            LabelLastName.Text = WebBrowserProfile.Document.GetElementById("secondName").GetAttribute("value").ToString
            LabelEducation.Text = WebBrowserProfile.Document.GetElementById("typeOfEducation").GetAttribute("value").ToString
            LabelTelephone.Text = WebBrowserProfile.Document.GetElementById("telephone").GetAttribute("value").ToString
        End If
    End Sub

    Private Sub TextBoxPassword_KeyDown(ByVal sender As Object, ByVal e As System.Windows.Forms.KeyEventArgs) Handles TextBoxPassword.KeyDown
        If e.KeyCode = Keys.Enter Then
            FunctionLogin()
        End If
    End Sub

    Function FunctionLogin()
        WebBrowserLogin.Document.GetElementById("login").SetAttribute("value", TextBoxUsername.Text)
        WebBrowserLogin.Document.GetElementById("pass").SetAttribute("value", TextBoxPassword.Text)
        WebBrowserLogin.Document.GetElementById("logsubmit").InvokeMember("click")
    End Function

    Private Sub ButtonRegister_Click(sender As Object, e As EventArgs) Handles ButtonRegister.Click
        If TextBoxKeyNumber.Text = key = False Then
            MsgBox("Wrong key!", MsgBoxStyle.OkOnly, "Message")
        Else
            If TextBoxKeyNumber.Text = key Then
                My.Settings.php_design = "php_design"
                Me.Close()
            Else
                MsgBox("Wrong key!", MsgBoxStyle.OkOnly, "Message")
            End If
        End If
    End Sub

    Private Sub ButtonBuyAKeyNumber_Click(ByVal sender As Object, ByVal e As EventArgs) Handles ButtonBuyAKeyNumber.Click
        System.Diagnostics.Process.Start("http://kazplus.kz/activation.php")
    End Sub

    Private Sub LabelKazPlusURL_Click(ByVal sender As Object, ByVal e As EventArgs) Handles LabelKazPlusURL.Click
        System.Diagnostics.Process.Start("http://kazplus.kz")
    End Sub
End Class