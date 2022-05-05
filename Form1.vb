Public Class Form1
    Dim conex As New OleDb.OleDbConnection("Provider=Microsoft.Jet.OLEDB.4.0;Data Source=..\..\bin\Debug\Base_Datos\Comunas.mdb")
    Dim ctrImg, ctrImg2, ctrLargo, numero, largo As Integer
    Dim Sql, comunas, vector(30), ruta, letra As String
    Dim cmd As New OleDb.OleDbCommand
    Dim dr As OleDb.OleDbDataReader
    Dim numAleatorio As New Random

    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load
        Try
            conex.Open()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try
        Sql = "Select * FROM Comunas"
        Dim adapta_1 As New OleDb.OleDbDataAdapter(Sql, conex)
        Dim datatbl_1 As New DataTable()
        adapta_1.Fill(datatbl_1)
        ctrImg = 1
        ctrImg2 = 0
        ctrLargo = 0
        ruta = "..\..\bin\Debug\Fotos_Colgado\"
    End Sub

    Private Sub FINToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles FINToolStripMenuItem.Click
        End
    End Sub

    Private Sub NuevoJuegoToolStripMenuItem_Click(sender As Object, e As EventArgs) Handles NuevoJuegoToolStripMenuItem.Click

        PictureBox1.Image = Image.FromFile(ruta & "FI_1.png")
        ctrImg = 1
        numero = numAleatorio.Next(1, 20)
        cmd.Connection = conex
        cmd.CommandType = CommandType.Text
        cmd.CommandText = "select * from comunas where numero = " & numero

        Try
            dr = cmd.ExecuteReader()
            If dr.HasRows Then
                While dr.Read()
                    Label28.Text = dr(1)
                    comunas = dr(1)
                End While
            End If
            dr.Close()
        Catch ex As Exception
            MsgBox(ex.ToString)
        End Try

        largo = Len(comunas)
        labelEnable(True)
        For Each labelV In Me.Controls
            If TypeOf labelV Is Label Then
                If labelV.tag >= 1 And labelV.tag <= 15 Then
                    labelV.Text = ""
                    labelV.Visible = False
                End If
            End If
        Next
        For Each labelV In Me.Controls
            If TypeOf labelV Is Label Then
                If labelV.tag >= 1 And labelV.tag <= largo Then
                    labelV.Visible = True
                End If
            End If
        Next

    End Sub

    Private Sub clickLabels(sender As Object, e As EventArgs) Handles Label9.Click, Label8.Click, Label7.Click, Label6.Click, Label5.Click, Label4.Click, Label3.Click, Label27.Click, Label26.Click, Label25.Click, Label24.Click, Label23.Click, Label22.Click, Label21.Click, Label20.Click, Label2.Click, Label19.Click, Label18.Click, Label17.Click, Label16.Click, Label15.Click, Label14.Click, Label13.Click, Label12.Click, Label11.Click, Label10.Click, Label1.Click

        For Each labelColor In Me.Controls
            If TypeOf labelColor Is Label Then
                If labelColor.tag >= 100 Then
                    labelColor.BackColor = Color.White
                End If
            End If
        Next
        sender.BackColor = Color.Red
        For i = 1 To largo
            vector(i) = Mid(comunas, i, 1)
        Next

        letra = sender.text
        ctrImg += 1
        For i = 1 To largo
            If vector(i) = letra Then
                ctrImg2 = 1
                ctrImg -= ctrImg2
                If ctrImg <= 0 Then
                    ctrImg = 1
                End If
                sender.tag = 100
                Select Case i
                    Case 1
                        forCase(Label30)
                        sender.Enabled = False
                    Case 2
                        forCase(Label31)
                        sender.Enabled = False
                    Case 3
                        forCase(Label32)
                        sender.Enabled = False
                    Case 4
                        forCase(Label33)
                        sender.Enabled = False
                    Case 5
                        forCase(Label34)
                        sender.Enabled = False
                    Case 6
                        forCase(Label35)
                        sender.Enabled = False
                    Case 7
                        forCase(Label36)
                        sender.Enabled = False
                    Case 8
                        forCase(Label37)
                        sender.Enabled = False
                    Case 9
                        forCase(Label38)
                        sender.Enabled = False
                    Case 10
                        forCase(Label39)
                        sender.Enabled = False
                    Case 11
                        forCase(Label40)
                        sender.Enabled = False
                    Case 12
                        forCase(Label41)
                        sender.Enabled = False
                    Case 13
                        forCase(Label42)
                        sender.Enabled = False
                    Case 14
                        forCase(Label43)
                        sender.Enabled = False
                    Case 15
                        forCase(Label44)
                        sender.Enabled = False
                End Select
            End If
            PictureBox1.Image = Image.FromFile(ruta & "FI_" & ctrImg & ".png")
            If ctrImg = 6 Then
                labelEnable(False)
                PictureBox1.Image = Image.FromFile(ruta & "FI_6.png")
                sender.tag = 100
                sender.BackColor = Color.White
                Label45.Visible = True
                Label45.Text = "PERDIO!"
                ctrImg = 1
                Exit For
            End If

            If ctrLargo = largo Then
                labelEnable(False)
                sender.tag = 100
                sender.BackColor = Color.White
                Label45.Visible = True
                Label45.Text = "GANO!"
                ctrLargo = 0
                Exit For
            End If
        Next

    End Sub
    Private Sub labelEnable(LaEbl)
        For Each labelEnabled In Me.Controls
            If TypeOf labelEnabled Is Label Then
                If labelEnabled.tag = 100 Then
                    labelEnabled.Enabled = LaEbl
                End If
            End If
        Next
    End Sub
    Private Sub forCase(labelText)
        labelText.Text = letra
        If letra = labelText.Text Then ctrLargo += 1
    End Sub
End Class