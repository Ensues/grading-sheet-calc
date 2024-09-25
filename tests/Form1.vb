Public Class Form1
    Private ave As Double = 0
    Private sum As Double = 0
    Private low As Double = 0
    Private x As Integer = 0
    Private y As Integer = 0
    Private Sub Form1_Load(sender As Object, e As EventArgs) Handles MyBase.Load

    End Sub

    Private Sub Label2_Click(sender As Object, e As EventArgs) Handles Label2.Click

    End Sub

    Private Sub Button2_Click(sender As Object, e As EventArgs) Handles bttnClear.Click
        txtbxSub1.Clear()
        txtbxSub2.Clear()
        txtbxSub3.Clear()
        txtbxSub4.Clear()
        txtbxSub5.Clear()
        txtbxRem.Clear()
        txtbxLowest.Clear()
        txtbxAve.Clear()
    End Sub

    Private Sub bttnProcess_Click(sender As Object, e As EventArgs) Handles bttnProcess.Click
        sum = Val(txtbxSub1.Text) + Val(txtbxSub2.Text) + Val(txtbxSub3.Text) + Val(txtbxSub4.Text) + Val(txtbxSub5.Text)
        ave = Val(sum / 5)

        If txtbxSub1.Text < txtbxSub2.Text And txtbxSub1.Text < txtbxSub3.Text And txtbxSub1.Text < txtbxSub4.Text And txtbxSub1.Text < txtbxSub5.Text Then
            txtbxLowest.Text = txtbxSub1.Text
        ElseIf txtbxSub2.Text < txtbxSub1.Text And txtbxSub2.Text < txtbxSub3.Text And txtbxSub2.Text < txtbxSub4.Text And txtbxSub2.Text < txtbxSub5.Text Then
            txtbxLowest.Text = txtbxSub2.Text
        ElseIf txtbxSub3.Text < txtbxSub1.Text And txtbxSub3.Text < txtbxSub2.Text And txtbxSub3.Text < txtbxSub4.Text And txtbxSub3.Text < txtbxSub5.Text Then
            txtbxLowest.Text = txtbxSub3.Text
        ElseIf txtbxSub4.Text < txtbxSub1.Text And txtbxSub4.Text < txtbxSub2.Text And txtbxSub4.Text < txtbxSub3.Text And txtbxSub4.Text < txtbxSub5.Text Then
            txtbxLowest.Text = txtbxSub4.Text
        ElseIf txtbxSub5.Text < txtbxSub1.Text And txtbxSub5.Text < txtbxSub2.Text And txtbxSub5.Text < txtbxSub3.Text And txtbxSub5.Text < txtbxSub4.Text Then
            txtbxLowest.Text = txtbxSub5.Text
        End If

        If 96 < ave And ave < 100 Then
            txtbxAve.Text = ave
            txtbxRem.Text = "Excellent"
        ElseIf 86 < ave And ave < 95 Then
            txtbxAve.Text = ave
            txtbxRem.Text = "Very Good"
        ElseIf 76 < ave And ave < 85 Then
            txtbxAve.Text = ave
            txtbxRem.Text = "good"
        ElseIf 66 < ave And ave < 75 Then
            txtbxAve.Text = ave
            txtbxRem.Text = "Satisfactory"
        ElseIf 60 < ave And ave < 65 Then
            txtbxAve.Text = ave
            txtbxRem.Text = "Pass"
        ElseIf 51 < ave And ave < 59 Then
            txtbxAve.Text = ave
            txtbxRem.Text = "Conditional Pass"
        ElseIf 41 < ave And ave < 50 Then
            txtbxAve.Text = ave
            txtbxRem.Text = "Conditional Fail"
        ElseIf 0 < ave And ave < 40 Then
            txtbxAve.Text = ave
            txtbxRem.Text = "Fail"
        ElseIf ave > -0 And 100 < ave Then
            MsgBox("Pameters not fitted", vbOKOnly + vbObjectError, "Error")
            Stop
        End If


    End Sub
End Class
