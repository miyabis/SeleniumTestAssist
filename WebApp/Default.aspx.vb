Public Class _Default
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        Debug.Print("hoge")
    End Sub

    Private Sub btnTest_Click(sender As Object, e As EventArgs) Handles btnTest.Click
        Me.txtTest.Text = "btnTest Click!"
        Threading.Thread.Sleep(300)
    End Sub

End Class