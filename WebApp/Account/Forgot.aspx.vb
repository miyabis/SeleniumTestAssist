Imports System
Imports System.Web
Imports Microsoft.AspNet.Identity
Imports Microsoft.AspNet.Identity.EntityFramework
Imports Microsoft.AspNet.Identity.Owin
Imports Owin

Partial Public Class ForgotPassword
    Inherits System.Web.UI.Page

    Protected Property StatusMessage() As String

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
    End Sub

    Protected Sub Forgot(sender As Object, e As EventArgs)
        If IsValid Then
            ' ユーザーのメール アドレスを検証
            Dim manager = Context.GetOwinContext().GetUserManager(Of ApplicationUserManager)()
            Dim user As ApplicationUser = manager.FindByName(Email.Text)
            If user Is Nothing OrElse Not manager.IsEmailConfirmed(user.Id) Then
                FailureText.Text = "ユーザーが存在しないか、未確認です。"
                ErrorMessage.Visible = True
                Return
            End If
            ' アカウントの確認とパスワードの再設定を有効にする方法については、http://go.microsoft.com/fwlink/?LinkID=320771 を参照してください
            ' コードとパスワードの再設定ページへのリダイレクトを記載した電子メールを送信します
            ' Dim code = manager.GeneratePasswordResetToken(user.Id)
            ' Dim callbackUrl = IdentityHelper.GetResetPasswordRedirectUrl(code, Request)
            ' manager.SendEmail(user.Id, "パスワード", "のリセット <a href=""" & callbackUrl & """>こちら</a> をクリックして、パスワードをリセットしてください.")
            loginForm.Visible = False
            DisplayEmail.Visible = True
        End If
    End Sub
End Class