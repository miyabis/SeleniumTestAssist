
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Support.PageObjects
Imports MiYABiS.SeleniumTestAssist

Public Class LoginPage
    Inherits SeleniumAction

    <FindsBy([Using]:="MainContent_Email")>
    Private _email As IWebElement

    <FindsBy([Using]:="MainContent_Password")>
    Private _password As IWebElement

    <FindsBy([Using]:="MainContent_RememberMe")>
    Private _rememberMe As IWebElement

    <FindsBy([Using]:="MainContent_btnLogIn")>
    Private _LogIn As IWebElement

    Public Sub New(ByVal driver As IWebDriver)
        MyBase.New(driver)
    End Sub

    Public Sub Email(ByVal value As String)
        Typing(_email, value)
    End Sub

    Public Sub Password(ByVal value As String)
        Typing(_password, value)
    End Sub

    Public Sub RememberMe(ByVal value As Boolean)
        Check(_rememberMe, value)
    End Sub

    Public Sub LogIn()
        Click(_LogIn)
    End Sub

End Class
