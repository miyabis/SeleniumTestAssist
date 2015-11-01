
Imports OpenQA.Selenium
Imports MiYABiS.SeleniumTestAssist

Public Class DemoPage
    Inherits SeleniumAction

    Public Overrides ReadOnly Property MyPageName As String
        Get
            Return "/Demo"
        End Get
    End Property

    Public Sub New(ByVal driver As IWebDriver, ByVal baseUrl As String)
        MyBase.New(driver, baseUrl)
    End Sub

    Public Sub TextBox1(ByVal value As String)
        Typing("TextBox1", value)
    End Sub

    Public Sub CheckBox1(ByVal value As Boolean)
        Check("CheckBox1", True)
    End Sub

    Public Sub DropDownList1(ByVal value As String)
        SelectByText("DropDownList1", value)
    End Sub

End Class
