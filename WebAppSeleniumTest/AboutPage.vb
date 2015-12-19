
Imports OpenQA.Selenium
Imports OpenQA.Selenium.Support.PageObjects
Imports MiYABiS.SeleniumTestAssist

Public Class AboutPage
    Inherits SeleniumAction

    <FindsBy(How:=How.CssSelector, [Using]:="div.container h2")>
    Private _hoge As IWebElement

    Public Sub New(ByVal driver As IWebDriver)
        MyBase.New(driver)
    End Sub

    Public Sub H2Assert()
        Assert.AreEqual("About.", _hoge.Text)
    End Sub

End Class
