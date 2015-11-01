Public Class Demo
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load
        Dim lst As New List(Of String)
        lst.Add("demo1")
        lst.Add("demo2")
        lst.Add("demo3")
        Me.DropDownList1.DataSource = lst
        Me.DropDownList1.DataBind()

        Me.CheckBoxList1.DataSource = lst
        Me.CheckBoxList1.DataBind()

        Me.RadioButtonList1.DataSource = lst
        Me.RadioButtonList1.DataBind()

        Me.ListBox1.DataSource = lst
        Me.ListBox1.DataBind()
    End Sub

End Class