Public Class _Default
    Inherits Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As EventArgs) Handles Me.Load
        If Not IsPostBack Then
            Dim s As New SOP
            s.loadItems("15")

            DDLItems.DataSource = s.itemList
            DDLItems.DataTextField = "Text"
            DDLItems.DataValueField = "ID"
            DDLItems.DataBind()
        End If
    End Sub

    Protected Sub Button1_Click(sender As Object, e As EventArgs) Handles Button1.Click
        Dim sel As String = DDLItems.SelectedItem.Value
        Dim selint As Integer = Convert.ToInt32(sel)
        Dim si As SOP.SOP_Items = New SOP.SOP_Items(selint)
        si.loadBranches(True)
        si.generateCompletionTracking()

    End Sub
End Class