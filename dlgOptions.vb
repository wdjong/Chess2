Imports System.Windows.Forms

Public Class DlgOptions
    Public Sub New()

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        cbShareThoughts.Checked = My.Settings.ShareThoughts

    End Sub

    Private Sub OK_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles OK_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.OK
        Me.Close()
    End Sub

    Private Sub Cancel_Button_Click(ByVal sender As System.Object, ByVal e As System.EventArgs) Handles Cancel_Button.Click
        Me.DialogResult = System.Windows.Forms.DialogResult.Cancel
        Me.Close()
    End Sub

    Private Sub CBShareThoughts_CheckedChanged(sender As Object, e As EventArgs) Handles cbShareThoughts.CheckedChanged
        My.Settings.ShareThoughts = cbShareThoughts.Checked
    End Sub
End Class
