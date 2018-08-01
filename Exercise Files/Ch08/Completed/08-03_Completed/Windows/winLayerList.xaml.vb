Public Class winLayerList

    Public Sub New(Layers As Objects.LayerObjectCollection)

        ' This call is required by the designer.
        InitializeComponent()

        ' Add any initialization after the InitializeComponent() call.
        Me.DataContext = Layers
    End Sub

    Private Sub btn_Ok_Click(sender As Object, e As Windows.RoutedEventArgs) Handles btn_Ok.Click
        Me.DialogResult = True
        Me.Close()
    End Sub

    Private Sub btn_Cancel_Click(sender As Object, e As Windows.RoutedEventArgs) Handles btn_Cancel.Click
        Me.DialogResult = False
        Me.Close()
    End Sub
End Class
