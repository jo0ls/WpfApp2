Public Class Menu

  Sub New()

    ' This call is required by the designer.
    InitializeComponent()

    ' We get a datacontext set from the parent window, a bit after the constructor runs.
    ' So the tree binds to the RootNodes property, with each tree node bound to a MenuItem
    ' Each node has children defined by the Items property on the MenuItem
  End Sub

  Private Sub TreeView_SelectedItemChanged(sender As Object, e As RoutedPropertyChangedEventArgs(Of Object))
    Dim vm = TryCast(DataContext, MenuViewModel)
    If vm IsNot Nothing Then
      vm.SelectedMenuItem = TryCast(tree1.SelectedItem, MenuItem) ' read onl!
    End If
  End Sub

End Class
