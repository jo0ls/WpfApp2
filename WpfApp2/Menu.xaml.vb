Public Class Menu

  Sub New()

    ' This call is required by the designer.
    InitializeComponent()

    ' Add any initialization after the InitializeComponent() call.

    ' Add any initialization after the InitializeComponent() call.
    Dim o = Me.DataContext
    Dim vm = CType(o, MenuViewModel)
    For Each n In vm.RootNodes
      tree1.Items.Add(n)
    Next
  End Sub
End Class
