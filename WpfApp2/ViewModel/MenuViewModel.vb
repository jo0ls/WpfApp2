Public Class MenuViewModel

  Private cono As String
  Private connectionString As String
  Private dtSSMENUS As New DataTable

  Property RootNodes As List(Of MenuITem)

  Public Sub New()

    If DesignerProperties.GetIsInDesignMode(New DependencyObject()) Then Return

    cono = System.Windows.Application.Current.Resources.GetString("CONO")
    connectionString = Windows.Application.Current.Resources.GetString("ConnectionString")
    GetSSMENUS()

    RootNodes = New List(Of MenuITem)
    Dim lookup As New Dictionary(Of String, List(Of MenuITem)) ' parent -> nodes
    For Each dr As DataRow In dtSSMENUS.Rows
      Dim MNU_PARENT As String = dr.GetTrimStringOrEmpty("MNU_PARENT")
      Dim MNU_PARAM As String = dr.GetTrimStringOrEmpty("MNU_PARAM")
      Dim MNU_DESCRIPTION As String = dr.GetTrimStringOrEmpty("MNU_DESCRIPTION")
      Dim item As New MenuITem With {.Caption = MNU_DESCRIPTION, .Parameter = MNU_PARAM}
      If MNU_PARENT = "" Then
        RootNodes.Add(item)
      Else
        Dim items As List(Of MenuITem) = Nothing
        If lookup.TryGetValue(MNU_PARENT, items) = False Then
          items = New List(Of MenuITem)
          lookup(MNU_PARENT) = items
        End If
        items.Add(item)
      End If
    Next

    Dim s As New Stack(Of MenuITem)(RootNodes.ToArray)
    While s.Count > 0
      Dim item = s.Pop
      Dim items As List(Of MenuITem) = Nothing
      If lookup.TryGetValue(item.Parameter, items) Then
        For Each mi In items
          item.Items.Add(mi)
          s.Push(mi)
        Next
      Else
        ' no children
      End If
    End While


  End Sub


  Private Function GetSSMENUS() As Boolean

    Using conn = GetConnection()
      Using cmd As New SqlCommand("SELECT * FROM dbo.[SSMENUS] WHERE MNU_NAME='SSMENUX'", conn)
        dtSSMENUS.Load(cmd.ExecuteReader(CommandBehavior.SingleResult))
      End Using
    End Using
    Return True
  End Function

  Private Function GetConnection() As SqlConnection
    Dim conn As New SqlConnection(connectionString)
    conn.Open()
    Return conn
  End Function

End Class
