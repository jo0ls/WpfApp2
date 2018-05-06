Imports System.IO
Imports System.Text

Public Class MenuViewModel
  Inherits DependencyObject
  Implements INotifyPropertyChanged

  Private cono As String
  Private connectionString As String
  Private OPR_CODE As String
  Private dtSSMENUS As New DataTable
  Private designMode As Boolean

  Property RootNodes As List(Of MenuItem)

  Shared ReadOnly SelectedMenuItemProperty As DependencyProperty =
    DependencyProperty.Register("SelectedMenuItem", GetType(MenuItem), GetType(Menu))
  Public Event PropertyChanged As PropertyChangedEventHandler Implements INotifyPropertyChanged.PropertyChanged

  Public Property SelectedMenuItem As MenuItem
    Get
      Return TryCast(GetValue(SelectedMenuItemProperty), MenuItem)
    End Get
    Set(value As MenuItem)
      SetValue(SelectedMenuItemProperty, value)
      RaiseEvent PropertyChanged(Me, New PropertyChangedEventArgs(NameOf(SelectedMenuItem)))
    End Set
  End Property

  Public Sub New()

    If DesignerProperties.GetIsInDesignMode(New DependencyObject()) Then
      designMode = True
      cono = "1"
      GetDummySSMENUS()
    Else
      cono = System.Windows.Application.Current.Resources.GetString("CONO")
      connectionString = Windows.Application.Current.Resources.GetString("ConnectionString")
      OPR_CODE = Windows.Application.Current.Resources.GetString("OPR_CODE")
      GetSSMENUS()
    End If

    RootNodes = New List(Of MenuItem)
    Dim lookup As New Dictionary(Of String, List(Of MenuItem)) ' parent -> nodes
    For Each dr As DataRow In dtSSMENUS.Rows
      Dim item = MenuItem.FromDataRow(dr)
      If item.Parent = "" Then
        RootNodes.Add(item)
      Else
        Dim items As List(Of MenuItem) = Nothing
        If lookup.TryGetValue(item.Parent, items) = False Then
          items = New List(Of MenuItem)
          lookup(item.Parent) = items
        End If
        items.Add(item)
      End If
    Next

    Dim s As New Stack(Of MenuItem)(RootNodes.ToArray)
    While s.Count > 0
      Dim item = s.Pop
      Dim items As List(Of MenuItem) = Nothing
      If lookup.TryGetValue(item.Parameter, items) Then
        item.Items = New ObservableCollection(Of MenuItem)
        For Each mi In items
          item.Items.Add(mi)
          s.Push(mi)
        Next
      Else
        ' no children
        If item.Parameter = "MXLVIEWS" Then
          If Me.designMode Then
            GetViewsDummy(item)
          Else
            GetViews(item)
            If item.Items IsNot Nothing Then GetCustomViews(item.Items)
          End If
        Else
          item.IsLeaf = True
        End If
      End If
    End While
  End Sub

  Private Function GetBaseViewName(sb As StringBuilder) As StringBuilder
    Dim dotIndex As Integer?
    Dim underscoreIndex As Integer?
    For i As Integer = 0 To sb.Length - 1
      Select Case sb(i)
        Case "."c
          dotIndex = i
        Case "_"c
          underscoreIndex = i
      End Select
    Next

    ' assuming dot before underscore...

    If underscoreIndex.HasValue Then
      ' remove after underscore
      sb.Remove(underscoreIndex.Value, sb.Length - underscoreIndex.Value)
    End If

    If dotIndex.HasValue Then
      '  [dbo].concept_something_blah_boo_co1
      ' strip up to dot
      sb.Remove(0, dotIndex.Value + 1)
    End If

    Return sb
  End Function

  Private Sub GetCustomViews(children As IEnumerable(Of MenuItem))
    Dim sql As String
    If String.IsNullOrWhiteSpace(OPR_CODE) Then
      sql = $"SELECT CustomName, OPR_CODE FROM SS_Custom_Views WHERE BaseViewName=@baseViewName And ISNULL(OPR_CODE, '') = ''"
    Else
      sql = $"SELECT CustomName, OPR_CODE FROM SS_Custom_Views WHERE BaseViewName=@baseViewName AND ISNULL(OPR_CODE, '') IN (@OPR_CODE, '')"
    End If
    Using connection = GetConnection()

      Using cmd As New SqlCommand(sql, connection)
        If Not String.IsNullOrEmpty(OPR_CODE) Then cmd.Parameters.Add("@OPR_CODE", SqlDbType.VarChar, 10).Value = OPR_CODE
        Dim pViewName = cmd.Parameters.Add("@baseViewName", SqlDbType.VarChar, 384)
        Dim sb As New StringBuilder
        For Each item In children
          sb.Clear()
          sb.Append(item.Args)
          Dim args = sb.ToString ' keep for the menu item
          GetBaseViewName(sb) ' trims it back without [dbo]. or _co1 assumes 1 dot really, and dot is before the last underscore.
          pViewName.Value = sb.ToString

          Using reader = cmd.ExecuteReader(CommandBehavior.SingleResult)
            ' if nothing is read, then there's no node beneath
            item.IsLeaf = True ' assume true, say false if anything below
            While reader.Read
              item.IsLeaf = False

              ' add 1 default node 1st
              If item.Items Is Nothing Then
                Dim defaultMi As New MenuItem With {
                .Parameter = "PXLCUSTOMGLOBAL",
                .Args = item.Args, 'args of higher node.
                .Caption = item.Caption & " (Default)",
                .CustomViewName = "",
                .Flag = "^",
                .MenuType = "ECV",
                .Parent = item.Parameter
              }
                item.Items = New ObservableCollection(Of MenuItem)
                item.Items.Add(defaultMi)
              End If

              ' and the nodes for saved views
              Dim mi As New MenuItem
              mi.Caption = reader.GetTrimStringOrEmpty(0)
              If String.IsNullOrWhiteSpace(reader.GetTrimStringOrEmpty(1)) Then
                ' OPR_CODE blank means a global view
                mi.Parameter = "PXLCUSTOMGLOBAL"
              Else
                mi.Parameter = "PXLCUSTOM"
              End If
              mi.Args = args
              mi.MenuType = "ECV"
              mi.Flag = "^"
              mi.CustomViewName = mi.Caption
              mi.Parent = item.Parameter
              item.Items.Add(mi)
            End While
          End Using
        Next
      End Using
    End Using
  End Sub

  Private Sub GetViewsDummy(item As MenuItem)
    Dim filter = If(item.Args = "", "concept", item.Args)
    Dim coPart = "_co" & CInt(cono)
    Try
      Dim descBuilder As New StringBuilder

      Dim reader As New StringReader(My.Resources.ResourceManager.GetString(filter))
      While reader.Peek > -1
        Dim TABLE_OWNER = "dbo"
        Dim TABLE_NAME = reader.ReadLine
        descBuilder.Clear()
        descBuilder.Append(TABLE_NAME.Substring(filter.Length).TrimStart({"_"c})) ' remove prefix, e.g. concept_mgmt
        ' remove _co1 from end
        If TABLE_NAME.EndsWith(coPart) Then descBuilder.Remove(descBuilder.Length - (coPart.Length), coPart.Length)
        descBuilder.Replace("_", " ")
        ' proper case
        ' Dim desc2 = ti.ToTitleCase(descBuilder.ToString)
        ProperCase(descBuilder)

        Dim child As New MenuItem
        child.Parameter = "PXLVIEWS"
        child.Args = $"[{TABLE_OWNER}].{TABLE_NAME}"
        child.Flag = "^" ' I think this means 'go and look for custom views'
        child.Caption = descBuilder.ToString
        child.MenuType = "P"
        child.Parent = item.Parameter ' the menu is linked on parent node's "param" = this child's "parent"
        child.CustomViewName = ""
        If item.Items Is Nothing Then item.Items = New ObservableCollection(Of MenuItem)
        item.Items.Add(child)
      End While
    Catch ex As Exception
    End Try
  End Sub

  Private Sub GetViews(item As MenuItem)
    Dim filter = If(item.Args = "", "concept", item.Args)
    Dim coPart = "_co" & CInt(cono) ' _co1 for string manipulation
    ' Dim ti As Globalization.TextInfo = Globalization.CultureInfo.InvariantCulture.TextInfo
    Try
      Dim descBuilder As New StringBuilder
      Using connection = GetConnection()
        Using cmd As New SqlCommand("dbo.[csp_views]", connection)
          cmd.CommandType = CommandType.StoredProcedure
          cmd.Parameters.Add("@table_name", SqlDbType.VarChar, 384).Value = $"{filter}%co{CInt(cono)}"
          Using reader As SqlDataReader = cmd.ExecuteReader()
            While reader.Read
              Dim TABLE_OWNER = reader.GetTrimStringOrEmpty(1)
              Dim TABLE_NAME = reader.GetTrimStringOrEmpty(2)

              descBuilder.Clear()
              descBuilder.Append(TABLE_NAME.Substring(filter.Length).TrimStart({"_"c})) ' remove prefix, e.g. concept_mgmt
              ' remove _co1 from end
              If TABLE_NAME.EndsWith(coPart) Then descBuilder.Remove(descBuilder.Length - (coPart.Length), coPart.Length)
              descBuilder.Replace("_", " ")
              ' proper case
              ' Dim desc2 = ti.ToTitleCase(descBuilder.ToString)
              ProperCase(descBuilder)

              Dim child As New MenuItem
              child.Parameter = "PXLVIEWS"
              child.Args = $"[{TABLE_OWNER}].{TABLE_NAME}"
              child.Flag = "^" ' I think this means 'go and look for custom views'
              child.Caption = descBuilder.ToString
              child.MenuType = "P"
              child.Parent = item.Parameter ' the menu is linked on parent node's "param" = this child's "parent"
              child.CustomViewName = ""
              If item.Items Is Nothing Then item.Items = New ObservableCollection(Of MenuItem)
              item.Items.Add(child)
            End While
          End Using
        End Using
      End Using

    Catch ex As Exception
    End Try
  End Sub

  Shared Function ProperCase(desc As StringBuilder) As StringBuilder
    If desc.Length > 1 Then
      For i As Integer = 0 To desc.Length - 1
        If i = 0 OrElse desc(i - 1) = " " Then
          desc(i) = Char.ToUpper(desc(i))
        End If
      Next
    End If
    Return desc
  End Function

  Private Function GetSSMENUS() As Boolean

    Using conn = GetConnection()
      Using cmd As New SqlCommand("SELECT * FROM dbo.[SSMENUS] WHERE MNU_NAME='SSMENUX'", conn)
        dtSSMENUS.Load(cmd.ExecuteReader(CommandBehavior.SingleResult))
      End Using
    End Using
    Return True
  End Function

  Private Function GetDummySSMENUS() As Boolean
    dtSSMENUS = New DataTable
    dtSSMENUS.Columns.Add("id", GetType(Integer))

    dtSSMENUS.Columns.Add("MNU_NAME", GetType(String))
    dtSSMENUS.Columns.Add("MNU_PARENT", GetType(String))
    dtSSMENUS.Columns.Add("MNU_DESCRIPTION", GetType(String))
    dtSSMENUS.Columns.Add("MNU_TYPE", GetType(String))
    dtSSMENUS.Columns.Add("MNU_PARAM", GetType(String))
    dtSSMENUS.Columns.Add("MNU_ARGS", GetType(String))
    dtSSMENUS.Columns.Add("MNU_SPECIALFLAG", GetType(String))
    Dim sr As New StringReader(My.Resources.DummySSMENUS)
    While sr.Peek > -1
      Dim line As String = sr.ReadLine
      Dim parts() As String = line.Split({","}, StringSplitOptions.None)
      Dim dr As DataRow = dtSSMENUS.NewRow
      dr(0) = CInt(parts(0))
      dr(1) = parts(1)
      dr(2) = parts(2)
      dr(3) = parts(3)
      dr(4) = parts(4)
      dr(5) = parts(5)
      dr(6) = parts(6)
      dtSSMENUS.Rows.Add(dr)
    End While
    Return True
  End Function

  Private Function GetConnection() As SqlConnection
    Dim conn As New SqlConnection(connectionString)
    conn.Open()
    Return conn
  End Function

End Class
