

Public Class MenuItem
  Inherits DependencyObject

  Property Id As Integer ' int PK
  Property Items As ObservableCollection(Of MenuItem)
  Property Caption As String ' MNU_DESCRIPTION varchar(80)
  Property Parameter As String ' MNU_PARAM vachar(50)
  Property MenuType As String ' MNU_TYPE varchar(1)
  Property Args As String  ' MNU_ARGS varchar(250)
  Property Flag As String ' MNU_SPECIALFLAG varchar(10)
  Property Parent As String ' MNU_PARENT varchar(30)
  Property IsLeaf As Boolean
  Property CustomViewName As String

  Public Shared Function FromDataRow(dr As DataRow) As MenuItem
    Dim item As New MenuItem
    item.Id = dr.GetIntegerOrZero("id")
    item.Parent = dr.GetTrimStringOrEmpty("MNU_PARENT")
    item.Parameter = dr.GetTrimStringOrEmpty("MNU_PARAM")
    item.Caption = dr.GetTrimStringOrEmpty("MNU_DESCRIPTION")
    item.MenuType = dr.GetTrimStringOrEmpty("MNU_TYPE")
    item.Args = dr.GetTrimStringOrEmpty("MNU_ARGS")
    item.Flag = dr.GetTrimStringOrEmpty("MNU_SPECIALFLAG")
    Return item
  End Function


End Class
