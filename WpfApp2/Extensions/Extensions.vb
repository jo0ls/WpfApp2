Imports System.Data
Imports System.Data.SqlClient
Imports System.Runtime.CompilerServices

Module Extensions

  <Extension>
  Function GetString(dic As ResourceDictionary, key As String) As String
    ' resourcedictionary type is similar to a dic, but slightly unhelpfully names things differently and
    ' doesn't have the same methods.
    If dic.Contains(key) Then
      Return TryCast(dic(key), String)
    End If
    Return Nothing
  End Function



#Region " DataRow "
  <Extension>
  <DebuggerStepThrough>
  Public Function GetTrimStringOrEmpty(dr As DataRow, field As String) As String
    Return If(dr.Field(Of String)(field), "").Trim
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetTrimStringOrEmpty(dr As DataRow, field As String, version As DataRowVersion) As String
    Return If(dr.Field(Of String)(field, version), "").Trim
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetTrimStringOrEmpty(dr As DataRow, field As Integer) As String
    Return If(dr.Field(Of String)(field), "").Trim
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetTrimStringOrEmpty(dr As DataRow, field As Integer, version As DataRowVersion) As String
    Return If(dr.Field(Of String)(field, version), "").Trim
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetTrimStringOrEmpty(dr As DataRow, dc As DataColumn) As String
    Return dr.GetTrimStringOrEmpty(dc.ColumnName)
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetTrimStringOrEmpty(dr As DataRow, dc As DataColumn, version As DataRowVersion) As String
    Return dr.GetTrimStringOrEmpty(dc.ColumnName, version)
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDecimalOrZero(dr As DataRow, field As String) As Decimal
    Return dr.Field(Of Decimal?)(field).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDecimalOrZero(dr As DataRow, field As String, version As DataRowVersion) As Decimal
    Return dr.Field(Of Decimal?)(field, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDecimalOrZero(dr As DataRow, field As Integer) As Decimal
    Return dr.Field(Of Decimal?)(field).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDecimalOrZero(dr As DataRow, field As Integer, version As DataRowVersion) As Decimal
    Return dr.Field(Of Decimal?)(field, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDecimalOrZero(dr As DataRow, dc As DataColumn) As Decimal
    Return dr.Field(Of Decimal?)(dc.ColumnName).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDecimalOrZero(dr As DataRow, dc As DataColumn, version As DataRowVersion) As Decimal
    Return dr.Field(Of Decimal?)(dc.ColumnName, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetIntegerOrZero(dr As DataRow, field As String) As Integer
    Return dr.Field(Of Integer?)(field).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetIntegerOrZero(dr As DataRow, field As String, version As DataRowVersion) As Integer
    Return dr.Field(Of Integer?)(field, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetIntegerOrZero(dr As DataRow, field As Integer) As Integer
    Return dr.Field(Of Integer?)(field).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetIntegerOrZero(dr As DataRow, field As Integer, version As DataRowVersion) As Integer
    Return dr.Field(Of Integer?)(field, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetIntegerOrZero(dr As DataRow, dc As DataColumn) As Integer
    Return dr.Field(Of Integer?)(dc.ColumnName).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetIntegerOrZero(dr As DataRow, dc As DataColumn, version As DataRowVersion) As Integer
    Return dr.Field(Of Integer?)(dc.ColumnName, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDateOrMin(dr As DataRow, field As String) As Date
    Return dr.Field(Of DateTime?)(field).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDateOrMin(dr As DataRow, field As String, version As DataRowVersion) As Date
    Return dr.Field(Of DateTime?)(field, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDateOrMin(dr As DataRow, field As Integer) As Date
    Return dr.Field(Of DateTime?)(field).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDateOrMin(dr As DataRow, field As Integer, version As DataRowVersion) As Date
    Return dr.Field(Of DateTime?)(field, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDateOrMin(dr As DataRow, dc As DataColumn) As Date
    Return dr.Field(Of DateTime?)(dc.ColumnName).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDateOrMin(dr As DataRow, dc As DataColumn, version As DataRowVersion) As Date
    Return dr.Field(Of DateTime?)(dc.ColumnName, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetBooleanOrFalse(dr As DataRow, field As String) As Boolean
    Return dr.Field(Of Boolean?)(field).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetBooleanOrFalse(dr As DataRow, field As String, version As DataRowVersion) As Boolean
    Return dr.Field(Of Boolean?)(field, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetBooleanOrFalse(dr As DataRow, field As Integer) As Boolean
    Return dr.Field(Of Boolean?)(field).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetBooleanOrFalse(dr As DataRow, field As Integer, version As DataRowVersion) As Boolean
    Return dr.Field(Of Boolean?)(field, version).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetBooleanOrFalse(dr As DataRow, dc As DataColumn) As Boolean
    Return dr.Field(Of Boolean?)(dc.ColumnName).GetValueOrDefault
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetBooleanOrFalse(dr As DataRow, dc As DataColumn, version As DataRowVersion) As Boolean
    Return dr.Field(Of Boolean?)(dc.ColumnName, version).GetValueOrDefault
  End Function

  <Extension()>
  <DebuggerStepThrough()>
  Public Function UpdateFieldIfChanged(ByVal dr As DataRow, ByVal fieldToUpdate As String, ByVal value As String) As Boolean
    If (dr.RowState And DataRowState.Deleted) <> 0 Then
      If Debugger.IsAttached Then Debugger.Break()
      Return False
    End If
    Dim curVal As String = dr.Field(Of String)(fieldToUpdate)
    If curVal Is Nothing Then
      If value IsNot Nothing Then
        dr(fieldToUpdate) = value
        Return True
      End If
    Else
      If value Is Nothing Then
        dr(fieldToUpdate) = DBNull.Value
        Return True
      ElseIf value <> curVal Then
        dr(fieldToUpdate) = value
        Return True
      End If
    End If
    Return False
  End Function

  <Extension()>
  <DebuggerStepThrough()>
  Public Sub UpdateFieldIfChanged(Of T As {Structure, IEquatable(Of T)})(ByVal dr As DataRow, ByVal fieldToUpdate As String, ByVal value As Nullable(Of T))
    If (dr.RowState And DataRowState.Deleted) <> 0 Then
      If Debugger.IsAttached Then Debugger.Break()
      Return
    End If
    Dim curVal As Nullable(Of T) = dr.Field(Of Nullable(Of T))(fieldToUpdate)
    If curVal.HasValue Then
      If value.HasValue Then
        If curVal.Value.Equals(value.Value) = False Then
          dr(fieldToUpdate) = value.Value
        End If
      Else
        dr(fieldToUpdate) = DBNull.Value
      End If
    Else
      If value.HasValue Then
        dr(fieldToUpdate) = value
      End If
    End If
  End Sub



#End Region

#Region " SqlDataReader "

  <Extension>
  <DebuggerStepThrough>
  Public Function GetTrimStringOrEmpty(reader As SqlDataReader, field As Integer) As String
    Return If(reader.IsDBNull(field), "", reader.GetString(field).Trim)
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDecimalOrZero(reader As SqlDataReader, field As Integer) As Decimal
    Return If(reader.IsDBNull(field), 0, reader.GetDecimal(field))
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetIntegerOrZero(reader As IDataReader, field As Integer) As Integer
    Return If(reader.IsDBNull(field), 0, reader.GetInt32(field))
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDateOrMin(reader As SqlDataReader, field As Integer) As Date
    Return If(reader.IsDBNull(field), Date.MinValue, reader.GetDateTime(field))
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetBooleanOrFalse(reader As SqlDataReader, field As Integer) As Boolean
    Return If(reader.IsDBNull(field), False, reader.GetBoolean(field))
  End Function
#End Region

#Region " SqlCommand "
  <Extension>
  <DebuggerStepThrough>
  Public Function GetTrimStringOrEmpty(cmd As SqlCommand) As String
    Return If(TryCast(cmd.ExecuteScalar, String), "").Trim
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDecimalOrZero(cmd As SqlCommand) As Decimal
    Dim o As Object = cmd.ExecuteScalar
    If TypeOf (o) Is Decimal Then
      Return CDec(o)
    Else
      Return 0
    End If
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetIntegerOrZero(cmd As SqlCommand) As Integer
    Dim o As Object = cmd.ExecuteScalar
    If TypeOf (o) Is Integer Then
      Return CInt(o)
    Else
      Return 0
    End If
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDateOrMin(cmd As SqlCommand) As Date
    Dim o As Object = cmd.ExecuteScalar
    If TypeOf (o) Is Date Then
      Return CDate(o)
    Else
      Return Date.MinValue
    End If
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetBooleanOrFalse(cmd As SqlCommand) As Boolean
    Dim o As Object = cmd.ExecuteScalar
    If TypeOf (o) Is Boolean Then
      Return CBool(o)
    Else
      Return False
    End If
  End Function

#End Region

#Region " String "

  <Extension>
  <DebuggerStepThrough>
  Public Function IsNotIn(lhs As String, ParamArray vals() As String) As Boolean
    Return Not lhs.IsIn(vals)
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function IsNotIn(lhs As String, comparisonType As StringComparison, ParamArray vals() As String) As Boolean
    Return Not lhs.IsIn(StringComparison.InvariantCultureIgnoreCase, vals)
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function IsIn(lhs As String, ParamArray vals() As String) As Boolean
    Return lhs.IsIn(StringComparison.InvariantCultureIgnoreCase, vals)
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function IsIn(lhs As String, comparisonType As StringComparison, ParamArray vals() As String) As Boolean
    For Each s As String In vals
      If lhs.Equals(s, comparisonType) Then Return True
    Next
    Return False
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function TrimToLength(lhs As String, length As Integer) As String
    If lhs Is Nothing Then Return Nothing
    lhs = lhs.Trim
    If lhs.Length > length Then
      lhs = lhs.Substring(0, length)
    End If
    Return lhs.Trim
  End Function

#End Region

#Region " ToDbObj "
  <Extension>
  <DebuggerStepThrough>
  Public Function ToDbObj(val As Date?) As Object
    If val.HasValue AndAlso val.Value <> Date.MinValue Then ' 2015.1.7 minvalue
      Return val.Value
    Else
      Return DBNull.Value
    End If
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function ToDbObj(val As Decimal?) As Object
    If val.HasValue Then
      Return val.Value
    Else
      Return DBNull.Value
    End If
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function ToDbObj(val As Integer?) As Object
    If val.HasValue Then
      Return val.Value
    Else
      Return DBNull.Value
    End If
  End Function
#End Region

#Region " DataSet "
  <Extension>
  <DebuggerStepThrough>
  Public Function TryGetTable(ds As DataSet, tableName As String) As DataTable
    If ds.Tables.Contains(tableName) Then
      Return ds.Tables(tableName)
    Else
      Return Nothing
    End If
  End Function
#End Region

#Region " DataRowView "

  <Extension>
  <DebuggerStepThrough>
  Public Function GetTrimStringOrEmpty(drv As DataRowView, field As String) As String
    Return If(TryCast(drv(field), String), "").Trim
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDecimalOrZero(drv As DataRowView, field As String) As Decimal
    If drv(field) Is DBNull.Value Then
      Return 0
    Else
      Return DirectCast(drv(field), Decimal)
    End If
  End Function


  <Extension>
  <DebuggerStepThrough>
  Public Function GetIntegerOrZero(drv As DataRowView, field As String) As Integer
    If drv(field) Is DBNull.Value Then
      Return 0
    Else
      Return DirectCast(drv(field), Integer)
    End If
  End Function

  <Extension>
  <DebuggerStepThrough>
  Public Function GetDateOrMin(drv As DataRowView, field As String) As Date
    If drv(field) Is DBNull.Value Then
      Return Date.MinValue
    Else
      Return DirectCast(drv(field), Date)
    End If
  End Function


  <Extension>
  <DebuggerStepThrough>
  Public Function GetBooleanOrFalse(drv As DataRowView, field As String) As Boolean
    If drv(field) Is DBNull.Value Then
      Return False
    Else
      Return DirectCast(drv(field), Boolean)
    End If
  End Function

#End Region

#Region " DataTable "
  ' caller can AcceptChanges!
  <Extension>
  Public Sub TrimTable(dt As DataTable)
    Dim stringCols As New List(Of DataColumn)
    Dim stringType As Type = GetType(String)
    For Each dc As DataColumn In dt.Columns
      If dc.DataType Is stringType Then
        stringCols.Add(dc)
      End If
    Next
    For Each dr As DataRow In dt.Rows
      For Each dc As DataColumn In stringCols
        dr(dc) = dr.GetTrimStringOrEmpty(dc)
      Next
    Next
  End Sub
#End Region



End Module
