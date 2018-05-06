' Application-level events, such as Startup, Exit, and DispatcherUnhandledException
' can be handled in this file.
Imports System.Data
Imports System.Data.SqlClient
Imports System.Globalization
Imports System.Windows.Threading

Class Application

  Private Sub Application_Startup(sender As Object, e As StartupEventArgs) Handles Me.Startup

    ' we have e.args which is the command line args as a string array.
    ' The are parsed out into the local variable args, which is a dictionary of string to string, so e.g. args("CONO") might be "1"
    ' We save this dic in a global resource dictionary(string, object) with key "CommandLineArgs". If there's some arg without an = then it's value will be nothing. 
    ' We also pull out / parse frequently used stuff like CONO, CO_NO, Locale, Bespoke into entries in the global resource dictionary.
    ' So elsewhere in the program we will mostly be looking in the resources dic, but if there's some weird command line arg we can dig furhter into CommandLineArgs.
    Dim args = ParseArgs(e.Args)


    SetENV(args) ' required, might throw. Then where do we go...
    SetCono(args)
    SetConnectionString(args)
    ' decision descisions - use frequently used code, and all its references or just keep it local?
    SetLocale()
    SetBespoke()
    SetOPR()
  End Sub

  Private Function ParseArgs(commandLineArgs() As String) As Dictionary(Of String, String)
    Dim args As New Dictionary(Of String, String)
    If commandLineArgs IsNot Nothing Then
      For Each s As String In commandLineArgs
        Dim indexOfEquals = s.IndexOf("="c)
        If indexOfEquals = -1 Then
          args.Add(s, Nothing)
        Else
          args.Add(s.Substring(0, indexOfEquals), s.Substring(indexOfEquals + 1, s.Length - indexOfEquals - 1))
        End If
      Next
    End If
    Resources("CommandLineArgs") = args ' Application.Current.Resouces = a global dictionary(of String, object)
    Return args
  End Function

  Private Sub SetENV(args As Dictionary(Of String, String))
    Dim env As String = Nothing
    If args.TryGetValue("ENV", env) = False OrElse String.IsNullOrWhiteSpace(env) Then
      Throw New Exception("ENV was not specified on the command line")
    End If
    Resources("ENV") = env.Trim

    Dim env_ As String = Nothing
    If args.TryGetValue("ENV_", env_) = False Then
      env_ = env ' environment with spaces should be defined on command line if required, otw defaults to ENV. Used by tbred printing routines.
    End If
    Resources("ENV") = env_.Trim
  End Sub

  Private Sub SetCono(args As Dictionary(Of String, String))
    ' cono is optional, defaulting to 1
    Dim cono As String = Nothing
    Dim icono As Integer = 0
    If Not args.TryGetValue("CONO", cono) OrElse String.IsNullOrWhiteSpace(cono) OrElse Integer.TryParse(cono, icono) = False Then
      icono = 1
    End If
    Resources("CONO") = icono.ToString
    Resources("CO_NO") = icono.ToString.PadLeft(2, "0"c) ' some tables have the company number padded out to 2 d.p.
  End Sub

  Private Sub SetConnectionString(args As Dictionary(Of String, String))
    ' hmm. well, for now...
    ' Resources("ConnectionString") = "Data Source=172.20.0.6\SQL2008;Initial Catalog=JulianDev;User ID=jdmsql;Password=coughsyrupgreenhouse;" ' TODO - create this properly
    Resources("ConnectionString") = "Data Source=DESKTOP-K96C0T9\CONCEPTDB;Initial Catalog=JulianDev;Integrated Security=True"
  End Sub

  Private Sub SetLocale()
    Resources("Locale") = If(CultureInfo.CurrentCulture.EnglishName.Contains("United States"), "US", "")
  End Sub

  Private Sub SetOpr()
    Resources("OPR_CODE") = "julian.mcf"

  End Sub

  Private Sub SetBespoke()
    Dim cs = Resources.GetString("ConnectionString")
    If String.IsNullOrWhiteSpace(cs) Then Throw New Exception("Connection string was not initialised")
    Using conn As New SqlConnection(cs)
      Try
        conn.Open()
      Catch ex As Exception
        Throw New Exception("Error opening sql conneciton", ex)
      End Try
      Try
        Using cmd As New SqlCommand("SELECT BESPOKE_LIB_NAME FROM dbo.[SSCOMP] WHERE CO_NO=@CO_NO", conn)
          cmd.Parameters.Add("@CO_NO", SqlDbType.VarChar, 2).Value = Resources.GetString("CO_NO")
          Dim result = cmd.ExecuteScalar
          Resources("Bespoke") = If(TryCast(result, String), "").Trim
        End Using
      Catch ex As Exception
        ' 1st sql run by the program.
        Throw New Exception("Error executing sql command to determine bespoke flag." & ex.Message)
      End Try
    End Using
  End Sub

  Private Sub Application_DispatcherUnhandledException(sender As Object, e As DispatcherUnhandledExceptionEventArgs) Handles Me.DispatcherUnhandledException
#If DEBUG Then
    e.Handled = True ' let vs handle it
#Else
    ShowException(e)
#End If
  End Sub

  Private Sub ShowException(e As DispatcherUnhandledExceptionEventArgs)
    e.Handled = True
    MessageBox.Show(e.Exception.Message, "Application Error", MessageBoxButton.OK)
    Application.Current.Shutdown()
  End Sub

End Class
