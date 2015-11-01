
Imports System.Text

''' <summary>
''' IIS Express 制御
''' </summary>
''' <remarks></remarks>
Public Class IISExpressManager

#Region " Declare "

	''' <summary>
	''' IISExpress.exe パス
	''' </summary>
	''' <remarks></remarks>
	Private Const C_EXE As String = "IIS Express\iisexpress.exe"

	''' <summary>
	''' appcmd.exe パス
	''' </summary>
	''' <remarks></remarks>
	Private Const C_ADDEXE As String = "IIS Express\appcmd.exe"
	Private Const C_PATH As String = " /path:""{0}"""
	Private Const C_VPATH As String = " /vpath:""{0}"""
	Private Const C_PORT As String = " /port:{0}"
	Private Const C_CLR As String = " /clr:{0}"
	Private Const C_SITE As String = " /site:{0}"
	Private Const C_SYSTRAY As String = " /systray:false"
	Private Const C_NTLM As String = " /ntlm"

	Private Shared _iis As Process

#Region " Logging For Log4net "
	''' <summary>Logging For Log4net</summary>
	Private Shared ReadOnly _mylog As log4net.ILog = log4net.LogManager.GetLogger(String.Empty)
#End Region
#End Region

#Region " Property "

	''' <summary>
	''' プロジェクト名
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Property ProjectName As String

	''' <summary>
	''' アプリケーションパス
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Property Path As String

	''' <summary>
	''' 仮想ディレクトリ
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Property VPath As String

	''' <summary>
	''' 割当ポート
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Property Port As String

    ''' <summary>
    ''' アプリケーションで使用する.NETライタイムのバージョン
    ''' </summary>
    ''' <value></value>
    ''' <returns></returns>
    ''' <remarks>Default:v4.0</remarks>
    Public Shared Property Clr As String

	''' <summary>
	''' サイト名
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Property Site As String

	''' <summary>
	''' システムトレイに表示するか
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Property SysTray As Boolean = True

	''' <summary>
	''' 
	''' </summary>
	''' <value></value>
	''' <returns></returns>
	''' <remarks></remarks>
	Public Shared Property Ntlm As Boolean

#End Region

#Region " Method "

	''' <summary>
	''' 開始
	''' </summary>
	''' <remarks></remarks>
	Public Shared Sub Start()
		If _iis IsNot Nothing Then
			Return
		End If

		Dim tt As Threading.Thread = New Threading.Thread(AddressOf _startIIS)
		tt.IsBackground = True
		tt.Start()
	End Sub

	''' <summary>
	''' 終了
	''' </summary>
	''' <remarks></remarks>
	Public Shared Sub [Stop]()
		If _iis Is Nothing Then
			Return
		End If

		'_iis.CloseMainWindow()
		'_iis.Close()
		Try
			_iis.Kill()
			_iis.Dispose()
		Catch ex As Exception
			_mylog.Debug(ex.Message)
		End Try
	End Sub

	''' <summary>
	''' アプリケーションパスの取得
	''' </summary>
	''' <param name="name"></param>
	''' <returns></returns>
	''' <remarks></remarks>
	Private Shared Function _getApplicationPath(ByVal name As String) As String
		Dim solutionFolder As String = System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(System.IO.Path.GetDirectoryName(AppDomain.CurrentDomain.BaseDirectory)))
		Return System.IO.Path.Combine(solutionFolder, name)
	End Function

	''' <summary>
	''' IIS開始
	''' </summary>
	''' <remarks></remarks>
	Private Shared Sub _startIIS()
		Dim psInfo As New ProcessStartInfo()
		Dim programfiles As String
		Dim solutionFolder As String

		programfiles = Environment.GetFolderPath(Environment.SpecialFolder.ProgramFiles)

		'Dim addExe As Process
		'addExe = Process.Start(System.IO.Path.Combine(programfiles, C_ADDEXE), String.Format("delete site http://localhost:{0}/{1}", Port, Site))
		'addExe = Process.Start(System.IO.Path.Combine(programfiles, C_ADDEXE), String.Format("add site /name:""{0}"" /bindings:http://*:{1}: /physicalPath:""{2}""", Site, Port, Path))
		'addExe = Process.Start(System.IO.Path.Combine(programfiles, C_ADDEXE), String.Format("add app /site.name:""{0}-Site"" /path:{0} /physicalPath:""{2}""", Site, Port, Path))

		Dim args As StringBuilder = New StringBuilder

		If String.IsNullOrEmpty(Path) Then
			solutionFolder = _getApplicationPath(ProjectName)
		Else
			solutionFolder = Path
		End If
		args.AppendFormat(C_PATH, solutionFolder)
		If Not String.IsNullOrEmpty(VPath) Then
			args.AppendFormat(C_VPATH, VPath)
		End If
		If Not String.IsNullOrEmpty(Site) Then
			args.AppendFormat(C_SITE, Site)
		End If
		If Not String.IsNullOrEmpty(Port) Then
			args.AppendFormat(C_PORT, Port)
		End If
		If Not String.IsNullOrEmpty(Clr) Then
			args.AppendFormat(C_CLR, Clr)
		End If
		If Not SysTray Then
			args.AppendFormat(C_SYSTRAY)
		End If
		If Ntlm Then
			args.AppendFormat(C_NTLM)
		End If

		psInfo.LoadUserProfile = True
		psInfo.UseShellExecute = False
		psInfo.FileName = System.IO.Path.Combine(programfiles, C_EXE)
		psInfo.Arguments = args.ToString

		_mylog.Debug(psInfo.Arguments)

		Try
			_iis = Process.Start(psInfo)
			_iis.WaitForExit()
		Catch ex As Exception
			_mylog.Error(ex)
			'_iis.CloseMainWindow()
			_iis.Dispose()
		End Try
	End Sub

#End Region

End Class
