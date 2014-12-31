Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics


Public Class IDE_Utility
	Dim DTE As DTE2

	Public Sub New(dte2_instance As DTE2)
		DTE = dte2_instance
	End Sub

	Function GetActivePane()
		Dim ActWindow = DTE.ActiveDocument.ActiveWindow
		Return ActWindow.Object.ActivePane
	End Function

	' デバッグ出力ウィンドウをクリア
	Sub ClearDebugOutputWindow()
		DTE.ExecuteCommand("Edit.ClearOutputWindow")
	End Sub

	' 出力ウィンドウの指定した名前を持つペインを開く。
	' 存在しないペイン名が指定された場合は、自動的に生成される。
	Public Function GetOutputWindowPane(ByVal Name As String, Optional ByVal show As Boolean = True) As OutputWindowPane
		Dim win As Window = DTE.Windows.Item(EnvDTE.Constants.vsWindowKindOutput)
		If show Then win.Visible = True
		Dim ow As OutputWindow = win.Object
		Dim owpane As OutputWindowPane
		Try
			owpane = ow.OutputWindowPanes.Item(Name)
		Catch e As System.Exception
			owpane = ow.OutputWindowPanes.Add(Name)
		End Try
		owpane.Activate()
		Return owpane
	End Function
End Class
