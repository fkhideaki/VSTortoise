Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics

Imports System.CodeDom


Public Module IDECompile

    ' 選択範囲の文字列をソースコードとみなして処理を実行する
	Public Sub CompileSelected(DTE As DTE2)
		Dim str_select As String = DTE.ActiveDocument.Selection.Text

		Dim src As String = ""
		src += "Imports System" + vbNewLine
		src += "Imports Microsoft" + vbNewLine
		src += "Imports Microsoft.VisualBasic" + vbNewLine
		src += "Imports Microsoft.VisualBasic.Constants" + vbNewLine
		src += "Public Class MainClass" + vbNewLine
		src += "  Public Shared out as String = """"" + vbNewLine
		src += "  Public Shared Function GetOut as String" + vbNewLine
		src += "    return out" + vbNewLine
		src += "  End Function " + vbNewLine
		src += "  Private Shared Sub WriteLine(s as String)" + vbNewLine
		src += "    out += s + vbNewLine" + vbNewLine
		src += "  End Sub" + vbNewLine
		src += "  Private Shared Sub Write(s as String)" + vbNewLine
		src += "    out += s" + vbNewLine
		src += "  End Sub" + vbNewLine
		src += "Public Shared Sub MainFunc()" + vbNewLine
		src += str_select + vbNewLine
		src += "End Sub" + vbNewLine
		src += "End Class" + vbNewLine

		Dim compile_result As Compiler.CompilerResults = CompileString(DTE, src)
		If compile_result Is Nothing Then Return

		'コンパイルしたアセンブリを取得
		Dim compiled_asm As System.Reflection.Assembly = compile_result.CompiledAssembly

		' 処理実行
		Dim t As Type = compiled_asm.GetType("MainClass")  'MainClassクラスのTypeを取得
		t.InvokeMember("MainFunc", System.Reflection.BindingFlags.InvokeMethod, Nothing, Nothing, Nothing)
		Dim s As String = DirectCast(t.InvokeMember("GetOut", System.Reflection.BindingFlags.InvokeMethod, Nothing, Nothing, Nothing), String)

		Dim iu As New IDE_Utility(DTE)
		iu.GetOutputWindowPane("IDECompile Result : ").Clear()
		iu.GetOutputWindowPane("IDECompile Result : ").OutputString(s)
	End Sub

    'ソースコードのコンパイルしてアセンブリを生成する
	Private Function CompileString(DTE As DTE2, ByVal SourceCode As String) As Compiler.CompilerResults
		' [VB/CSharp]CodeProviderクラスのインスタンスを作成
		' [VB/CSharp]CodeProviderクラスはSystem.CodeDom.Compiler.CodeDomProviderの派生クラス
		Dim codeProvider As Compiler.CodeDomProvider
		codeProvider = New Microsoft.VisualBasic.VBCodeProvider()

		' 作成したCodeProviderから、ICodeCompilerインターフェイスを取得
		Dim codeCompiler As Compiler.ICodeCompiler
		codeCompiler = codeProvider.CreateCompiler()

		'コンパイルに使用するパラメータ
		Dim params As New Compiler.CompilerParameters()
		With params
			.GenerateExecutable = False
			.IncludeDebugInformation = True
			.GenerateInMemory = True
			'.ReferencedAssemblies.Add("Microsoft.VisualBasic.dll")
		End With

		'コンパイル実行
		Try
			Dim iu As New IDE_Utility(DTE)

			Dim compiled_asm As Compiler.CompilerResults
			compiled_asm = codeCompiler.CompileAssemblyFromSource(params, SourceCode)
			If compiled_asm.Errors.HasErrors = True Then
				'コンパイラからの出力を表示
				iu.GetOutputWindowPane("IDECompile Error : ").Clear()
				For Each msg As String In compiled_asm.Output
					iu.GetOutputWindowPane("IDECompile Error : ").OutputString(msg + vbNewLine)
				Next

				Return Nothing
			End If

			Return compiled_asm
		Catch ex As Exception
			Dim err_str As String
			err_str = "ERROR : " + vbNewLine + vbNewLine + ex.Message
			MsgBox(err_str)
			Return Nothing
		End Try
	End Function

End Module
