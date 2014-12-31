Imports System
Imports EnvDTE
Imports EnvDTE80
Imports System.Diagnostics

Public Module ExtTools

	Private Sub OpenCurrentFrom(DTE As DTE2, ByVal program_name As String)
		Dim filename As String = ""
		If Not get_current_doc(DTE, filename) Then Return
		Dim cmdline As String = """" + program_name + """ """ + filename + """"

		DTE.ExecuteCommand("Tools.Shell", cmdline)
	End Sub

    
	Sub OpenCurrent_SakuraEditor(DTE As DTE2)
		OpenCurrentFrom(DTE, "C:\Program Files\sakura\sakura.exe")
	End Sub
End Module
