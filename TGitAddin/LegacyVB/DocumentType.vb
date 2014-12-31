Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics

Public Module DocumentType

	Public Enum DOCTYPE
		CPP
		VB
		CS
		TEXT
		UNKNOWN
	End Enum


	Public Function GetCurrentDocumentType(DTE As DTE2) As DOCTYPE
		Dim ext As String = IO.Path.GetExtension(DTE.ActiveDocument.Name).ToLower()

		If (ext = ".c") Then Return DOCTYPE.CPP
		If (ext = ".cpp") Then Return DOCTYPE.CPP
		If (ext = ".h") Then Return DOCTYPE.CPP

		If (ext = ".cs") Then Return DOCTYPE.CS

		If (ext = ".vb") Then Return DOCTYPE.VB

		If (ext = ".txt") Then Return DOCTYPE.TEXT

		Return DOCTYPE.UNKNOWN
	End Function


	Public Function GetCurrentDocPropertyPageName(DTE As DTE2) As String
		Select Case GetCurrentDocumentType(DTE)
			Case DOCTYPE.CPP : Return "C/C++"
			Case DOCTYPE.CS : Return "CSharp"
			Case DOCTYPE.VB : Return "Basic"
			Case DOCTYPE.TEXT : Return "Basic"
			Case DOCTYPE.UNKNOWN : Throw New Exception("不明なドキュメント")
		End Select

		Return ""
	End Function
End Module
