﻿Imports System
Imports EnvDTE
Imports EnvDTE80
Imports EnvDTE100
Imports System.Diagnostics


Public Class TGit
	Inherits Tortoise

	Sub New(parent As DTE2)
		MyBase.New(parent)
	End Sub

	Protected Overrides Function GetTPName() As String
		Return """TortoiseGitProc.exe"""
	End Function
End Class
