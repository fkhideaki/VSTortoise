Imports System.Text
Imports TortoiseCore
Imports Microsoft.VisualStudio.TestTools.UnitTesting

<TestClass()> Public Class UnitTest1

	<TestMethod()> Public Sub TPTypeCehckTest()
		Dim t As TPTypeCehck.EngineType = TPTypeCehck.GetEngineType("E:\ws\git\VSPlugins\TGitAddin")
		Assert.AreEqual(TPTypeCehck.EngineType.Git, t)
	End Sub

End Class