Imports System.IO
Imports FlaUI.Core
Imports FlaUI.UIA3
Imports NUnit.Framework

Namespace Laila.Shell.Tests.UI

    Public Class Tests



        <Test>
        Public Sub Test1()
            Dim testDir = TestContext.CurrentContext.TestDirectory
            Dim appPath = Path.GetFullPath(Path.Combine(
                testDir,
                "..", "..", "..", "..",    ' up from /bin/x64/Debug/net6.0/
                "Laila.Shell.SampleApp",
                "bin", "x64", "Debug", "net6.0-windows7.0",
                "Laila.Shell.SampleApp2.exe"
            ))

            Using app = Application.Launch(appPath)
                Using automation = New UIA3Automation()
                    Threading.Thread.Sleep(2000)
                    Dim window = app.GetMainWindow(automation)
                    Dim listBox = window.FindFirstDescendant(Function(c) c.ByAutomationId("tabControl"))
                    Dim tabItem = listBox.FindFirstDescendant(Function(c) c.ByAutomationId("TabGrid"))
                    Dim tabItem2 = listBox.FindAllChildren()
                    Dim treeView = listBox.FindFirstDescendant(Function(c) c.ByAutomationId("PART_ItemsHolder"))
                    Dim tv = listBox.FindAllNested(automation.ConditionFactory.ByClassName("Grid"))

                    Assert.IsTrue(Not listBox Is Nothing)
                End Using
            End Using
        End Sub

    End Class

End Namespace