Imports Data
Imports System.Data
Imports System.Diagnostics
Imports System.Web.UI

Partial Class CodeGenPage
    Inherits System.Web.UI.Page

    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs) Handles Me.Load

        'Dim code As New NLTCodeGenerator(True, True)
        'code.BuildClassAndInterfaceFiles()

    End Sub

    Protected Sub ListBox1_SelectedIndexChanged(ByVal sender As Object, ByVal e As System.EventArgs) Handles ListBox1.SelectedIndexChanged

        Dim tableList As New SortedList
        Dim itemCount As Integer = 0
        For Each item As ListItem In ListBox1.Items
            If item IsNot Nothing Then
                If item.Selected = True Then
                    tableList.Add(itemCount, item.Text)
                    itemCount += 1
                End If
            End If
        Next

        If itemCount > 0 Then
            Dim codePreviewer As New NewLeaf.CodeGenerator
            Dim preview As String = codePreviewer.BuildClassAndInterfaceFilePreviewer(tableList)
            Me.PanelCode.Controls.Clear()
            Me.PanelCode.Controls.Add(New LiteralControl(preview))
        Else
            Me.PanelCode.Controls.Clear()
            Me.PanelCode.Controls.Add(New LiteralControl("Select a table to view the code generation"))
        End If
    End Sub

    Protected Sub lbtnGenerateCode_Click(ByVal sender As Object, ByVal e As System.EventArgs) Handles lbtnGenerateCode.Click

        Dim tableList As New SortedList
        Dim itemCount As Integer = 0
        For Each item As ListItem In ListBox1.Items
            If item IsNot Nothing Then
                If item.Selected = True Then
                    tableList.Add(itemCount, item.Text)
                    itemCount += 1
                End If
            End If
        Next

        If itemCount > 0 Then
            Dim codegen As NewLeaf.CodeGenerator
            If Me.chkOverwriteFiles.Checked = True Then
                codegen = New NewLeaf.CodeGenerator(True, True)
            Else
                codegen = New NewLeaf.CodeGenerator(True, False)
            End If
            Dim success As Boolean = codegen.BuildClassAndInterfaceFiles(tableList)
            Me.PanelCode.Controls.Clear()
            If success = True Then
                Dim preview As String = codegen.BuildClassAndInterfaceFilePreviewer(tableList)
                Me.PanelCode.Controls.Add(New LiteralControl("The following code was created in their respective files. To view them, stop and refresh the project files.<br/><br/>" & preview))
            Else
                Me.PanelCode.Controls.Add(New LiteralControl("An error occurred while attempting to build the project files."))
            End If
        Else
            Me.PanelCode.Controls.Clear()
            Me.PanelCode.Controls.Add(New LiteralControl("Select a table to view the code generation"))
        End If

    End Sub


    '    Dim ifolder As System.IO.DirectoryInfo = Nothing
    '    Dim cfolder As System.IO.DirectoryInfo = Nothing
    '    cfolder = New System.IO.DirectoryInfo(My.Request.PhysicalApplicationPath & "\App_Code\BLL\")
    '    ifolder = New System.IO.DirectoryInfo(My.Request.PhysicalApplicationPath & "\App_Code\BLL\Interfaces\")

    '    Me.PanelCode.Controls.Clear()

    '    If ifolder.Exists Then
    '        Me.PanelCode.Controls.Add(New LiteralControl("Interface Files<br/>"))
    '        For Each file As String In System.IO.Directory.GetFiles(ifolder.FullName, "*.vb")
    '            Me.PanelCode.Controls.Add(New LiteralControl(file & "<br/>"))
    '        Next
    '    End If

    '    Me.PanelCode.Controls.Add(New LiteralControl("<br/>"))

    '    If cfolder.Exists Then
    '        Me.PanelCode.Controls.Add(New LiteralControl("Class Files<br/>"))
    '        For Each file As String In System.IO.Directory.GetFiles(cfolder.FullName, "*.vb")
    '            Me.PanelCode.Controls.Add(New LiteralControl(file & "<br/>"))
    '        Next
    '    End If

End Class

