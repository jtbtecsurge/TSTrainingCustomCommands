Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Route.Middle
Imports Ingr.SP3D.Common.Middle.Interfaces

Public Class SetObjPropPPL : Inherits BaseModalCommand

    Private Const Name As String = "Enter new Name for the selected object:"
    Private Const Desc As String = "Enter new Description for the selected object:"
    Public Overrides Sub OnStart(instanceId As Integer, argument As Object)
        Dim oSelectSet As SelectSet = ClientServiceProvider.SelectSet
        Dim newName As String
        Dim newDesc As String

        If oSelectSet.SelectedObjects.Count = 0 Then
            MsgBox("No objects selected.")
            Return
        Else
            ' Ask user for new name
            newName = InputBox(Name, "Rename Object")
            newDesc = InputBox(Desc, "Change Object Description")
            ' Cancel if user left it blank or clicked Cancel
            If String.IsNullOrWhiteSpace(newName) Or String.IsNullOrWhiteSpace(newDesc) Then
                MsgBox("Operation cancelled.")
                Return
            End If
            ' Apply name to each selected object
            For Each obj As BusinessObject In oSelectSet.SelectedObjects
                obj.SetPropertyValue(newName, "IJNamedItem", "Name")
                obj.SetPropertyValue(newDesc, "IJPipelineSystem", "Description")
            Next
            ' Commit the changes
            MiddleServiceProvider.TransactionMgr.Compute()
            MiddleServiceProvider.TransactionMgr.Commit("Rename Object(s)")
            MsgBox("Object(s) renamed successfully.")
        End If
    End Sub
End Class
