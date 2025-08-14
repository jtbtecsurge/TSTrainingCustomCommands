Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Route.Middle

Public Class SetObjPropHS : Inherits BaseModalCommand

    Private Const Name As String = "Enter new Name for the selected object:"
    Public Overrides Sub OnStart(instanceId As Integer, argument As Object)
        Dim oSelectSet As SelectSet = ClientServiceProvider.SelectSet
        Dim newName As String

        If oSelectSet.SelectedObjects.Count = 0 Then
            MsgBox("No objects selected.")
            Return
        Else
            ' Ask user for new name
            newName = InputBox(Name, "Rename Object")
            ' Cancel if user left it blank or clicked Cancel
            If String.IsNullOrWhiteSpace(newName) Then
                MsgBox("Operation cancelled.")
                Return
            End If
            ' Apply name to each selected object
            For Each obj As BusinessObject In oSelectSet.SelectedObjects
                obj.SetPropertyValue(newName, "IJNamedItem", "Name")

            Next
            ' Commit the changes
            MiddleServiceProvider.TransactionMgr.Compute()
            MiddleServiceProvider.TransactionMgr.Commit("Rename Object(s)")
            MsgBox("Object(s) renamed successfully.")
        End If
    End Sub

End Class
