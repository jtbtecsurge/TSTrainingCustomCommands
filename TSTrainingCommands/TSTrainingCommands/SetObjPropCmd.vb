Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle.Services
Imports Ingr.SP3D.Route.Middle

Public Class SetObjPropCmd : Inherits BaseModalCommand

    Public Overrides Sub OnStart(instanceId As Integer, argument As Object)
        Dim oSelectSet As SelectSet = ClientServiceProvider.SelectSet

        If oSelectSet.SelectedObjects.Count = 0 Then
            MsgBox("No Objects Selected")
        Else
            'MsgBox(oSelectSet.SelectedObjects(0).ToString)
            Dim oBO As BusinessObject = oSelectSet.SelectedObjects(0)
            oBO.SetPropertyValue("Inuslated Pipeline", "IJPipelineSystem", "Description")
            oBO.SetPropertyValue("OMV_001", "IJPipelineSystem", "SequenceNumber")
            oBO.SetPropertyValue("OMV_Pipe009", "IJNamedItem", "Name")
            'Read the  data from Excel or Text file
            MiddleServiceProvider.TransactionMgr.Compute()
            MiddleServiceProvider.TransactionMgr.Commit("Set value")
        End If


    End Sub

End Class
