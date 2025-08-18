Imports System.Drawing
Imports System.Reflection
Imports System.Windows.Forms
Imports Ingr.SP3D.Common.Client
Imports Ingr.SP3D.Common.Client.Services
Imports Ingr.SP3D.Common.Middle
Imports Ingr.SP3D.Common.Middle.Services
Imports Intergraph.CommonToolkit.Client.Interfaces

Public Class EditPropsForm
    Inherits Form

    Private TargetObject As BusinessObject
    Private txtName As TextBox
    Private txtDesc As TextBox
    Private chkInsl As CheckBox
    Private txtSqnm As TextBox
    Private txtRmnm As TextBox
    Private txtEDes As TextBox
    Private txtSeId As TextBox
    Private txtTag As TextBox
    Private txtCDes As TextBox
    Private txtISNo As TextBox
    Private txtSDes As TextBox
    Private txtTDes As TextBox
    Private txtUmdt As TextBox
    Private txtAsMa As TextBox
    Private txtPiMa As TextBox
    Private txtStCo As TextBox
    Private txtWoCe As TextBox

    ' This code creates the form and gathers data/input from the user
    Public Sub New(obj As BusinessObject)
        Dim currentTop As Integer = 40
        Dim spacing As Integer = 30
        TargetObject = obj

        Dim lblType As New Label With {
            .Text = "Object Type: " & GetObjectTypeName(obj),
            .Top = 10, .Left = 10, .Width = 360
        }
        Me.Controls.Add(lblType)

        currentTop = AvailableProperty(obj, currentTop, spacing)
        currentTop = ButtonApply(obj, currentTop, spacing)

        ChangeLightAccentColor()
        Me.Icon = New Icon("C:\Users\s3d\Downloads\document.ico")
        Me.Text = "Edit Properties"
        Me.Width = 325
        Me.Height = 75 + currentTop
        Me.StartPosition = FormStartPosition.CenterScreen
    End Sub

    Private Sub ChangeLightAccentColor()
        Dim rnd As New Random

        Dim r As Integer = rnd.Next(180, 256)
        Dim g As Integer = rnd.Next(180, 256)
        Dim b As Integer = rnd.Next(180, 256)

        Dim randomLightColor As Color = Color.FromArgb(r, g, b)
        Me.BackColor = randomLightColor

        For Each ctrl As Control In Me.Controls
            ctrl.BackColor = randomLightColor
        Next
    End Sub

    ' Method for Button Apply 
    Private Function ButtonApply(obj As BusinessObject, currentTop As Integer, spacing As Integer) As Integer
        ' Button 
        If {("IJNamedItem", "Name"),
            ("IJPipelineSystem", "Description"),
            ("IJOccInsulation", "IsInsulated"),
            ("IURoomNo", "RoomNo"),
            ("IJEquipment", "Description"),
            ("IJPipelineSystem", "SequenceNumber"),
            ("IJSequence", "Id"),
            ("IJRtePipePathFeat", "Tag"),
            ("ISPGCoordinateSystemProperties", "Description"),
            ("IJRtePipePart", "IsoSheetNo"),
            ("IJSmartInteropCommonAttribute", "Description"),
            ("ISIOTag", "TagTypeDescription"),
            ("IJSmartInteropUnmappedData", "UnmappedData"),
            ("ISIOStructDetailingProps", "AssemblyMark"),
            ("ISIOStructDetailingProps", "PieceMark"),
            ("ISIOStructState", "StageCode"),
            ("ISIOStructuralPart", "WorkCenter")
             }.Any(Function(p) PropertyExists(obj, p.Item1, p.Item2)) Then
            Dim btnApply As New Button With {.Text = "Apply", .Top = currentTop, .Left = 120}
            AddHandler btnApply.Click, AddressOf ApplyChanges
            Me.Controls.Add(btnApply)
        Else
            Dim lblNP As New Label With {.Text = "No Editable Property/s", .Top = currentTop, .Left = 10, .Width = 200}
            currentTop += spacing
            Dim btnNA As New Button With {.Text = "Ok", .Top = currentTop, .Left = 120}
            AddHandler btnNA.Click, AddressOf Close
            Me.Controls.Add(lblNP)
            Me.Controls.Add(btnNA)
        End If
        Return currentTop
    End Function

    'Method where it checks and shows the availbale Property that can be Modified
    Private Function AvailableProperty(obj As BusinessObject, currentTop As Integer, spacing As Integer) As Integer
        ' Name
        If PropertyExists(obj, "IJNamedItem", "Name") Then
            Dim lbl1 As New Label With {.Text = "Name:", .Top = currentTop, .Left = 0}
            txtName = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("IJNamedItem", "Name"))}
            Me.Controls.Add(lbl1)
            Me.Controls.Add(txtName)
            currentTop += spacing

        End If

        ' Description
        If PropertyExists(obj, "IJPipelineSystem", "Description") Then
            Dim lbl2 As New Label With {.Text = "Description:", .Top = currentTop, .Left = 0}
            txtDesc = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("IJPipelineSystem", "Description"))}
            Me.Controls.Add(lbl2)
            Me.Controls.Add(txtDesc)
            currentTop += spacing

        End If

        ' Sequence Number
        If PropertyExists(obj, "IJPipelineSystem", "SequenceNumber") Then
            Dim lbl3 As New Label With {.Text = "Sequence Number:", .Top = currentTop, .Left = 0}
            txtSqnm = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("IJPipelineSystem", "SequenceNumber"))}
            Me.Controls.Add(lbl3)
            Me.Controls.Add(txtSqnm)
            currentTop += spacing
        End If

        ' Insulation
        If PropertyExists(obj, "IJOccInsulation", "IsInsulated") Then
            Dim isInsl As Boolean
            Dim value As String = Convert.ToString(obj.GetPropertyValue("IJOccInsulation", "IsInsulated"))
            If String.Equals(value, "True", StringComparison.OrdinalIgnoreCase) Then
                isInsl = True
                Dim lbl4 As New Label With {.Text = "Insulation:", .Top = currentTop, .Left = 0}
                chkInsl = New CheckBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                            .Checked = isInsl}
                Me.Controls.Add(lbl4)
                Me.Controls.Add(chkInsl)
                currentTop += spacing
            ElseIf String.Equals(value, "False", StringComparison.OrdinalIgnoreCase) Then
                isInsl = False
                Dim lbl4 As New Label With {.Text = "Allow Insulation:", .Top = currentTop, .Left = 0}
                chkInsl = New CheckBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                            .Checked = isInsl}
                Me.Controls.Add(lbl4)
                Me.Controls.Add(chkInsl)
                currentTop += spacing
            End If
        End If

        ' Room Number
        If PropertyExists(obj, "IURoomNo", "RoomNo") Then
            Dim lbl5 As New Label With {.Text = "Room Number:", .Top = currentTop, .Left = 0}
            txtRmnm = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("IURoomNo", "RoomNo"))}
            Me.Controls.Add(lbl5)
            Me.Controls.Add(txtRmnm)
            currentTop += spacing
        End If

        ' Equipment Description
        If PropertyExists(obj, "IJEquipment", "Description") Then
            Dim lbl6 As New Label With {.Text = "Description:", .Top = currentTop, .Left = 0}
            txtEDes = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("IJEquipment", "Description"))}
            Me.Controls.Add(lbl6)
            Me.Controls.Add(txtEDes)
            currentTop += spacing
        End If

        ' Sequence Id
        If PropertyExists(obj, "IJSequence", "Id") Then
            Dim lbl7 As New Label With {.Text = "Sequence ID:", .Top = currentTop, .Left = 0}
            txtSeId = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("IJSequence", "Id"))}
            Me.Controls.Add(lbl7)
            Me.Controls.Add(txtSeId)
            currentTop += spacing
        End If

        ' Tag
        If PropertyExists(obj, "IJRtePipePathFeat", "Tag") Then
            Dim lbl8 As New Label With {.Text = "Tag:", .Top = currentTop, .Left = 0}
            txtSeId = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("IJRtePipePathFeat", "Tag"))}
            Me.Controls.Add(lbl8)
            Me.Controls.Add(txtTag)
            currentTop += spacing
        End If

        ' Coordinates Description
        If PropertyExists(obj, "ISPGCoordinateSystemProperties", "Description") Then
            Dim lbl9 As New Label With {.Text = "Description:", .Top = currentTop, .Left = 0}
            txtCDes = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("ISPGCoordinateSystemProperties", "Description"))}
            Me.Controls.Add(lbl9)
            Me.Controls.Add(txtCDes)
            currentTop += spacing
        End If

        ' Isometric Sheet Number
        If PropertyExists(obj, "IJRtePipePart", "IsoSheetNo") Then
            Dim lbl10 As New Label With {.Text = "Iso Sheet No:", .Top = currentTop, .Left = 0}
            txtISNo = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("IJRtePipePart", "IsoSheetNo"))}
            Me.Controls.Add(lbl10)
            Me.Controls.Add(txtISNo)
            currentTop += spacing
        End If

        ' Smart Interop Description
        If PropertyExists(obj, "IJSmartInteropCommonAttribute", "Description") Then
            Dim lbl11 As New Label With {.Text = "Description:", .Top = currentTop, .Left = 0}
            txtSDes = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("IJSmartInteropCommonAttribute", "Description"))}
            Me.Controls.Add(lbl11)
            Me.Controls.Add(txtSDes)
            currentTop += spacing
        End If

        ' Tag Type Description
        If PropertyExists(obj, "ISIOTag", "TagTypeDescription") Then
            Dim lbl12 As New Label With {.Text = "Tag Type Description:", .Top = currentTop, .Left = 0}
            txtTDes = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("ISIOTag", "TagTypeDescription"))}
            Me.Controls.Add(lbl12)
            Me.Controls.Add(txtTDes)
            currentTop += spacing
        End If

        ' Unmapped Data
        If PropertyExists(obj, "IJSmartInteropUnmappedData", "UnmappedData") Then
            Dim lbl13 As New Label With {.Text = "Unmapped Data:", .Top = currentTop, .Left = 0}
            txtUmdt = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("IJSmartInteropUnmappedData", "UnmappedData"))}
            Me.Controls.Add(lbl13)
            Me.Controls.Add(txtUmdt)
            currentTop += spacing
        End If

        ' Assembly Mark
        If PropertyExists(obj, "ISIOStructDetailingProps", "AssemblyMark") Then
            Dim lbl14 As New Label With {.Text = "Assembly Mark:", .Top = currentTop, .Left = 0}
            txtAsMa = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("ISIOStructDetailingProps", "AssemblyMark"))}
            Me.Controls.Add(lbl14)
            Me.Controls.Add(txtAsMa)
            currentTop += spacing
        End If

        ' Piece Mark
        If PropertyExists(obj, "ISIOStructDetailingProps", "PieceMark") Then
            Dim lbl15 As New Label With {.Text = "Piece Mark:", .Top = currentTop, .Left = 0}
            txtPiMa = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("ISIOStructDetailingProps", "PieceMark"))}
            Me.Controls.Add(lbl15)
            Me.Controls.Add(txtPiMa)
            currentTop += spacing
        End If

        ' Stage Code
        If PropertyExists(obj, "ISIOStructState", "StageCode") Then
            Dim lbl16 As New Label With {.Text = "Stage Code:", .Top = currentTop, .Left = 0}
            txtStCo = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("ISIOStructState", "StageCode"))}
            Me.Controls.Add(lbl16)
            Me.Controls.Add(txtStCo)
            currentTop += spacing
        End If

        ' Work Center
        If PropertyExists(obj, "ISIOStructuralPart", "WorkCenter") Then
            Dim lbl17 As New Label With {.Text = "Work Center:", .Top = currentTop, .Left = 0}
            txtWoCe = New TextBox With {.Top = currentTop, .Left = 100, .Width = 200,
                                        .Text = Convert.ToString(obj.GetPropertyValue("ISIOStructuralPart", "WorkCenter"))}
            Me.Controls.Add(lbl17)
            Me.Controls.Add(txtWoCe)
            currentTop += spacing
        End If

        Return currentTop
    End Function

    ' This Function Gets The Name of the type of an Object that's been Selected
    Private Function GetObjectTypeName(obj As BusinessObject) As String
        Try
            Dim pi As PropertyInfo = obj.GetType().GetProperty("TypeName", BindingFlags.Public Or BindingFlags.Instance)
            If pi IsNot Nothing Then
                Return Convert.ToString(pi.GetValue(obj, Nothing))
            End If
        Catch
        End Try
        Return obj.GetType().Name
    End Function

    ' This checks if the Property existed on the object selected
    Private Function PropertyExists(obj As BusinessObject, iface As String, prop As String) As Boolean
        Try
            obj.GetPropertyValue(iface, prop)
            Return True
        Catch
            Return False
        End Try
    End Function

    ' This Method Apply the Changes to the Smart 3D Database PS: If Multiple Objects Selected, All Selected Objects' Properties will be modified also.
    ' Only Similar Objects From the First Selected Object Will Be Modified
    Private Sub ApplyChanges(sender As Object, e As EventArgs)
        Try
            Dim oSelectSet As SelectSet = ClientServiceProvider.SelectSet
            If oSelectSet.SelectedObjects.Count = 0 Then Exit Sub
            Dim firstType As Type = oSelectSet.SelectedObjects(0).GetType()

            For Each TargetObject As BusinessObject In oSelectSet.SelectedObjects
                If TargetObject.GetType() Is firstType Then
                    If txtName IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtName.Text) Then
                        TargetObject.SetPropertyValue(txtName.Text, "IJNamedItem", "Name")
                    End If
                    If txtDesc IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtDesc.Text) Then
                        TargetObject.SetPropertyValue(txtDesc.Text, "IJPipelineSystem", "Description")
                    End If
                    If txtSqnm IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtSqnm.Text) Then
                        TargetObject.SetPropertyValue(txtSqnm.Text, "IJPipelineSystem", "SequenceNumber")
                    End If
                    If chkInsl IsNot Nothing AndAlso Not String.IsNullOrEmpty(chkInsl.Checked) Then
                        TargetObject.SetPropertyValue(chkInsl.Checked, "IJOccInsulation", "IsInsulated")
                    End If
                    If txtRmnm IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtRmnm.Text) Then
                        TargetObject.SetPropertyValue(txtRmnm.Text, "IURoomNo", "RoomNo")
                    End If
                    If txtEDes IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtEDes.Text) Then
                        TargetObject.SetPropertyValue(txtEDes.Text, "IJEquipment", "Description")
                    End If
                    If txtSeId IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtSeId.Text) Then
                        TargetObject.SetPropertyValue(txtSeId.Text, "IJSequence", "Id")
                    End If
                    If txtTag IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtTag.Text) Then
                        TargetObject.SetPropertyValue(txtTag.Text, "IJRtePipePathFeat", "Tag")
                    End If
                    If txtCDes IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtCDes.Text) Then
                        TargetObject.SetPropertyValue(txtCDes.Text, "ISPGCoordinateSystemProperties", "Description")
                    End If
                    If txtISNo IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtISNo.Text) Then
                        TargetObject.SetPropertyValue(txtISNo.Text, "IJRtePipePart", "IsoSheetNo")
                    End If
                    If txtSDes IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtSDes.Text) Then
                        TargetObject.SetPropertyValue(txtSDes.Text, "IJSmartInteropCommonAttribute", "Description")
                    End If
                    If txtTDes IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtTDes.Text) Then
                        TargetObject.SetPropertyValue(txtTDes.Text, "ISIOTag", "TagTypeDescription")
                    End If
                    If txtUmdt IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtUmdt.Text) Then
                        TargetObject.SetPropertyValue(txtUmdt.Text, "IJSmartInteropUnmappedData", "UnmappedData")
                    End If
                    If txtAsMa IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtAsMa.Text) Then
                        TargetObject.SetPropertyValue(txtAsMa.Text, "ISIOStructDetailingProps", "AssemblyMark")
                    End If
                    If txtPiMa IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtPiMa.Text) Then
                        TargetObject.SetPropertyValue(txtPiMa.Text, "ISIOStructDetailingProps", "PieceMark")
                    End If
                    If txtStCo IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtStCo.Text) Then
                        TargetObject.SetPropertyValue(txtStCo.Text, "ISIOStructState", "StageCode")
                    End If
                    If txtWoCe IsNot Nothing AndAlso Not String.IsNullOrEmpty(txtWoCe.Text) Then
                        TargetObject.SetPropertyValue(txtWoCe.Text, "ISIOStructuralPart", "WorkCenter")
                    End If
                End If

            Next

            MiddleServiceProvider.TransactionMgr.Compute()
            MiddleServiceProvider.TransactionMgr.Commit("Property Values")
            MessageBox.Show("Properties updated successfully!", "Done")
            Me.Close()
        Catch ex As Exception
            MessageBox.Show("Error: " & ex.Message, "Failed")
        End Try
    End Sub
End Class

' This is the main where it will check if there's been a selected object
Public Class QuickEditFormCmd
    Inherits BaseModalCommand

    Public Overrides Sub OnStart(instanceId As Integer, argument As Object)
        Dim oSelectSet As SelectSet = ClientServiceProvider.SelectSet
        If oSelectSet Is Nothing OrElse oSelectSet.SelectedObjects.Count = 0 Then
            MessageBox.Show("No objects selected.", "Info")
            Exit Sub
        End If

        Dim selObj = oSelectSet.SelectedObjects(0)
        Dim busObj As BusinessObject = TryCast(selObj, BusinessObject)
        If busObj Is Nothing Then
            MessageBox.Show("Selected item is not a BusinessObject.", "Info")
            Exit Sub
        End If

        Dim frm As New EditPropsForm(busObj)
        frm.ShowDialog()
    End Sub
End Class
