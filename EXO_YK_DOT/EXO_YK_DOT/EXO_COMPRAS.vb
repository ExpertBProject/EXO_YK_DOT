Imports System.Xml
Imports SAPbouiCOM
Public Class EXO_COMPRAS
    Inherits EXO_Generales.EXO_DLLBase
    Public Sub New(ByRef general As EXO_Generales.EXO_General, actualizar As Boolean)
        MyBase.New(general, actualizar)

    End Sub

    Public Overrides Function filtros() As EventFilters
        Dim filtrosXML As Xml.XmlDocument = New Xml.XmlDocument
        filtrosXML.LoadXml(objGlobal.Functions.leerEmbebido(Me.GetType(), "XML_FILTROS.xml"))
        Dim filtro As SAPbouiCOM.EventFilters = New SAPbouiCOM.EventFilters()
        filtro.LoadFromXML(filtrosXML.OuterXml)

        Return filtro
    End Function
    Public Overrides Function menus() As XmlDocument
        Return Nothing
    End Function

    Public Overrides Function SBOApp_ItemEvent(ByRef infoEvento As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim res As Boolean = True
        Dim oForm As SAPbouiCOM.Form = SboApp.Forms.Item(infoEvento.FormUID)

        Dim Path As String = ""
        Dim XmlDoc As New System.Xml.XmlDocument
        Dim EXO_Functions As New EXO_BasicDLL.EXO_Generic_Forms_Functions(Me.objGlobal.conexionSAP)

        Try
            If infoEvento.InnerEvent = False Then
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "143", "182"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED
                                    If EventHandler_ItemPressed_After(objGlobal, infoEvento) = False Then
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_FORM_RESIZE

                            End Select
                    End Select
                ElseIf infoEvento.BeforeAction = True Then
                    Select Case infoEvento.FormTypeEx
                        Case "143", "182"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_COMBO_SELECT

                                Case SAPbouiCOM.BoEventTypes.et_CLICK

                                Case SAPbouiCOM.BoEventTypes.et_ITEM_PRESSED

                                Case SAPbouiCOM.BoEventTypes.et_VALIDATE

                                Case SAPbouiCOM.BoEventTypes.et_KEY_DOWN

                                Case SAPbouiCOM.BoEventTypes.et_MATRIX_LINK_PRESSED

                            End Select
                    End Select
                End If
            Else
                If infoEvento.BeforeAction = False Then
                    Select Case infoEvento.FormTypeEx
                        Case "143", "182"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_FORM_VISIBLE

                                Case SAPbouiCOM.BoEventTypes.et_FORM_LOAD
                                    If EventHandler_FORM_LOAD_After(objGlobal, infoEvento) = False Then
                                        GC.Collect()
                                        Return False
                                    End If
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                            End Select
                    End Select
                Else
                    Select Case infoEvento.FormTypeEx
                        Case "143", "182"
                            Select Case infoEvento.EventType
                                Case SAPbouiCOM.BoEventTypes.et_CHOOSE_FROM_LIST

                                Case SAPbouiCOM.BoEventTypes.et_PICKER_CLICKED

                            End Select
                    End Select
                End If
            End If
            Return MyBase.SBOApp_ItemEvent(infoEvento)
        Catch exCOM As System.Runtime.InteropServices.COMException
            objGlobal.conexionSAP.Mostrar_Error(exCOM, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        Catch ex As Exception
            objGlobal.conexionSAP.Mostrar_Error(ex, EXO_Generales.EXO_SAP.EXO_TipoMensaje.Excepcion)
            Return False
        End Try
    End Function
    Private Function EventHandler_FORM_LOAD_After(ByRef objGlobal As EXO_Generales.EXO_General, ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim oItem As SAPbouiCOM.Item
        EventHandler_FORM_LOAD_After = False

        Try
            oForm = objGlobal.conexionSAP.SBOApp.Forms.Item(pVal.FormUID)
            oForm.Freeze(True)
            'Botón de DOT
            oItem = oForm.Items.Add("btn_GDOT", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            oItem.Left = oForm.Items.Item("2").Left + oForm.Items.Item("2").Width + 5
            oItem.Width = oForm.Items.Item("2").Width + oForm.Items.Item("2").Height
            oItem.Top = oForm.Items.Item("2").Top
            oItem.Height = oForm.Items.Item("2").Height
            oItem.Enabled = False
            Dim oBtnAct As SAPbouiCOM.Button
            oBtnAct = CType(oItem.Specific, Button)
            oBtnAct.Caption = "Generar DOT"
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Find, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Add, SAPbouiCOM.BoModeVisualBehavior.mvb_False)
            oItem.SetAutoManagedAttribute(SAPbouiCOM.BoAutoManagedAttr.ama_Editable, SAPbouiCOM.BoAutoFormMode.afm_Ok, SAPbouiCOM.BoModeVisualBehavior.mvb_True)
            EventHandler_FORM_LOAD_After = True


        Catch exCOM As System.Runtime.InteropServices.COMException
            oForm.Freeze(False)
            Throw exCOM
        Catch ex As Exception
            oForm.Freeze(False)
            Throw ex
        Finally
            oForm.Freeze(False)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oForm, Object))
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))

        End Try
    End Function
    Private Function EventHandler_ItemPressed_After(ByRef objGlobal As EXO_Generales.EXO_General, ByRef pVal As EXO_Generales.EXO_infoItemEvent) As Boolean
        Dim oForm As SAPbouiCOM.Form = Nothing
        Dim oRs As SAPbobsCOM.Recordset = CType(objGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
        Dim sSQL As String = ""
        Dim sDocEntry As String = 0 : Dim sDocNum As String = "" : Dim sObjType As String = ""

        Dim dtArt As System.Data.DataTable = Nothing
        Dim dr As DataRow
        EventHandler_ItemPressed_After = False

        Try
            oForm = objGlobal.conexionSAP.SBOApp.Forms.Item(pVal.FormUID)
            Dim sTable_Origen As String = CType(oForm.Items.Item("4").Specific, SAPbouiCOM.EditText).DataBind.TableName
            Dim sTable_Origen_Lin As String = CType(CType(oForm.Items.Item("38").Specific, SAPbouiCOM.Matrix).Columns.Item("1").Cells.Item(1).Specific, SAPbouiCOM.EditText).DataBind.TableName
            sDocEntry = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("DocEntry", 0).Trim
            sDocNum = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("DocNum", 0).Trim
            sObjType = oForm.DataSources.DBDataSources.Item(sTable_Origen).GetValue("ObjType", 0).Trim
            If pVal.ItemUID = "btn_GDOT" Then 'Botón UDO Generar DOT
                dtArt = New System.Data.DataTable
                dtArt.Columns.Add("Articulo", GetType(String))
                dtArt.Columns.Add("Descripcion", GetType(String))
                dtArt.Columns.Add("Cantidad", GetType(Double))
                'Recorremos las líneas agrupadas y guardamos en un datatable los artículos y las cantidades.
                sSQL = "SELECT ""ItemCode"",""Dscription"", Sum(""Quantity"") ""Cantidad"" "
                sSQL &= " FROM """ & sTable_Origen_Lin & """ WHERE DocEntry=" & sDocEntry
                sSQL &= " GROUP BY ""ItemCode"", ""Dscription"" ORDER BY ""ItemCode"" "
                oRs.DoQuery(sSQL)
                If oRs.RecordCount > 0 Then
                    For i = 0 To oRs.RecordCount - 1
                        dr = dtArt.NewRow()
                        dr("Articulo") = oRs.Fields.Item("ItemCode").Value.ToString
                        dr("Descripcion") = oRs.Fields.Item("Dscription").Value.ToString
                        dr("Cantidad") = CType(oRs.Fields.Item("Cantidad").Value.ToString, Double)
                        dtArt.Rows.Add(dr)
                        oRs.MoveNext()
                    Next

                    CargarUDOADOT(objGlobal, dtArt, sDocEntry, sDocNum, sObjType)

                Else
                    objGlobal.conexionSAP.SBOApp.StatusBar.SetText("(EXO) - Error inesperado. No encuentra las líneas del documento", SAPbouiCOM.BoMessageTime.bmt_Short, SAPbouiCOM.BoStatusBarMessageType.smt_Error)
                End If
            End If

            EventHandler_ItemPressed_After = True

        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.Form(oForm)
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#Region "Métodos auxiliares"
    Public Function CargarUDOADOT(ByRef objGlobal As EXO_Generales.EXO_General, ByRef dtArt As System.Data.DataTable, ByVal sDocEntry As String, ByVal sDocNum As String, ByVal sObjType As String) As Boolean
        Dim sSQL As String = ""
        Dim oRs As SAPbobsCOM.Recordset = Nothing

        CargarUDOADOT = False

        Try
            EXO_ADOT._sDocEntry = sDocEntry.Trim
            EXO_ADOT._sDocNum = sDocNum.Trim
            EXO_ADOT._sObjType = sObjType.Trim
            EXO_ADOT._dtArt = dtArt
            oRs = CType(objGlobal.conexionSAP.compañia.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset), SAPbobsCOM.Recordset)
            sSQL = " SELECT ""DocEntry"" FROM ""@EXO_ADOT"" WHERE ""U_EXO_DOCENTRY""=" & sDocEntry.Trim & " and ""U_EXO_DOCNUM""='" & sDocNum.Trim & "' and ""U_EXO_OTYPE""='" & sObjType.Trim & "' "
            oRs.DoQuery(sSQL)
            If oRs.RecordCount > 0 Then
                Dim sEntry As String = oRs.Fields.Item("DocEntry").Value.ToString.Trim
                objGlobal.conexionSAP.cargaFormUdoBD_Clave("EXO_ADOT", sEntry)
            Else
                objGlobal.conexionSAP.cargaFormUdoBD("EXO_ADOT")
            End If

            CargarUDOADOT = True
        Catch exCOM As System.Runtime.InteropServices.COMException
            Throw exCOM
        Catch ex As Exception
            Throw ex
        Finally
            EXO_CleanCOM.CLiberaCOM.liberaCOM(CType(oRs, Object))
        End Try
    End Function
#End Region
End Class
