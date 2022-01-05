Public Class ORDR

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Private coForm As SAPbouiCOM.Form           '//FORMA
    Private Directorio As String

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        Directorio = oCatchingEvents.csDirectory
    End Sub

    '//----- AGREGA ELEMENTOS A LA FORMA
    Public Sub addFormItems(ByVal FormUID As String)
        Dim loItem As SAPbouiCOM.Item
        Dim loButton As SAPbouiCOM.Button
        Dim lsItemRef As String

        Try
            '//AGREGA BOTON MOVIMIENTOS EN PEDIDOS DE COMPRAS
            coForm = cSBOApplication.Forms.Item(FormUID)
            lsItemRef = "2"
            loItem = coForm.Items.Add("btOwtr", SAPbouiCOM.BoFormItemTypes.it_BUTTON)
            loItem.Left = coForm.Items.Item(lsItemRef).Left + coForm.Items.Item(lsItemRef).Width + coForm.Items.Item(lsItemRef).Width + 40 + 10
            loItem.Top = coForm.Items.Item(lsItemRef).Top
            loItem.Width = coForm.Items.Item(lsItemRef).Width + 40
            loItem.Height = coForm.Items.Item(lsItemRef).Height
            loButton = loItem.Specific
            loButton.Caption = "Trasladar"

        Catch ex As Exception
            cSBOApplication.MessageBox("DocumentoSBO. agregar elementos a la forma. " & ex.Message)
        Finally
            coForm = Nothing
            loItem = Nothing
            loButton = Nothing
        End Try
    End Sub

End Class
