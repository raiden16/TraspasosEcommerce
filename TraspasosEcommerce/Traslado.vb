Public Class Traslado

    Private cSBOApplication As SAPbouiCOM.Application '//OBJETO DE APLICACION
    Private cSBOCompany As SAPbobsCOM.Company     '//OBJETO DE CONEXION
    Dim conexionSQL As Sap.Data.Hana.HanaConnection

    '//----- METODO DE CREACION DE LA CLASE
    Public Sub New()
        MyBase.New()
        cSBOApplication = oCatchingEvents.SBOApplication
        cSBOCompany = oCatchingEvents.SBOCompany
        conectar()
    End Sub


    Public Function conectar() As Boolean
        Dim stCadenaConexion As String
        Try

            conectar = False

            ''---- Cargamos datos de archivo de configuracion

            '---- objeto compañia
            conexionSQL = New Sap.Data.Hana.HanaConnection

            '---- armamos cadena de conexion
            stCadenaConexion = "DRIVER={B1CRHPROXY32};UID=" & My.Settings.UserSQL & ";PWD=" & My.Settings.PassSQL & ";SERVERNODE=" & My.Settings.Server

            '---- realizamos conexion
            conexionSQL = New Sap.Data.Hana.HanaConnection(stCadenaConexion)

            conexionSQL.Open()

        Catch ex As Exception
            cSBOApplication.MessageBox("Error al conectar con HANA . " & ex.Message)
        End Try
    End Function


    Public Function Separar(ByVal DocNum As String)

        Dim stQueryH, stQueryH1 As String
        Dim oRecSetH, oRecSetH1 As SAPbobsCOM.Recordset
        Dim oStockTransfer As SAPbobsCOM.StockTransfer
        Dim CantidadR, CantidadL, Quantity As Double
        Dim ItemCode, Lote, Boxes, Delivery, Package, BatchNumber As String
        Dim llError As Long
        Dim lsError As String

        oRecSetH = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)
        oRecSetH1 = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.BoRecordset)

        oStockTransfer = cSBOCompany.GetBusinessObject(SAPbobsCOM.BoObjectTypes.oStockTransfer)

        Try

            stQueryH = "Select T1.""ItemCode"",T1.""Quantity"",T2.""ManBtchNum"",T1.""Quantity""/T2.""SalPackUn"" as ""Package"", case when T0.""CardCode""='XAXX010101002' then 'AMAZON' when T0.""CardCode""='XAXX010101003' then 'MERCADOLIBRE' else 'ECOMMERCE' end as ""Delivery"" from ORDR T0 Inner Join RDR1 T1 on T1.""DocEntry""=T0.""DocEntry"" Inner Join OITM T2 on T2.""ItemCode""=T1.""ItemCode"" Inner Join OITB T3 on T3.""ItmsGrpCod""=T2.""ItmsGrpCod"" where T3.""ItmsGrpNam"" in ('TAPETES','FOLLAJE SINTETICO') and T0.""DocNum""='" & DocNum & "'"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oStockTransfer.DocDate = DateTime.Now
                oStockTransfer.FromWarehouse = "001B"
                oStockTransfer.ToWarehouse = "702"
                oStockTransfer.Comments = DocNum
                oStockTransfer.ElectronicProtocols.GenerationType = 1
                oStockTransfer.ElectronicProtocols.Add()

                CreateTemporalTable()

                InsertTemporalTable(DocNum, "001B")

                oRecSetH1.MoveFirst()

                For x = 0 To oRecSetH.RecordCount - 1

                    ItemCode = oRecSetH.Fields.Item("ItemCode").Value
                    Quantity = oRecSetH.Fields.Item("Quantity").Value
                    Lote = oRecSetH.Fields.Item("ManBtchNum").Value
                    Boxes = oRecSetH.Fields.Item("Package").Value
                    Delivery = oRecSetH.Fields.Item("Delivery").Value
                    Package = oRecSetH.Fields.Item("Package").Value


                    oStockTransfer.Lines.ItemCode = ItemCode
                    oStockTransfer.Lines.FromWarehouseCode = "001B"
                    oStockTransfer.Lines.WarehouseCode = "702"
                    oStockTransfer.Lines.Quantity = Quantity
                    oStockTransfer.Lines.UserFields.Fields.Item("U_CajasReq").Value = Boxes
                    oStockTransfer.Lines.UserFields.Fields.Item("U_DeliveryType").Value = Delivery
                    oStockTransfer.Lines.UserFields.Fields.Item("U_NumPaq").Value = Package


                    If Lote = "Y" Then

                        stQueryH1 = "Select * from """ & cSBOCompany.CompanyDB & """.ListaLotes where ""ITEMCODE""='" & ItemCode & "' and ""CANTIDADLOTE"">0 order by ""CREATEDATE"" Desc"
                        oRecSetH1.DoQuery(stQueryH1)

                        If oRecSetH1.RecordCount > 0 Then

                            oRecSetH1.MoveFirst()
                            CantidadR = Format(Quantity, "0.000")

                            For l = 0 To oRecSetH1.RecordCount - 1

                                CantidadL = Format(oRecSetH1.Fields.Item("CANTIDADLOTE").Value, "0.000")

                                If CantidadR > CantidadL Then

                                    CantidadR = Format(CantidadR - CantidadL, "0.000")

                                    BatchNumber = oRecSetH1.Fields.Item("BatchNum").Value

                                    oStockTransfer.Lines.BatchNumbers.BatchNumber = BatchNumber
                                    oStockTransfer.Lines.BatchNumbers.Quantity = CantidadL

                                    oStockTransfer.Lines.BatchNumbers.Add()

                                    UpdateTemporalTable(BatchNumber, ItemCode, CantidadL - CantidadL)

                                    l = 0

                                Else

                                    BatchNumber = oRecSetH1.Fields.Item("BatchNum").Value

                                    oStockTransfer.Lines.BatchNumbers.BatchNumber = BatchNumber
                                    oStockTransfer.Lines.BatchNumbers.Quantity = CantidadR

                                    oStockTransfer.Lines.BatchNumbers.Add()

                                    UpdateTemporalTable(BatchNumber, ItemCode, CantidadL - CantidadR)

                                    l = oRecSetH1.RecordCount - 1

                                End If

                                oRecSetH1.MoveNext()

                            Next

                        End If

                    End If

                    oStockTransfer.Lines.Add()
                    oRecSetH.MoveNext()

                Next

                If oStockTransfer.Add() <> 0 Then

                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)
                    cSBOApplication.MessageBox("Error al trasladar TAPETES o FOLLAJE")
                    DropTemporalTable()

                Else

                    cSBOApplication.MessageBox("Traslado de tapetes o follaje creado con éxito")
                    DropTemporalTable()

                End If

            End If

            stQueryH = "Select T1.""ItemCode"",T1.""Quantity"",T2.""ManBtchNum"",T1.""Quantity""/T2.""SalPackUn"" as ""Package"", case when T0.""CardCode""='XAXX010101002' then 'AMAZON' when T0.""CardCode""='XAXX010101003' then 'MERCADOLIBRE' else 'ECOMMERCE' end as ""Delivery"" from ORDR T0 Inner Join RDR1 T1 on T1.""DocEntry""=T0.""DocEntry"" Inner Join OITM T2 on T2.""ItemCode""=T1.""ItemCode"" Inner Join OITB T3 on T3.""ItmsGrpCod""=T2.""ItmsGrpCod"" where T3.""ItmsGrpNam"" in ('PASTO ARTIFICIAL') and T0.""DocNum""='" & DocNum & "'"
            oRecSetH.DoQuery(stQueryH)

            If oRecSetH.RecordCount > 0 Then

                oStockTransfer.DocDate = DateTime.Now
                oStockTransfer.FromWarehouse = "001"
                oStockTransfer.ToWarehouse = "702"
                oStockTransfer.Comments = DocNum
                oStockTransfer.ElectronicProtocols.GenerationType = 1
                oStockTransfer.ElectronicProtocols.Add()

                CreateTemporalTable()

                InsertTemporalTable(DocNum, "001")

                oRecSetH1.MoveFirst()

                For x = 0 To oRecSetH.RecordCount - 1

                    ItemCode = oRecSetH.Fields.Item("ItemCode").Value
                    Quantity = oRecSetH.Fields.Item("Quantity").Value
                    Lote = oRecSetH.Fields.Item("ManBtchNum").Value
                    Boxes = oRecSetH.Fields.Item("Package").Value
                    Delivery = oRecSetH.Fields.Item("Delivery").Value
                    Package = oRecSetH.Fields.Item("Package").Value


                    oStockTransfer.Lines.ItemCode = ItemCode
                    oStockTransfer.Lines.FromWarehouseCode = "001"
                    oStockTransfer.Lines.WarehouseCode = "702"
                    oStockTransfer.Lines.Quantity = Quantity
                    oStockTransfer.Lines.UserFields.Fields.Item("U_CajasReq").Value = Boxes
                    oStockTransfer.Lines.UserFields.Fields.Item("U_DeliveryType").Value = Delivery
                    oStockTransfer.Lines.UserFields.Fields.Item("U_NumPaq").Value = Package


                    If Lote = "Y" Then

                        stQueryH1 = "Select * from """ & cSBOCompany.CompanyDB & """.ListaLotes where ""ITEMCODE""='" & ItemCode & "' and ""CANTIDADLOTE"">0 order by ""CREATEDATE"" Desc"
                        oRecSetH1.DoQuery(stQueryH1)

                        If oRecSetH1.RecordCount > 0 Then

                            oRecSetH1.MoveFirst()
                            CantidadR = Format(Quantity, "0.000")

                            For l = 0 To oRecSetH1.RecordCount - 1

                                CantidadL = Format(oRecSetH1.Fields.Item("CANTIDADLOTE").Value, "0.000")

                                If CantidadR > CantidadL Then

                                    CantidadR = Format(CantidadR - CantidadL, "0.000")

                                    BatchNumber = oRecSetH1.Fields.Item("BatchNum").Value

                                    oStockTransfer.Lines.BatchNumbers.BatchNumber = BatchNumber
                                    oStockTransfer.Lines.BatchNumbers.Quantity = CantidadL

                                    oStockTransfer.Lines.BatchNumbers.Add()

                                    UpdateTemporalTable(BatchNumber, ItemCode, CantidadL - CantidadL)

                                    l = 0

                                Else

                                    BatchNumber = oRecSetH1.Fields.Item("BatchNum").Value

                                    oStockTransfer.Lines.BatchNumbers.BatchNumber = BatchNumber
                                    oStockTransfer.Lines.BatchNumbers.Quantity = CantidadR

                                    oStockTransfer.Lines.BatchNumbers.Add()

                                    UpdateTemporalTable(BatchNumber, ItemCode, CantidadL - CantidadR)

                                    l = oRecSetH1.RecordCount - 1

                                End If

                                oRecSetH1.MoveNext()

                            Next

                        End If

                    End If

                    oStockTransfer.Lines.Add()
                    oRecSetH.MoveNext()

                Next

                If oStockTransfer.Add() <> 0 Then

                    cSBOCompany.GetLastError(llError, lsError)
                    Err.Raise(-1, 1, lsError)
                    cSBOApplication.MessageBox("Error al trasladar PASTO")
                    DropTemporalTable()

                Else

                    cSBOApplication.MessageBox("Traslado de pasto creado con éxito")
                    DropTemporalTable()

                End If

            End If

            conexionSQL.Close()

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al crear el traslado. " & ex.Message)
            DropTemporalTable()
            conexionSQL.Close()

        End Try

    End Function

    Public Function CreateTemporalTable()

        Dim stQueryH1 As String
        Dim comm As New Sap.Data.Hana.HanaCommand
        Dim DA As New Sap.Data.Hana.HanaDataAdapter
        Dim ds As New DataSet

        Try

            stQueryH1 = "CREATE COLUMN TABLE """ & cSBOCompany.CompanyDB & """.ListaLotes (BatchNum NVARCHAR(50), ItemCode NVARCHAR(50), WhsCode NVARCHAR(5), CantidadLote Double, CreateDate NVARCHAR(20));"
            comm.CommandText = stQueryH1
            comm.Connection = conexionSQL
            DA.SelectCommand = comm
            DA.Fill(ds)

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al CreateTemporalTable. " & ex.Message)
            conexionSQL.Close()

        End Try

    End Function


    Public Function InsertTemporalTable(ByVal DocNum As String, ByVal FromWhsCode As String)

        Dim stQueryH1, stQueryH2 As String
        Dim comm, comm2 As New Sap.Data.Hana.HanaCommand
        Dim DA, DA2 As New Sap.Data.Hana.HanaDataAdapter
        Dim ds, ds2 As New DataSet
        Dim tabla As DataTable
        Dim BatchNum, ItemCode, WhsCode As String
        Dim CantidadLote As Double
        Dim CreateDate As String

        Try

            stQueryH1 = "Call """ & cSBOCompany.CompanyDB & """.LotesORDR(" & DocNum & ",'" & FromWhsCode & "')"
            comm.CommandText = stQueryH1
            comm.Connection = conexionSQL
            DA.SelectCommand = comm
            DA.Fill(ds)

            If ds.Tables(0).Rows.Count > 0 Then

                tabla = ds.Tables(0)

                For i = 0 To ds.Tables(0).Rows.Count - 1

                    BatchNum = ds.Tables(0).Rows(i).Item("BatchNum")
                    ItemCode = ds.Tables(0).Rows(i).Item("ItemCode")
                    WhsCode = ds.Tables(0).Rows(i).Item("WhsCode")
                    CantidadLote = ds.Tables(0).Rows(i).Item("CantidadLote")
                    CreateDate = ds.Tables(0).Rows(i).Item("CreateDate")

                    stQueryH2 = "Insert Into """ & cSBOCompany.CompanyDB & """.ListaLotes values ('" & BatchNum & "','" & ItemCode & "','" & WhsCode & "'," & CantidadLote & ",'" & CreateDate & "')"
                    comm2.CommandText = stQueryH2
                    comm2.Connection = conexionSQL
                    DA2.SelectCommand = comm2
                    DA2.Fill(ds2)

                Next

            End If

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al InsertTemporalTable. " & ex.Message)
            conexionSQL.Close()

        End Try

    End Function


    Public Function UpdateTemporalTable(ByVal Lote As String, ByVal ItemCode As String, ByVal Cantidad As Double)

        Dim stQueryH1 As String
        Dim comm As New Sap.Data.Hana.HanaCommand
        Dim DA As New Sap.Data.Hana.HanaDataAdapter
        Dim ds As New DataSet

        Try

            stQueryH1 = "Update """ & cSBOCompany.CompanyDB & """.ListaLotes set ""CANTIDADLOTE""=" & Cantidad & " where ""BATCHNUM""='" & Lote & "' and ""ITEMCODE""='" & ItemCode & "'"
            comm.CommandText = stQueryH1
            comm.Connection = conexionSQL
            DA.SelectCommand = comm
            DA.Fill(ds)

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al UpdateTemporalTable. " & ex.Message)
            conexionSQL.Close()

        End Try

    End Function


    Public Function DropTemporalTable()

        Dim stQueryH1 As String
        Dim comm As New Sap.Data.Hana.HanaCommand
        Dim DA As New Sap.Data.Hana.HanaDataAdapter
        Dim ds As New DataSet

        Try

            stQueryH1 = "Drop table """ & cSBOCompany.CompanyDB & """.ListaLotes"
            comm.CommandText = stQueryH1
            comm.Connection = conexionSQL
            DA.SelectCommand = comm
            DA.Fill(ds)

        Catch ex As Exception

            cSBOApplication.MessageBox("Error al DropTemporalTable. " & ex.Message)
            conexionSQL.Close()

        End Try

    End Function

End Class
