Imports System.Data.SqlClient
Imports System.Windows.Forms.VisualStyles.VisualStyleElement.StartPanel

Public Class CShopTransfer
    Inherits CUtility
    Public Function SaveShopTransHead(Reference As String, TransferDate As Date, ShopRef As String, ShopName As String, ToShopRef As String, ToShopName As String, TotalQtyOut As Integer, TotalQtyIn As Integer) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblShopTransfers (Reference,TransferDate,ShopRef,ShopName,ToShopRef,ToShopName,TotalQtyOut,TotalQtyIn,CreatedBy,CreatedDate) VALUES (@Reference,@TransferDate,@ShopRef,@ShopName,@ToShopRef,@ToShopName,@TotalQtyOut,@TotalQtyIn,@CreatedBy,@CreatedDate)"
                With .Parameters
                    .AddWithValue("@Reference", Reference)
                    .AddWithValue("@TransferDate", CDate(TransferDate))
                    .AddWithValue("@ShopRef", ShopRef)
                    .AddWithValue("@ShopName", ShopName)
                    .AddWithValue("@ToShopRef", ToShopRef)
                    .AddWithValue("@ToShopName", ToShopName)
                    .AddWithValue("@TotalQtyOut", CInt(TotalQtyOut))
                    .AddWithValue("@TotalQtyIn", CInt(TotalQtyIn))
                    .AddWithValue("@CreatedBy", username)
                    .AddWithValue("@CreatedDate", Date.Now)
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
    Public Function SaveShopTransLines(TransferID As Integer, smtoid As Integer, smtiid As Integer, StockCode As String, CurrQty As Integer, TOQty As Integer, TIQty As Integer) As Boolean
        Using conn As New SqlConnection(GetConnString(1))
            Dim inscmd As New SqlCommand
            With inscmd
                .Connection = conn
                .Connection.Open()
                .CommandType = CommandType.Text
                .CommandText = "INSERT INTO tblShopTransferLines (TransferID,SMTOID,SMTIID,StockCode,CurrentQty,TOQty,TIQty) VALUES (@TransferID,@SMTOID,@SMTIID,@StockCode,@CurrentQty,@TOQty,@TIQty)"
                With .Parameters
                    .AddWithValue("@TransferID", TransferID)
                    .AddWithValue("@SMTOID", CInt(smtoid))
                    .AddWithValue("@SMTIID", CInt(smtiid))
                    .AddWithValue("@StockCode", StockCode)
                    .AddWithValue("@CurrentQty", CInt(CurrQty))
                    .AddWithValue("@TOQty", CInt(TOQty))
                    .AddWithValue("@TIQty", CInt(TIQty))
                End With
                .ExecuteNonQuery()
            End With
        End Using
        Return True
    End Function
    Public Function SaveShopTransferHead() As Integer

        Return Result
    End Function
    Public Function SaveShopTransferLine() As Boolean

        Return True
    End Function
    Public Function UpdateShopTransferHead() As Boolean

        Return True
    End Function
    Public Function UpdateShopTransferLine() As Boolean

        Return True
    End Function
    Public Function DeleteShopTransfer() As Boolean

        Return True
    End Function
End Class
