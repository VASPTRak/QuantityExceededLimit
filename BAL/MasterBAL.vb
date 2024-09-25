Imports System.Data.SqlClient
Imports log4net
Imports log4net.Config
Public Class MasterBAL
    Private Shared ReadOnly log As ILog = LogManager.GetLogger(GetType(MasterBAL))

    Shared Sub New()
        XmlConfigurator.Configure()
    End Sub

    Public Function GetQuantityExceededLimitInfo() As DataTable
        Dim dal = New GeneralizedDAL()
        Try

            Dim ds As DataSet = New DataSet()

            ds = dal.ExecuteStoredProcedureGetDataSet("usp_tt_Transaction_GetQuantityExceededLimitInfo")

            Return ds.Tables(0)

        Catch ex As Exception

            log.Error("Error occurred in GetCompanyBrandByCustomerID Exception is :" + ex.Message)
            Return Nothing
        Finally

        End Try
    End Function

    Public Function GetTransactionById(TransactionId As Integer, IsDeleted As Boolean) As DataTable
        Dim dal = New GeneralizedDAL()
        Try

            Dim ds As DataSet = New DataSet()

            Dim Param As SqlParameter() = New SqlParameter(1) {}

            Param(0) = New SqlParameter("@TransactionId", SqlDbType.Int)
            Param(0).Direction = ParameterDirection.Input
            Param(0).Value = TransactionId

            Param(1) = New SqlParameter("@IsDeleted", SqlDbType.Bit)
            Param(1).Direction = ParameterDirection.Input
            Param(1).Value = IsDeleted

            ds = dal.ExecuteStoredProcedureGetDataSet("usp_tt_Transaction_GetTransactionById", Param)

            Return ds.Tables(0)

        Catch ex As Exception

            log.Error("Error occurred in GetTransactionById Exception is :" + ex.Message)
            Return Nothing
        Finally

        End Try
    End Function

    Public Function GetCompanyBrandByCustomerID(CustomerId As Integer) As DataTable
        Dim dal = New GeneralizedDAL()
        Try

            Dim ds As DataSet = New DataSet()

            Dim Param As SqlParameter() = New SqlParameter(0) {}

            Param(0) = New SqlParameter("@CustomerId", SqlDbType.Int)
            Param(0).Direction = ParameterDirection.Input
            Param(0).Value = CustomerId

            ds = dal.ExecuteStoredProcedureGetDataSet("usp_tt_CompanyBrand_GetCompanyBrandByCustomerID", Param)

            Return ds.Tables(0)

        Catch ex As Exception

            log.Error("Error occurred in GetCompanyBrandByCustomerID Exception is :" + ex.Message)
            Return Nothing
        Finally

        End Try
    End Function
End Class
