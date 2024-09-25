Imports System.Configuration
Imports System.IO
Imports System.Net
Imports System.Net.Mail
Imports System.Transactions
Imports log4net
Imports log4net.Config
Imports MathNet.Numerics.Distributions
Imports NPOI.HSSF.Record
Imports NPOI.HSSF.UserModel
Imports NPOI.SS.Formula.Functions
Imports NPOI.SS.UserModel
Imports NPOI.SS.Util
Imports NPOI.XSSF.UserModel
Imports SixLabors.Fonts

Module Module1
    Private ReadOnly log As ILog = LogManager.GetLogger(GetType(Module1))
    Sub Main()
        XmlConfigurator.Configure()
        log.Info("Execution started")

        Try
            'Dim body As String = String.Empty
            'Using sr As New StreamReader(ConfigurationManager.AppSettings("PathForQuantityExceededLimitEmailTemplate"))
            '    body = sr.ReadToEnd()
            'End Using
            'Dim filePath As String = ConfigurationManager.AppSettings("PathForSaveQuantityExceededLimitFile").ToString()
            'log.Debug("Template Path: " + ConfigurationManager.AppSettings("PathForQuantityExceededLimitEmailTemplate") + " File folder: " + ConfigurationManager.AppSettings("PathForSaveQuantityExceededLimitFile").ToString())
            StartProcessing()
        Catch ex As Exception
            log.Error("Error occurred in Main. ex is :" & ex.Message)
        End Try
        log.Info("Execution stopped")
    End Sub

    Private Sub StartProcessing()
        Try
            log.Info("In StartProcessing")

            Dim OBJMasterBAL = New MasterBAL()
            OBJMasterBAL = New MasterBAL()

            Dim dtQuantityExceededLimit As DataTable = New DataTable()
            Dim dtSingleTransacton As DataTable = New DataTable()

            dtQuantityExceededLimit = OBJMasterBAL.GetQuantityExceededLimitInfo()

            If dtQuantityExceededLimit IsNot Nothing Then
                If dtQuantityExceededLimit.Rows.Count > 0 Then

                    Dim dtQuantityExceededLimitInfo As DataTable = New DataTable("QuantityExceededLimitInfo")

                    dtQuantityExceededLimitInfo.Columns.Add("CompanyName", GetType(System.String))
                    dtQuantityExceededLimitInfo.Columns.Add("VehicleNumber", GetType(System.String))
                    dtQuantityExceededLimitInfo.Columns.Add("TransactionNumber", GetType(System.String))
                    dtQuantityExceededLimitInfo.Columns.Add("DateValue", GetType(System.String))
                    dtQuantityExceededLimitInfo.Columns.Add("TimeValue", GetType(System.String))
                    dtQuantityExceededLimitInfo.Columns.Add("FuelQuantityValue", GetType(System.Decimal))

                    Dim TransactionId As Integer = 0
                    For i = 0 To dtQuantityExceededLimit.Rows.Count - 1
                        Try
                            TransactionId = dtQuantityExceededLimit.Rows(i)("TransactionId")
                            Dim FuelQuantity As Integer = dtQuantityExceededLimit.Rows(i)("FuelQuantity")
                            Dim TransactionQtyLimit As Integer = dtQuantityExceededLimit.Rows(i)("TransactionQtyLimit")

                            dtSingleTransacton = OBJMasterBAL.GetTransactionById(TransactionId, False)

                            Dim TransactionCustomerName As String = ""
                            Dim TransactionCustomerId As Integer = 0
                            Dim transactionDate As String = ""
                            Dim transactionTime As String = ""
                            Dim transactionVehicleNumber As String = ""
                            Dim transactionSiteID As Integer = 0
                            Dim TransactionIsOtherRequire As Boolean = False
                            Dim TransactionNumber As String = ""
                            Dim TransactionOtherLabel As String = ""
                            Dim TransactionFuelQuantity As String = ""

                            If dtSingleTransacton IsNot Nothing Then
                                If dtSingleTransacton.Rows.Count > 0 Then
                                    TransactionCustomerName = dtSingleTransacton.Rows(0)("Company").ToString()
                                    TransactionCustomerId = dtSingleTransacton.Rows(0)("CustomerId").ToString()

                                    Try
                                        transactionDate = dtSingleTransacton.Rows(0)("Date")
                                    Catch ex As Exception
                                        log.Error("Error occurred in sendTransactionEmail transactionDate. Transaction # " & TransactionId & ". ex is :" & ex.Message)
                                    End Try

                                    Try
                                        transactionTime = dtSingleTransacton.Rows(0)("Time")
                                    Catch ex As Exception
                                        log.Error("Error occurred in sendTransactionEmail transactionTime. Transaction # " & TransactionId & ". ex is :" & ex.Message)
                                    End Try

                                    Try
                                        transactionSiteID = dtSingleTransacton.Rows(0)("SiteID").ToString()
                                    Catch ex As Exception
                                        log.Error("Error occurred in sendTransactionEmail transactionSiteID. Transaction # " & TransactionId & ". ex is :" & ex.Message)
                                    End Try

                                    Try
                                        transactionVehicleNumber = dtSingleTransacton.Rows(0)("VehicleNumber").ToString().Trim()
                                    Catch ex As Exception
                                        log.Error("Error occurred in sendTransactionEmail TransactionVehicleNumber. Transaction # " & TransactionId & ". ex is :" & ex.Message)
                                    End Try

                                    Try
                                        TransactionIsOtherRequire = dtSingleTransacton.Rows(0)("IsOtherRequire")
                                    Catch ex As Exception
                                        log.Error("Error occurred in sendTransactionEmail TransactionIsOtherRequire. Transaction # " & TransactionId & ". ex is :" & ex.Message)
                                    End Try

                                    Try
                                        TransactionNumber = dtSingleTransacton.Rows(0)("TransactionNumber").ToString()
                                    Catch ex As Exception
                                        log.Error("Error occurred in sendTransactionEmail TransactionNumber. Transaction # " & TransactionId & ". ex is :" & ex.Message)
                                    End Try

                                    Try
                                        TransactionOtherLabel = dtSingleTransacton.Rows(0)("OtherLabel").ToString()
                                    Catch ex As Exception
                                        log.Error("Error occurred in sendTransactionEmail TransactionOtherLabel. Transaction # " & TransactionId & ". ex is :" & ex.Message)
                                    End Try

                                    Try
                                        TransactionFuelQuantity = dtSingleTransacton.Rows(0)("FuelQuantity").ToString()
                                    Catch ex As Exception
                                        log.Error("Error occurred in sendTransactionEmail TransactionFuelQuantity. Transaction # " & TransactionId & ". ex is :" & ex.Message)
                                    End Try
                                End If
                            End If

                            dtQuantityExceededLimitInfo.Rows.Add(TransactionCustomerName, transactionVehicleNumber, TransactionNumber, transactionDate, transactionTime, TransactionFuelQuantity)
                        Catch ex As Exception
                            log.Error("Error occurred in sending transaction quantity exceeded limit email for TransactionId : " & TransactionId & ". ex is :" & ex.Message)
                        End Try
                    Next

                    If dtQuantityExceededLimitInfo IsNot Nothing Then
                        If dtQuantityExceededLimitInfo.Rows.Count > 0 Then
                            ExcelProcessing(dtQuantityExceededLimitInfo)
                        Else
                            log.Error("No data Generated..")
                        End If
                    End If

                End If
            End If

        Catch ex As Exception
            log.Error("Error occurred StartProcessing. ex is :" & ex.Message)
        End Try
    End Sub

    Private Sub ExcelProcessing(dtData As DataTable)
        Try
            log.Info("In ExcelProcessing")

            '========== Main Support Data ======================
            Dim dtTransactions As DataTable = dtData
            Dim dtFinal As DataTable = New DataTable()

            Dim columns As String = "Company Name,Vehicle Number,Transaction Number,Transaction Date,Transaction Time,Transaction Fuel Quantity"
            Dim columnsOfArray() As String = columns.Split(",")

            For l As Integer = 0 To columnsOfArray.Count - 1
                dtFinal.Columns.Add(columnsOfArray(l).ToString(), System.Type.[GetType]("System.String"))
            Next

            Dim drColumnNameNew As DataRow = dtFinal.NewRow()

            Dim columnNameList As List(Of String) = New List(Of String)
            For Each column As DataColumn In dtFinal.Columns
                columnNameList.Add(column.ColumnName)
            Next

            drColumnNameNew.ItemArray = columnNameList.ToArray()
            dtFinal.Rows.Add(drColumnNameNew)

            For Each dr As DataRow In dtTransactions.Rows
                Try

                    Dim drNew As DataRow = dtFinal.NewRow()
                    drNew("Company Name") = dr("CompanyName").ToString()
                    drNew("Vehicle Number") = dr("VehicleNumber").ToString()
                    drNew("Transaction Number") = dr("TransactionNumber").ToString()
                    drNew("Transaction Date") = dr("DateValue").ToString()
                    drNew("Transaction Time") = dr("TimeValue").ToString()
                    drNew("Transaction Fuel Quantity") = dr("FuelQuantityValue").ToString()

                    dtFinal.Rows.Add(drNew)
                Catch ex As Exception
                    log.Error("Error occurred in ExcelProcessing for loop. ex is :" & ex.Message)
                End Try
            Next
            '==================================================================

            Dim FileName = WriteExcelWithNPOI("xlsx", dtFinal)

            SendEmail(FileName) '

        Catch ex As Exception
            log.Error("Error occurred in ExcelProcessing. ex is :" & ex.Message)
        End Try
    End Sub

    Public Function WriteExcelWithNPOI(ByVal extension As String, ByVal dtData As DataTable) As String
        Dim fullPath = ""
        Try

            Dim workbook As IWorkbook

            If extension = "xlsx" Then
                workbook = New XSSFWorkbook()
            ElseIf extension = "xls" Then
                workbook = New HSSFWorkbook()
            Else
                Throw New Exception("This format is not supported")
            End If

            Dim rowCounter As Integer = 0
            Dim sheet1 As ISheet = workbook.CreateSheet("Sheet 1")

            Dim boldFont As XSSFFont = workbook.CreateFont()
            boldFont.IsBold = True

            ' create bordered cell style
            Dim borderedHeaderCellStyle As XSSFCellStyle = workbook.CreateCellStyle()
            borderedHeaderCellStyle.SetFont(boldFont)
            borderedHeaderCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Medium
            borderedHeaderCellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Medium
            borderedHeaderCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Medium
            borderedHeaderCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Medium

            Dim borderedCellStyle As XSSFCellStyle = workbook.CreateCellStyle()
            borderedCellStyle.BorderLeft = NPOI.SS.UserModel.BorderStyle.Thin
            borderedCellStyle.BorderTop = NPOI.SS.UserModel.BorderStyle.Thin
            borderedCellStyle.BorderRight = NPOI.SS.UserModel.BorderStyle.Thin
            borderedCellStyle.BorderBottom = NPOI.SS.UserModel.BorderStyle.Thin

            '========== Main Transaction Data ======================
            For i As Integer = 0 To dtData.Rows.Count - 1
                rowCounter = rowCounter + 1
                Dim row As IRow = sheet1.CreateRow(i + 1) ' One row already added before so to balance added 1+
                For j As Integer = 0 To dtData.Columns.Count - 1
                    Dim cell As ICell = row.CreateCell(j)
                    Dim columnName As String = dtData.Columns(j).ToString()
                    cell.SetCellValue(dtData.Rows(i)(columnName).ToString())
                    If i = 0 Then
                        cell.CellStyle = borderedHeaderCellStyle
                    Else
                        cell.CellStyle = borderedCellStyle
                    End If
                Next
            Next

            '==================================================================

            Dim filePath As String = ConfigurationManager.AppSettings("PathForSaveQuantityExceededLimitFile").ToString()
            fullPath = filePath & DateTime.Now.ToString("yyyy_MM_dd_hh_mm_ss_fff") & ".xls"


            Using exportData = New MemoryStream()
                workbook.Write(exportData)
                Try
                    If (File.Exists(fullPath)) Then
                        'File.Delete(fullPath)
                        System.IO.File.Delete(fullPath)
                    End If
                Catch ex As Exception
                    log.Info("In WriteExcelWithNPOI => exception in delete file. filename: " & fullPath & "; exception is: " & ex.ToString())
                End Try


                log.Info("In WriteExcelWithNPOI => step 9")
                Dim bw As BinaryWriter = New BinaryWriter(File.Open(fullPath, FileMode.OpenOrCreate))

                bw.Write(exportData.ToArray())

                log.Info("In WriteExcelWithNPOI => step 10")
                bw.Close()
                workbook.Close()

            End Using

        Catch ex As Exception
            If Not (ex.Message.Contains("Thread was being aborted") = True) Then
                log.Error("Error occurred in WriteExcelWithNPOI. Exception is :" + ex.Message)
            End If
        End Try
        Return fullPath
    End Function

    Private Sub SendEmail(FileName As String)
        Try

            Dim body As String = String.Empty
            Using sr As New StreamReader(ConfigurationManager.AppSettings("PathForQuantityExceededLimitEmailTemplate"))
                body = sr.ReadToEnd()
            End Using
            '------------------

            Try
                body = body.Replace("ImageSign", "<img src=""https://www.fluidsecure.net/Content/Images/FluidSECURELogo.png"" style=""width:200px""/>")
                body = body.Replace("SupportTeamName", "FluidSecure Support Team")
                body = body.Replace("supportemail", "support@fluidsecure.com")
                body = body.Replace("SupportPhoneNumber", "1-850-878-4585")
                body = body.Replace("SupportLine1", "Press ""0"" During Normal Business Hours:  Monday - Friday 8:00am - 5:00pm (EST)")
                body = body.Replace("SupportLine2", "Press ""7"" After Normal Business Hours")
                body = body.Replace("websiteURLHREF", "https://www.fluidsecure.com")
                body = body.Replace("webisteURL", "www.fluidsecure.com")
            Catch ex As Exception
                body = body.Replace("ImageSign", "")
            End Try

            Dim mailClient As New SmtpClient(ConfigurationManager.AppSettings("smtpServer"))
            mailClient.UseDefaultCredentials = False
            mailClient.Credentials = New NetworkCredential(ConfigurationManager.AppSettings("emailAccount"), ConfigurationManager.AppSettings("emailPassword"))
            mailClient.Port = Convert.ToInt32(ConfigurationManager.AppSettings("smtpPort"))

            Dim messageSend As New MailMessage()
            messageSend.Body = body
            messageSend.IsBodyHtml = True
            messageSend.Subject = ConfigurationManager.AppSettings("TransactionQuantityExceededLimitSubject")
            messageSend.From = New MailAddress(ConfigurationManager.AppSettings("FromEmail"))

            If FileExists(FileName) Then
                Dim attach As Attachment = New Attachment(FileName)
                attach.Name = ConfigurationManager.AppSettings("FileName").ToString()
                messageSend.Attachments.Add(attach)
            End If

            mailClient.EnableSsl = Convert.ToBoolean(ConfigurationManager.AppSettings("EnableSsl"))

            Dim SupportEmail As String() = ConfigurationManager.AppSettings("TransactionQuantityExceededLimitEmail").ToString().Split(";")

            Dim EmailsList As String = ""

            For index = 0 To SupportEmail.Length - 1
                If (SupportEmail(index).Trim() = "") Then
                    Continue For
                End If
                EmailsList = EmailsList & SupportEmail(index).Trim() & ";"
                messageSend.To.Add(New MailAddress(SupportEmail(index).Trim(), ""))
            Next

            Try
                mailClient.Send(messageSend)
                log.Info("At SendEmailForQuantityExceededLimit. Email send To : " & EmailsList & ".")
            Catch ex As Exception
                log.Error("Error occurred in sending transaction quantity exceeded limit email to EmailId : " & EmailsList & ". ex is :" & ex.Message)
            End Try


            Try
                System.IO.File.Delete(FileName)
            Catch ex As Exception
                log.Error("When deleting file after email send : " + ex.ToString())
            End Try


        Catch ex As Exception
            log.Debug("Exception occurred in SendEmail. ex is :" & ex.ToString())
        End Try
    End Sub

    Private Function FileExists(ByVal FileFullPath As String) As Boolean
        Try
            If FileFullPath = "" Then Return False

            Dim f As New IO.FileInfo(FileFullPath)
            Return f.Exists

        Catch ex As Exception
            log.Error("Exception occurred in FileExists. ex is :" & ex.ToString())
            Return False
        End Try
    End Function
End Module
