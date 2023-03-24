﻿Imports System.Text
Imports System.Reflection
Imports System.Net.Mail
Imports System
Imports System.Globalization
Imports System.IO

Module Module1

    Dim ClsGnrl As New ClsGeneral
    Dim ClsAutRep As New ClsAutoReport
    Dim start_date, end_date, current_date As DateTime
    Dim current_HHmm As Integer 'jam dan menit yang sedang berjalan
    Dim status_sudah_email As Boolean 'jika sudah email maka nilai true, jika belum nilai false

    Dim tableManpower As New DataTable 'sheet1 dan content email
    Dim tableType As DataTable = Nothing 'sheet2


    Sub Main()

        Console.WriteLine("APPLICATION CALCULATION AND REPORT OPMJ")
        Console.WriteLine("")


        Console.WriteLine("##START SET CONFIG")
        ClsConfig.get_variable_setting()
        Console.WriteLine("##FINISH SET CONFIG")
        Console.WriteLine("")


        Console.WriteLine("---> Setup date...")
        Console.WriteLine("")
        ''contoh irisan bulan berjalan dengan bulan sebelumnya
        'current_date = Convert.ToDateTime("2/18/2023 07:45:00")
        current_date = Now

        'TRIAL DATA
        'end_date = Convert.ToDateTime("12/06/2022 07:00:00")
        'start_date = Convert.ToDateTime("12/06/2022 07:00:00")
        'current_date = Convert.ToDateTime("12/06/2022 07:00:00")

        'Jika < jam 08.29 WIB masih menggunakan periode tanggal kemarin (menggunakan jam 08.29 karena proses kalkulasi pasti ada delay dan pasti > 07.30 WIB)
        'menggunakan 08.29 WIB karena kalkulasi pertama per-hari dijam 08.30
        'fungsi ini untuk mengikuti periode hari produksi DTI 07.30 s/d 07.30 dan juga handle periode pada irisan bulan
        '(akhir bulan dengan awal bulan contoh : tgl. 01-11-22 jam 07.00 dihitung masih bulan sebelumnya)
        If Int32.Parse(current_date.ToString("HHmm")) <= 829 Then
            current_date = current_date.AddDays(-1)
        End If

        start_date = DateSerial(Year(current_date), Month(current_date), 1)
        end_date = ClsGeneral.get_last_date(current_date) 'belum digunakan, lebih baik dengan current_date (periode data yg dibutuhkan saja)
        current_HHmm = Int32.Parse(current_date.ToString("HHmm"))


        Console.WriteLine("---> Check the calculation status and send an email...")
        Console.WriteLine("")
        'JIKA SUDAH KIRIM EMAIL DIJAM YANG SAMA, TIDAK KIRIM EMAIL LAGI
        status_sudah_email = ClsGnrl.cek_status_sudah_email(current_date, current_HHmm)

        If status_sudah_email = True Then
            'JIKA PROGRAM SUDAH KALKULASI DAN KIRIM EMAIL, DI-SKIP PROSES
            Console.WriteLine("PROGRAM SUDAH KALKULASI DAN KIRIM EMAIL")
        Else
            get_calculation(start_date, current_date, current_HHmm, end_date)
        End If

        'kirim laporan monitoring mail sender OPMJ hanya jam 07.30
        If Int32.Parse(current_date.ToString("HHmm")) >= 730 And Int32.Parse(current_date.ToString("HHmm")) <= 829 Then
            status_sudah_email = ClsGnrl.cek_status_sudah_email(current_date, 9999)
            If status_sudah_email = True Then
                'JIKA PROGRAM SUDAH KIRIM EMAIL MONITORING, DI-SKIP PROSES
                Console.WriteLine("PROGRAM SUDAH KALKULASI DAN KIRIM EMAIL MONITORING")
            Else
                'JIKA PROGRAM BELUM KIRIM EMAIL MONITORING, LAKUKAN KIRIM EMAIL MONITORING
                ClsAutRep.AutoReportMonitoring(start_date, current_date)
            End If
        End If

    End Sub

    Private Sub get_calculation(ByVal startDate As Date,
                                ByVal currentDate As Date,
                                ByVal currentHHmm As Integer,
                                ByVal endDate As Date)

        Console.WriteLine("##START DATA CALCULATION")

        ClsGnrl.StartCalc(startDate, endDate, tableManpower, tableType)

        Console.WriteLine("##FINISH DATA CALCULATION")

        export_excel_and_send_mail(startDate, currentDate, currentHHmm, tableManpower, tableType)

    End Sub

    Private Sub export_excel_and_send_mail(ByVal startDate As Date, ByVal currentDate As Date, ByVal currentHHmm As Integer,
                                           ByVal tableManpower As DataTable,
                                           ByVal tableType As DataTable)

        Console.WriteLine("")
        Console.WriteLine("##START EXPORT DATA TO EXCEL")
        Console.WriteLine("")

        Try

            Dim nama_file_template_n_path As String
            Dim nama_file_simpan As String
            Dim lokasi_simpan_file As String
            Dim mat_type As String = ""
            Dim OpenReport As Boolean = False 'Open file excel
            Dim ExcelOutputFile As String = ""

            'myCulture = New System.Globalization.CultureInfo("en-US", True)
            nama_file_template_n_path = System.AppDomain.CurrentDomain.BaseDirectory & ClsConfig.nama_file_template & ".xlsx"
            'dir = System.Environment.GetFolderPath(Environment.SpecialFolder.MyDocuments)
            lokasi_simpan_file = ClsConfig.lokasi_simpan_file

            Dim xlApp As Object = CreateObject("Excel.Application")
            Dim xlWorkBook As Object = xlApp.Workbooks.Open(nama_file_template_n_path)

            Dim i As Integer
            Dim starting_row As Integer
            Dim row_count As Integer
            Dim last_row As Integer

            '(SHEET1) OPMJ BY MANPOWER
            Dim xlWorkSheet1 As Object
            xlWorkSheet1 = xlWorkBook.WorkSheets(1)
            starting_row = 5
            row_count = tableManpower.Rows.Count
            last_row = row_count + starting_row

            xlWorkSheet1.Cells(1, 1) = "OPMJ BY MANPOWER"
            xlWorkSheet1.Cells(2, 1) = Format(startDate, "dd-MMM-yyyy") & " until " & Format(currentDate, "dd-MMM-yyyy")
            xlWorkSheet1.Cells(3, 1) = "Printed date : " & Format(Now, "dd-MMM-yyyy HH:mm")

            Console.WriteLine("")
            Console.WriteLine("Process Export Data SHEET1")
            Console.WriteLine("Total Export Record : " & row_count)
            For i = 0 To tableManpower.Rows.Count - 1
                If tableManpower(i)("SHIFT_DATE").ToString() <> "" Then
                    xlWorkSheet1.Cells(i + starting_row, 1) = tableManpower(i)("NIK").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 2) = tableManpower(i)("NAMA").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 3) = tableManpower(i)("GRP").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 4) = tableManpower(i)("HARI_01").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 5) = tableManpower(i)("HARI_02").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 6) = tableManpower(i)("HARI_03").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 7) = tableManpower(i)("HARI_04").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 8) = tableManpower(i)("HARI_05").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 9) = tableManpower(i)("HARI_06").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 10) = tableManpower(i)("HARI_07").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 11) = tableManpower(i)("HARI_08").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 12) = tableManpower(i)("HARI_09").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 13) = tableManpower(i)("HARI_10").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 14) = tableManpower(i)("HARI_11").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 15) = tableManpower(i)("HARI_12").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 16) = tableManpower(i)("HARI_13").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 17) = tableManpower(i)("HARI_14").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 18) = tableManpower(i)("HARI_15").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 19) = tableManpower(i)("HARI_16").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 20) = tableManpower(i)("HARI_17").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 21) = tableManpower(i)("HARI_18").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 22) = tableManpower(i)("HARI_19").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 23) = tableManpower(i)("HARI_20").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 24) = tableManpower(i)("HARI_21").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 25) = tableManpower(i)("HARI_22").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 26) = tableManpower(i)("HARI_23").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 27) = tableManpower(i)("HARI_24").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 28) = tableManpower(i)("HARI_25").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 29) = tableManpower(i)("HARI_26").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 30) = tableManpower(i)("HARI_27").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 31) = tableManpower(i)("HARI_28").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 32) = tableManpower(i)("HARI_29").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 33) = tableManpower(i)("HARI_30").ToString()
                    xlWorkSheet1.Cells(i + starting_row, 34) = tableManpower(i)("HARI_31").ToString()
                End If
                Console.WriteLine("Export Record (SHEET1) OPMJ BY MANPOWER : " & i + 1)
            Next

            xlWorkSheet1.Select()
            xlWorkSheet1.Rows(tableManpower.Rows.Count + starting_row & ":1048576").Delete()
            xlWorkSheet1.cells(1, 1).select()

            '(SHEET2) OPMJ BY TIPE
            Dim xlWorkSheet2 As Object
            xlWorkSheet2 = xlWorkBook.WorkSheets(2)
            starting_row = 5
            row_count = tableType.Rows.Count
            last_row = row_count + starting_row

            xlWorkSheet2.Cells(1, 1) = "OPMJ BY TIPE"
            xlWorkSheet2.Cells(2, 1) = Format(startDate, "dd-MMM-yyyy") & " until " & Format(currentDate, "dd-MMM-yyyy")
            xlWorkSheet2.Cells(3, 1) = "Printed date : " & Format(Now, "dd-MMM-yyyy HH:mm")

            Console.WriteLine("")
            Console.WriteLine("Process Export Data SHEET1")
            Console.WriteLine("Total Export Record : " & row_count)
            For i = 0 To tableType.Rows.Count - 1
                If tableType(i)("SHIFT_DATE").ToString() <> "" Then
                    xlWorkSheet2.Cells(i + starting_row, 1) = tableType(i)("DMC_CODE").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 2) = tableType(i)("HARI_01").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 3) = tableType(i)("HARI_02").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 4) = tableType(i)("HARI_03").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 5) = tableType(i)("HARI_04").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 6) = tableType(i)("HARI_05").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 7) = tableType(i)("HARI_06").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 8) = tableType(i)("HARI_07").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 9) = tableType(i)("HARI_08").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 10) = tableType(i)("HARI_09").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 11) = tableType(i)("HARI_10").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 12) = tableType(i)("HARI_11").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 13) = tableType(i)("HARI_12").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 14) = tableType(i)("HARI_13").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 15) = tableType(i)("HARI_14").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 16) = tableType(i)("HARI_15").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 17) = tableType(i)("HARI_16").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 18) = tableType(i)("HARI_17").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 19) = tableType(i)("HARI_18").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 20) = tableType(i)("HARI_19").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 21) = tableType(i)("HARI_20").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 22) = tableType(i)("HARI_21").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 23) = tableType(i)("HARI_22").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 24) = tableType(i)("HARI_23").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 25) = tableType(i)("HARI_24").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 26) = tableType(i)("HARI_25").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 27) = tableType(i)("HARI_26").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 28) = tableType(i)("HARI_27").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 29) = tableType(i)("HARI_28").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 30) = tableType(i)("HARI_29").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 31) = tableType(i)("HARI_30").ToString()
                    xlWorkSheet2.Cells(i + starting_row, 32) = tableType(i)("HARI_31").ToString()
                End If
                Console.WriteLine("Export Record (SHEET1) OPMJ BY TIPE : " & i + 1)
            Next

            xlWorkSheet2.Select()
            xlWorkSheet2.Rows(tableType.Rows.Count + starting_row & ":1048576").Delete()
            xlWorkSheet2.cells(1, 1).select()

            xlWorkSheet1.Select()
            xlWorkSheet1.cells(1, 1).select()

            nama_file_simpan = ClsConfig.nama_file_lampiran_email & "_" & Now.ToString("yyyyMMddHHmmss")
            'xlWorkSheet1.SaveAs(lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx") 'simpan hanya 1 sheet
            xlApp.ActiveWorkbook.SaveAs(lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx") 'simpan beberapa sheet

            xlApp.Quit()
            releaseObject(xlApp)
            releaseObject(xlWorkBook)
            releaseObject(xlWorkSheet1)
            releaseObject(xlWorkSheet2)

            Console.WriteLine("")
            Console.WriteLine("##FINISH EKSPOR DATA")

            ExcelOutputFile = lokasi_simpan_file & "\" & nama_file_simpan & ".xlsx"

            send_mail(ExcelOutputFile, startDate, currentDate, currentHHmm)

        Catch ex As Exception
            ClsConfig.create_log_error("[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Ekspor Data to Excel Error")
        Environment.Exit(0)
        End Try
    End Sub

    Private Sub releaseObject(ByVal obj As Object)
        Try
            System.Runtime.InteropServices.Marshal.ReleaseComObject(obj)
            obj = Nothing
        Catch ex As Exception
            obj = Nothing
        Finally
            GC.Collect()
        End Try
    End Sub

    Private Sub send_mail(ByVal AttachedFile As String,
                          ByVal start_date As DateTime,
                          ByVal current_date As DateTime,
                          ByVal currentHHmm As Integer)
        Console.WriteLine("")
        Console.WriteLine("##START CREATE AND SEND MAIL")

        Try

            If AttachedFile = "" Then Exit Sub
            Dim tbTemp As New DataTable
            Dim query As StringBuilder = New StringBuilder()

            Dim AddressMail_To As String = ""
            Dim body_message As New StringBuilder

            If Not create_body_msg(body_message, start_date, current_date) Then Exit Sub

            query.AppendLine(" SELECT MAILADDRESS FROM Z_TANTO_LIST Where OPMJ IN ('To') ORDER BY Asc_Email_Sort DESC ")
            'query.AppendLine(" SELECT MAILADDRESS FROM Z_TANTO_LIST Where name in('Sulton') ORDER BY Asc_Email_Sort DESC ")

            tbTemp = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_TxDTIPRD)
            query.Length = 0
            query.Capacity = 0

            If tbTemp.Rows.Count > 0 Then
                Dim vw As DataView = tbTemp.DefaultView
                Dim tb As Data.DataTable = vw.ToTable()
                Dim rdr As DataTableReader = tb.CreateDataReader()
                While rdr.Read
                    AddressMail_To = rdr("MAILADDRESS") & "," & AddressMail_To
                End While
                rdr.Close()
            End If

            If Microsoft.VisualBasic.Right(Trim(AddressMail_To), 1) = "," Then AddressMail_To = Microsoft.VisualBasic.Left(AddressMail_To, Len(AddressMail_To) - 1)

            SendExcelMailViaSMTP(AddressMail_To, body_message, AttachedFile, start_date, current_date, currentHHmm)

        Catch ex As Exception
            'Panggil fungsi send email agar kirim email ulang ketika terjadi kegagal email
            send_mail(AttachedFile, start_date, current_date, currentHHmm)

            ClsConfig.create_log_error("[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Create Email Error")
            Environment.Exit(0)
        End Try

        Console.WriteLine("##FINISH CREATE AND SEND MAIL")
    End Sub

    Private Function create_body_msg(ByRef body_str As StringBuilder,
                                    ByRef startDate As DateTime,
                                    ByRef currentDate As DateTime) As Boolean

        Dim Result As Boolean = False
        Dim body_str_temp As New StringBuilder

        'If dtSource.Rows.Count > 0 Then
        Result = True
        body_str_temp.AppendLine("<html>")
        body_str_temp.AppendLine("<body>")
        body_str_temp.AppendLine("Dear All, <br />")
        body_str_temp.AppendLine("<br />")
        body_str_temp.AppendLine(String.Concat(New String() {"This is Output Man Power Per Jam (OPMJ) by period : ", start_date.ToString("dd-MMM-yyyy"), " until ", current_date.ToString("dd-MMM-yyyy"), " <br />"}))
        body_str_temp.AppendLine("Please find the attached file for detailed information <br /><br />")
        body_str_temp.AppendLine("<table style='border-collapse: collapse'>")
        body_str_temp.AppendLine("</table>")
        body_str_temp.AppendLine("</body>")
        body_str_temp.AppendLine("</html>")
        'Else
        '    Result = False
        'End If
        body_str = body_str_temp
        create_body_msg = Result
    End Function

    Private Function SendExcelMailViaSMTP(
                                            ByVal strToAddress As String,
                                            ByVal BodyMsg As StringBuilder,
                                            ByVal AttachedFile As String,
                                            ByVal start_date As DateTime,
                                            ByVal current_date As DateTime,
                                            ByVal current_HHmm1 As Integer
                                          ) As Boolean

        Dim query As StringBuilder = New StringBuilder()
        Dim email_nama As String = ClsConfig.email_nama
        Dim email_password As String = ClsConfig.email_password
        Dim email_server_smtp As String = ClsConfig.email_server_smtp
        Dim email_server_port As String = ClsConfig.email_server_port
        Dim subject_email As String = ClsConfig.subject_email
        'Dim tls_1_2 = DirectCast(3072, System.Net.SecurityProtocolType) 'TLS 1.2 //oldx`
        Dim tls As Int32 = ClsConfig.tls 'Get tls from .ini
        Dim tls_1_2 = DirectCast(tls, System.Net.SecurityProtocolType) 'TLS 1.2
        'Dim date_now As String = Format(Now)

        Dim oMail As New MailMessage()
        Dim oSmtp As New SmtpClient
        oSmtp.UseDefaultCredentials = False
        oSmtp.Credentials = New Net.NetworkCredential(email_nama, email_password)
        oSmtp.Port = CInt(email_server_port)
        oSmtp.EnableSsl = True
        oSmtp.Host = email_server_smtp

        oMail = New MailMessage()
        oMail.From = New MailAddress(email_nama)
        oMail.To.Add(strToAddress)
        oMail.Subject = subject_email & " : " & start_date.ToString("dd-MMM-yyyy") & " until " & current_date.ToString("dd-MMM-yyyy")
        oMail.IsBodyHtml = True
        oMail.Body = BodyMsg.ToString
        oMail.Attachments.Add(New Attachment(AttachedFile))
        System.Net.ServicePointManager.Expect100Continue = False
        System.Net.ServicePointManager.SecurityProtocol = tls_1_2


        'SEND EMAIL
        Try
            status_sudah_email = ClsGnrl.cek_status_sudah_email(current_date, current_HHmm1)
            If status_sudah_email = True Then
                'JIKA PROGRAM SUDAH KALKULASI DAN KIRIM EMAIL, DI-SKIP PROSES
                Console.WriteLine("PROGRAM SUDAH KALKULASI DAN KIRIM EMAIL")
            Else
                oSmtp.Send(oMail)
                ClsGnrl.monitoring_email_opmj(current_date, current_HHmm)
            End If
        Catch ex As Exception
            'Panggil fungsi send email agar kirim email ulang ketika terjadi kegagal email
            SendExcelMailViaSMTP(strToAddress, BodyMsg, AttachedFile, start_date, current_date, current_HHmm)

            ClsConfig.create_log_error("[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Send Email Error")
            Environment.Exit(0)
        End Try

    End Function

End Module
