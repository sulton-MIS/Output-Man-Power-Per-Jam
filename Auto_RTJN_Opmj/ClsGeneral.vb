﻿Imports System.Text
Public Class ClsGeneral
    'perbedaan Public "Shared" dengan "tanpa Shared", 

    'Public "Shared" fungsinya bisa di panggil tanpa harus di-initialisasi
    'contoh = ClsGeneral.get_last_date(current_date)
    'Jika menggunakan "Shared" maka tidak bisa memanggil prosedure local

    '"tanpa Shared" fungsinya bisa di panggil jika sudah initialisasi (harus di-initialisasi-kan)
    'contoh = Dim ClsGnrl As New ClsGeneral -> lalu -> ClsGnrl.cek_status_sudah_email(current_date, current_HHmm)

    Public Shared Function get_last_date(ByVal tanggal As DateTime) As Date
        Dim month_int As Integer = Month(DateAdd("m", 1, tanggal)) ' bulan + 1
        Dim year_int As Integer = Year(DateAdd("m", 1, tanggal))
        Dim Date_result As Date = DateSerial(year_int, month_int, 1) ' setting jadi tanggal 1 awal bulan berikutnya
        'get_last_date = "9/21/2021 12:00:00 AM"
        get_last_date = DateAdd("d", -1, Date_result) 'dikurangi 1 hari
    End Function

    Public Function get_kolom_HHmm(ByRef current_HHmm As Integer) As String
        Dim kolom_HHmm As String = ""

        'dimulai dari jam 08.30, karena periode jam DTI dari jam 07.30, sehingga baru mulai kalkulasi laporan awal jam 08.30
        'karena jam 07.30 masih sebagai kalkulasi laporan hari kemarin
        Select Case current_HHmm
            Case 830 To 929 'No.1 dari jam 08.30 s/d 09.29 WIB
                kolom_HHmm = "Pukul_08_30"

            Case 930 To 1029 'No.2 dari jam 09.30 s/d 10.29 WIB
                kolom_HHmm = "Pukul_09_30"

            Case 1030 To 1129 'No.3 dari jam 10.30 s/d 11.29 WIB
                kolom_HHmm = "Pukul_10_30"

            Case 1130 To 1229 'No.4 dari jam 11.30 s/d 12.29 WIB
                kolom_HHmm = "Pukul_11_30"

            Case 1230 To 1329 'No.5 dari jam 12.30 s/d 13.29 WIB
                kolom_HHmm = "Pukul_12_30"

            Case 1330 To 1429 'No.6 dari jam 13.30 s/d 14.29 WIB
                kolom_HHmm = "Pukul_13_30"

            Case 1430 To 1529 'No.7 dari jam 14.30 s/d 15.29 WIB
                kolom_HHmm = "Pukul_14_30"

            Case 1530 To 1629 'No.8 dari jam 15.30 s/d 16.29 WIB
                kolom_HHmm = "Pukul_15_30"

            Case 1630 To 1729 'No.9 dari jam 16.30 s/d 17.29 WIB
                kolom_HHmm = "Pukul_16_30"

            Case 1730 To 1829 'No.10 dari jam 17.30 s/d 18.29 WIB
                kolom_HHmm = "Pukul_17_30"

            Case 1830 To 1929 'No.11 dari jam 18.30 s/d 19.29 WIB
                kolom_HHmm = "Pukul_18_30"

            Case 1930 To 2029 'No.12 dari jam 19.30 s/d 20.29 WIB
                kolom_HHmm = "Pukul_19_30"

            Case 2030 To 2129 'No.13 dari jam 20.30 s/d 21.29 WIB
                kolom_HHmm = "Pukul_20_30"

            Case 2130 To 2229 'No.14 dari jam 21.30 s/d 22.29 WIB
                kolom_HHmm = "Pukul_21_30"

            Case 2230 To 2329 'No.15 dari jam 22.30 s/d 23.29 WIB
                kolom_HHmm = "Pukul_22_30"

            Case 2330 To 2359 'No.16 dari jam 23.30 s/d 23.59 WIB 'karena menggunakan range integer, sehingga range di pecah 2
                kolom_HHmm = "Pukul_23_30"
            Case 0 To 29 'No.16 dari jam 00.00 s/d 00.29 WIB 'karena menggunakan range integer, sehingga range di pecah 2
                kolom_HHmm = "Pukul_23_30"

            Case 30 To 129 'No.17 dari jam 00.30 s/d 01.29 WIB 'karena menggunakan range integer, sehingga range di pecah 2
                kolom_HHmm = "Pukul_00_30"

            Case 130 To 229 'No.18 dari jam 01.30 s/d 02.29 WIB
                kolom_HHmm = "Pukul_01_30"

            Case 230 To 329 'No.19 dari jam 02.30 s/d 03.29 WIB
                kolom_HHmm = "Pukul_02_30"

            Case 330 To 429 'No.20 dari jam 03.30 s/d 04.29 WIB
                kolom_HHmm = "Pukul_03_30"

            Case 430 To 529 'No.21 dari jam 04.30 s/d 05.29 WIB
                kolom_HHmm = "Pukul_04_30"

            Case 530 To 629 'No.22 dari jam 05.30 s/d 06.29 WIB
                kolom_HHmm = "Pukul_05_30"

            Case 630 To 729 'No.23 dari jam 06.30 s/d 07.29 WIB
                kolom_HHmm = "Pukul_06_30"

            Case 730 To 829 'No.24 dari jam 07.30 s/d 08.29 WIB
                kolom_HHmm = "Pukul_07_30"

            Case Else 'status email monitoring mail sender
                kolom_HHmm = "EmailMonitoring"

        End Select

        get_kolom_HHmm = kolom_HHmm
    End Function

    Public Function cek_status_sudah_email(
                                            ByVal currentDate As Date,
                                            ByRef current_HHmm As Integer) As Boolean
        Dim hasil_cek As Boolean = False
        Dim query As New StringBuilder
        Dim dt As DataTable
        Dim kolom_HHmm As String = get_kolom_HHmm(current_HHmm)

        query.AppendLine(" select ")
        query.AppendLine("     date ")
        query.AppendLine("     ," & kolom_HHmm & " ")
        query.AppendLine("     ,last_email_sent ")
        query.AppendLine(" from ")
        query.AppendLine("     ad_dis_monitoring_maintenance ")
        query.AppendLine(" where ")
        query.AppendLine("     jenis_mail_sender = 'OPMJ' ")
        query.AppendLine("     and date = '" & currentDate.ToString("yyyy-MM-dd") & "' ")
        query.AppendLine("     and " & kolom_HHmm & " = 'OK' ")
        dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
        query.Length = 0
        query.Capacity = 0

        If dt.Rows.Count > 0 Then
            hasil_cek = True
        End If

        cek_status_sudah_email = hasil_cek
    End Function

    Public Sub monitoring_email_opmj(
                        ByVal currentDate As Date,
                        ByRef current_HHmm As Integer)

        Try

            Dim query As New StringBuilder
            Dim dt As DataTable
            Dim kolom_HHmm As String = get_kolom_HHmm(current_HHmm)

            query.AppendLine(" select ")
            query.AppendLine("     date ")
            query.AppendLine("     ,last_email_sent ")
            query.AppendLine(" from ")
            query.AppendLine("     ad_dis_monitoring_maintenance ")
            query.AppendLine(" where ")
            query.AppendLine("     jenis_mail_sender = 'OPMJ' ")
            query.AppendLine("     and date = '" & currentDate.ToString("yyyy-MM-dd") & "' ")
            dt = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
            query.Length = 0
            query.Capacity = 0

            If dt.Rows.Count > 0 Then
                query.AppendLine(" update ")
                query.AppendLine("     ad_dis_monitoring_maintenance ")
                query.AppendLine(" set ")
                query.AppendLine("     " & kolom_HHmm & " = 'OK' ")
                query.AppendLine("     ,jumlah_kegagalan = ( IIF(Pukul_08_30 = 'OK', 0, 1) + IIF(Pukul_09_30 = 'OK', 0, 1) + IIF(Pukul_10_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_11_30 = 'OK', 0, 1) + IIF(Pukul_12_30 = 'OK', 0, 1) + IIF(Pukul_13_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_14_30 = 'OK', 0, 1) + IIF(Pukul_15_30 = 'OK', 0, 1) + IIF(Pukul_16_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_17_30 = 'OK', 0, 1) + IIF(Pukul_18_30 = 'OK', 0, 1) + IIF(Pukul_19_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_20_30 = 'OK', 0, 1) + IIF(Pukul_21_30 = 'OK', 0, 1) + IIF(Pukul_22_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_23_30 = 'OK', 0, 1) + IIF(Pukul_00_30 = 'OK', 0, 1) + IIF(Pukul_01_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_02_30 = 'OK', 0, 1) + IIF(Pukul_03_30 = 'OK', 0, 1) + IIF(Pukul_04_30 = 'OK', 0, 1) + ")
                query.AppendLine("                           IIF(Pukul_05_30 = 'OK', 0, 1) + IIF(Pukul_06_30 = 'OK', 0, 1) + IIF(Pukul_07_30 = 'OK', 0, 1) ) ")
                query.AppendLine("     ,last_email_sent = getdate() ")
                query.AppendLine(" where ")
                query.AppendLine("     jenis_mail_sender = 'OPMJ' ")
                query.AppendLine("     and date = '" & currentDate.ToString("yyyy-MM-dd") & "' ")
                ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
                query.Length = 0
                query.Capacity = 0
                Console.WriteLine("Proses Update Data")
            Else
                query.AppendLine(" Insert Into ")
                query.AppendLine("     ad_dis_monitoring_maintenance ")
                query.AppendLine("     ( ")
                query.AppendLine("         date ")
                query.AppendLine("         ,jenis_mail_sender ")
                query.AppendLine("         ," & kolom_HHmm & " ")
                query.AppendLine("         ,jumlah_kegagalan ")
                query.AppendLine("         ,last_email_sent ")
                query.AppendLine("     ) ")
                query.AppendLine(" values ")
                query.AppendLine("     ( ")
                query.AppendLine("         '" & currentDate.ToString("yyyy-MM-dd") & "' ")
                query.AppendLine("         ,'OPMJ' ")
                query.AppendLine("         ,'OK' ")
                query.AppendLine("         ,0 ")
                query.AppendLine("         ,getdate() ")
                query.AppendLine("     ) ")
                ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
                query.Length = 0
                query.Capacity = 0
                Console.WriteLine("Proses Input Data")
            End If

        Catch ex As Exception
            ClsConfig.create_log_error("[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Proses insert/update monitoring OPMJ error")
            Environment.Exit(0)
        End Try
    End Sub

    Public Sub StartCalc(ByVal startDate As Date, ByVal endDate As Date,
           ByRef tableManpower As DataTable,
           ByRef tableType As DataTable
        )

        tableManpower = CalculateManpower(startDate, endDate)
        Console.WriteLine("")
        Console.WriteLine("---> Calculated OPMJ By Manpower Done")
        Console.WriteLine("")

        tableType = CalculateType(startDate, endDate)
        Console.WriteLine("")
        Console.WriteLine("---> Calculated OPMJ By Type Done")
        Console.WriteLine("")
    End Sub
    Private Shared Function EscapeQuote(ByVal dmc_code As String) As String
        Return Strings.Replace(dmc_code, "'", "''", 1, -1, CompareMethod.Binary)
    End Function

    Public Function CalculateManpower(ByVal startDate As Date, ByVal endDate As Date) As DataTable
        Dim query As New StringBuilder
        Dim shift_date As String

        Try
            shift_date = startDate.ToString("yyyyMM")

            Console.WriteLine("")
            Console.WriteLine("## PROSES GET DATA FROM DATABASE")
            Console.WriteLine("")

            query.AppendLine(" 
             Select 
             (LEFT(MAX([SC].shift_date), 6)) as SHIFT_DATE, 
             [SC].NIK as NIK,
             [SC].NAMA As NAMA,  
             [SC].GRP as GRP,
             ISNULL(SUM([SC].[HARI_01]), 0) As 'HARI_01', ISNULL(SUM([SC].[HARI_02]), 0) as 'HARI_02', ISNULL(SUM([SC].[HARI_03]), 0) as 'HARI_03', ISNULL(SUM([SC].[HARI_04]), 0) as 'HARI_04', ISNULL(SUM([SC].[HARI_05]), 0) as 'HARI_05',  
             ISNULL(SUM([SC].[HARI_06]), 0) As 'HARI_06', ISNULL(SUM([SC].[HARI_07]), 0) as 'HARI_07', ISNULL(SUM([SC].[HARI_08]), 0) as 'HARI_08', ISNULL(SUM([SC].[HARI_09]), 0) as 'HARI_09', ISNULL(SUM([SC].[HARI_10]), 0) as 'HARI_10',  
             ISNULL(SUM([SC].[HARI_11]), 0) As 'HARI_11', ISNULL(SUM([SC].[HARI_12]), 0) as 'HARI_12', ISNULL(SUM([SC].[HARI_13]), 0) as 'HARI_13', ISNULL(SUM([SC].[HARI_14]), 0) as 'HARI_14', ISNULL(SUM([SC].[HARI_15]), 0) as 'HARI_15',  
             ISNULL(SUM([SC].[HARI_16]), 0) As 'HARI_16', ISNULL(SUM([SC].[HARI_17]), 0) as 'HARI_17', ISNULL(SUM([SC].[HARI_18]), 0) as 'HARI_18', ISNULL(SUM([SC].[HARI_19]), 0) as 'HARI_19', ISNULL(SUM([SC].[HARI_20]), 0) as 'HARI_20',  
             ISNULL(SUM([SC].[HARI_21]), 0) As 'HARI_21', ISNULL(SUM([SC].[HARI_22]), 0) as 'HARI_22', ISNULL(SUM([SC].[HARI_23]), 0) as 'HARI_23', ISNULL(SUM([SC].[HARI_24]), 0) as 'HARI_24', ISNULL(SUM([SC].[HARI_25]), 0) as 'HARI_25',  
             ISNULL(SUM([SC].[HARI_26]), 0) As 'HARI_26', ISNULL(SUM([SC].[HARI_27]), 0) as 'HARI_27', ISNULL(SUM([SC].[HARI_28]), 0) as 'HARI_28', ISNULL(SUM([SC].[HARI_29]), 0) as 'HARI_29', ISNULL(SUM([SC].[HARI_30]), 0) as 'HARI_30',  
             ISNULL(SUM([SC].[HARI_31]), 0) As 'HARI_31'  
             FROM(
             SELECT 
             [J_KOTEI].shift_date as SHIFT_DATE,
             [J_KOTEI].NIK,
             [J_KOTEI].NAMA,
             [J_KOTEI].GRP,
             CASE WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '01') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_01', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '02') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_02', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '03') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_03', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '04') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_04', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '05') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_05', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '06') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_06', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '07') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_07', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '08') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_08', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '09') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_09', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '10') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_10', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '11') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_11', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '12') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_12', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '13') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_13', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '14') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_14', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '15') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_15', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '16') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_16', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '17') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_17', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '18') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_18', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '19') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_19', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '20') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_20', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '21') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_21', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '22') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_22', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '23') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_23', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '24') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_24', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '25') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_25', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '26') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_26', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '27') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_27', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '28') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_28', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '29') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_29', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '30') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_30', 
             Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '31') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_31'  
             FROM(
             SELECT  
             dbo.Z_RT_master_sagyosha.id_sagyosha AS NIK,
             dbo.Z_RT_master_sagyosha.name_sagyosha As NAMA, 
             dbo.Z_RT_master_sagyosha.grp AS GRP,
             Z_RT_data_J_kotei.id_hinmoku As TIPE, 
             Z_RT_data_J_kotei.shift_date AS SHIFT_DATE,
             SUM(amnt_OK) + SUM(amnt_NG) + SUM(amnt_PND) As QTY_INPUT, 
             Case WHEN Z_RT_data_J_kotei.id_hinmoku Like '%-S' then Z_PRTS.SIYOUW else 1 end  AS CAVITY, 
             (SUM(amnt_OK) + SUM(amnt_NG) + SUM(amnt_PND)) * CASE WHEN Z_RT_data_J_kotei.id_hinmoku Like '%-S' then Z_PRTS.SIYOUW else 1 end AS QTY_INPUT_PCS, 
             ((SUM(ttl_sagyo)) - (ISNULL(SUM(J_PAUSE_WORKING.ttl_pause_working),0)) - (ISNULL(SUM(J_SETTING_AFTER.ttl_setting_after), 0))) AS WORKING_TIME 
             FROM
                    dbo.Z_RT_data_J_kotei
                    INNER Join  
             dbo.Z_RT_data_J_sagyosha ON dbo.Z_RT_data_J_kotei.id_seisan = dbo.Z_RT_data_J_sagyosha.id_seisan  
             And dbo.Z_RT_data_J_kotei.id_kotei = dbo.Z_RT_data_J_sagyosha.id_kotei And dbo.Z_RT_data_J_kotei.id_kikai = dbo.Z_RT_data_J_sagyosha.id_kikai  
             And dbo.Z_RT_data_J_kotei.bunban = dbo.Z_RT_data_J_sagyosha.bunban  
             INNER Join  
             dbo.Z_RT_master_sagyosha ON dbo.Z_RT_data_J_sagyosha.id_sagyosha = dbo.Z_RT_master_sagyosha.id_sagyosha  
             INNER Join  
             dbo.Z_RT_master_kotei ON dbo.Z_RT_data_J_kotei.id_kotei = dbo.Z_RT_master_kotei.id_kotei  
             Left OUTER JOIN (Select id_seisan, id_kotei, id_kikai, id_maejotai, bunban, SUM(sbttl_jotai) AS ttl_pause_working  
             From dbo.Z_RT_data_J_kikai AS Z_RT_data_J_kikai_1  
             Where (id_Remarks <> 6) And (id_Remarks <> 2) And (id_Remarks <> 24)
             And (id_Remarks <> 23) And (id_Remarks <> 32) And (id_Remarks <> 111) And (id_Remarks <> 19)  
             And (id_Remarks <> 38) And (id_Remarks <> 9) And (id_Remarks <> 10)  
             And (id_Remarks <> 8) And (id_Remarks <> 11) And (id_Remarks <> 5) And (id_Remarks <> 44) And (id_Remarks <> 17)  
             And (id_Remarks <> 45) And (id_Remarks <> 31) And (id_Remarks <> 15) And (id_Remarks <> 129)  
             GROUP BY id_seisan, id_kotei, id_kikai, id_maejotai, bunban  
             HAVING(id_maejotai = 5)) As J_PAUSE_WORKING  
             On Z_RT_data_J_kotei.id_seisan = J_PAUSE_WORKING.id_seisan And Z_RT_data_J_kotei.id_kotei = J_PAUSE_WORKING.id_kotei And  
             Z_RT_data_J_kotei.id_kikai = J_PAUSE_WORKING.id_kikai And Z_RT_data_J_kotei.bunban = J_PAUSE_WORKING.bunban
                Left OUTER JOIN (Select id_seisan, id_kotei, id_kikai, id_maejotai, bunban, SUM(sbttl_jotai) AS ttl_setting_after  
             From dbo.Z_RT_data_J_kikai AS Z_RT_data_J_kikai_1  
             Where (id_Remarks = 6) And (id_Remarks = 2) And (id_Remarks = 24)
             Group By id_seisan, id_kotei, id_kikai, id_maejotai, bunban  
             HAVING(id_maejotai = 5)) As J_SETTING_AFTER  
             On Z_RT_data_J_kotei.id_seisan = J_SETTING_AFTER.id_seisan And Z_RT_data_J_kotei.id_kotei = J_SETTING_AFTER.id_kotei And Z_RT_data_J_kotei.id_kikai = J_SETTING_AFTER.id_kikai And  
             Z_RT_data_J_kotei.bunban = J_SETTING_AFTER.bunban
                Left OUTER JOIN Z_PRTS ON Z_RT_data_J_kotei.id_hinmoku = Z_PRTS.KCODE And (Z_PRTS.SDATE Like '000000001' OR Z_PRTS.SDATE LIKE '000000011') AND (EDATE like '999999991' OR EDATE like '999999999') 
                WHERE  1 = 1 
             And dbo.Z_RT_data_J_kotei.shift_date Like '%" & shift_date & "%' 
             And Z_RT_master_kotei.name_kotei in ( 
             'Double Sheet', 
             'Hariawase', 
             'Hariawase Awal', 
             'Hariawase Polycarbon', 
             'Pasang Ag Protection Sht', 
             'Pasang Anti Bacteri Film', 
             'Pasang EMI Shield', 
             'Pasang Overlay', 
             'Pasang Smoke Sheet', 
             'Pasang UV Cut Film') 
             And Z_RT_data_J_kotei.flg_sagyokanryo = 1 
             And dbo.Z_RT_master_sagyosha.flg_opmj = 1 
             GROUP BY  
             Z_RT_master_sagyosha.id_sagyosha, Z_RT_master_sagyosha.name_sagyosha, Z_RT_master_sagyosha.grp, shift_date, Z_RT_data_J_kotei.id_hinmoku, Z_PRTS.SIYOUW 
  
             ) as [J_KOTEI] 
             GROUP BY [J_KOTEI].NIK, [J_KOTEI].NAMA, [J_KOTEI].GRP, [J_KOTEI].SHIFT_DATE 
             ) AS [SC] 
             WHERE 1 = 1 
             GROUP BY [SC].NIK, [SC].NAMA, [SC].GRP 
             ORDER BY [SC].NIK 
         ")

        Dim GetManpower As DataTable = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_RTJN_PRD)


        query.Length = 0
        query.Capacity = 0

        Console.WriteLine("## FINISH GET DATA")
        Console.WriteLine("")
        Console.WriteLine("Total Calculation Record : " & GetManpower.Rows.Count)
        Console.WriteLine("")

        Console.WriteLine("## START CALCULATION")

            If GetManpower.Rows.Count > 0 Then

                Dim day = startDate
                Dim end_day = endDate

                While day <= end_day
                    For i = 0 To GetManpower.Rows.Count - 1
                        If day.ToString("yyyyMM") = GetManpower(i)("SHIFT_DATE") Then
                            Dim str_shiftdate As String = day.ToString("yyyyMM")
                            Dim NIK As String = GetManpower(i)("NIK")
                            Dim NAMA As String = GetManpower(i)("NAMA")
                            Dim GRP As String = GetManpower(i)("GRP")
                            Dim HARI_01 As Decimal = GetManpower(i)("HARI_01")
                            Dim HARI_02 As Decimal = GetManpower(i)("HARI_02")
                            Dim HARI_03 As Decimal = GetManpower(i)("HARI_03")
                            Dim HARI_04 As Decimal = GetManpower(i)("HARI_04")
                            Dim HARI_05 As Decimal = GetManpower(i)("HARI_05")
                            Dim HARI_06 As Decimal = GetManpower(i)("HARI_06")
                            Dim HARI_07 As Decimal = GetManpower(i)("HARI_07")
                            Dim HARI_08 As Decimal = GetManpower(i)("HARI_08")
                            Dim HARI_09 As Decimal = GetManpower(i)("HARI_09")
                            Dim HARI_10 As Decimal = GetManpower(i)("HARI_10")
                            Dim HARI_11 As Decimal = GetManpower(i)("HARI_11")
                            Dim HARI_12 As Decimal = GetManpower(i)("HARI_12")
                            Dim HARI_13 As Decimal = GetManpower(i)("HARI_13")
                            Dim HARI_14 As Decimal = GetManpower(i)("HARI_14")
                            Dim HARI_15 As Decimal = GetManpower(i)("HARI_15")
                            Dim HARI_16 As Decimal = GetManpower(i)("HARI_16")
                            Dim HARI_17 As Decimal = GetManpower(i)("HARI_17")
                            Dim HARI_18 As Decimal = GetManpower(i)("HARI_18")
                            Dim HARI_19 As Decimal = GetManpower(i)("HARI_19")
                            Dim HARI_20 As Decimal = GetManpower(i)("HARI_20")
                            Dim HARI_21 As Decimal = GetManpower(i)("HARI_21")
                            Dim HARI_22 As Decimal = GetManpower(i)("HARI_22")
                            Dim HARI_23 As Decimal = GetManpower(i)("HARI_23")
                            Dim HARI_24 As Decimal = GetManpower(i)("HARI_24")
                            Dim HARI_25 As Decimal = GetManpower(i)("HARI_25")
                            Dim HARI_26 As Decimal = GetManpower(i)("HARI_26")
                            Dim HARI_27 As Decimal = GetManpower(i)("HARI_27")
                            Dim HARI_28 As Decimal = GetManpower(i)("HARI_28")
                            Dim HARI_29 As Decimal = GetManpower(i)("HARI_29")
                            Dim HARI_30 As Decimal = GetManpower(i)("HARI_30")
                            Dim HARI_31 As Decimal = GetManpower(i)("HARI_31")

                            query.AppendLine("  Select ")
                            query.AppendLine("          id ")
                            query.AppendLine("         ,shift_date")
                            query.AppendLine("         ,nik ")
                            query.AppendLine("         ,nama ")
                            query.AppendLine("         ,grp ")
                            query.AppendLine("         ,hari_01 ")
                            query.AppendLine("         ,hari_02 ")
                            query.AppendLine("         ,hari_03 ")
                            query.AppendLine("         ,hari_04 ")
                            query.AppendLine("         ,hari_05 ")
                            query.AppendLine("         ,hari_06 ")
                            query.AppendLine("         ,hari_07 ")
                            query.AppendLine("         ,hari_08 ")
                            query.AppendLine("         ,hari_09 ")
                            query.AppendLine("         ,hari_10 ")
                            query.AppendLine("         ,hari_11 ")
                            query.AppendLine("         ,hari_12 ")
                            query.AppendLine("         ,hari_13 ")
                            query.AppendLine("         ,hari_14 ")
                            query.AppendLine("         ,hari_15 ")
                            query.AppendLine("         ,hari_16 ")
                            query.AppendLine("         ,hari_17 ")
                            query.AppendLine("         ,hari_18 ")
                            query.AppendLine("         ,hari_19 ")
                            query.AppendLine("         ,hari_20 ")
                            query.AppendLine("         ,hari_21 ")
                            query.AppendLine("         ,hari_22 ")
                            query.AppendLine("         ,hari_23 ")
                            query.AppendLine("         ,hari_24 ")
                            query.AppendLine("         ,hari_25 ")
                            query.AppendLine("         ,hari_26 ")
                            query.AppendLine("         ,hari_27 ")
                            query.AppendLine("         ,hari_28 ")
                            query.AppendLine("         ,hari_29 ")
                            query.AppendLine("         ,hari_30 ")
                            query.AppendLine("         ,hari_31 ")
                            query.AppendLine("  From ")
                            query.AppendLine("      ad_dis_rtjn_sum_qty_opmj ")
                            query.AppendLine(" where ")
                            query.AppendLine("     nik='" & NIK & "' ")
                            query.AppendLine("     and shift_date='" & str_shiftdate & "' ")

                            Dim GetTblManpower As DataTable = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)

                            query.Length = 0
                            query.Capacity = 0

                            If GetTblManpower.Rows.Count > 0 Then
                                query.AppendLine(" update ")
                                query.AppendLine("     ad_dis_rtjn_sum_qty_opmj ")
                                query.AppendLine(" set ")
                                query.AppendLine("     nik='" & NIK & "' ")
                                query.AppendLine("     ,nama='" & EscapeQuote(NAMA) & "' ")
                                query.AppendLine("     ,grp='" & GRP & "' ")
                                query.AppendLine("     ,hari_01='" & HARI_01 & "' ")
                                query.AppendLine("     ,hari_02='" & HARI_02 & "' ")
                                query.AppendLine("     ,hari_03='" & HARI_03 & "' ")
                                query.AppendLine("     ,hari_04='" & HARI_04 & "' ")
                                query.AppendLine("     ,hari_05='" & HARI_05 & "' ")
                                query.AppendLine("     ,hari_06='" & HARI_06 & "' ")
                                query.AppendLine("     ,hari_07='" & HARI_07 & "' ")
                                query.AppendLine("     ,hari_08='" & HARI_08 & "' ")
                                query.AppendLine("     ,hari_09='" & HARI_09 & "' ")
                                query.AppendLine("     ,hari_10='" & HARI_10 & "' ")
                                query.AppendLine("     ,hari_11='" & HARI_11 & "' ")
                                query.AppendLine("     ,hari_12='" & HARI_12 & "' ")
                                query.AppendLine("     ,hari_13='" & HARI_13 & "' ")
                                query.AppendLine("     ,hari_14='" & HARI_14 & "' ")
                                query.AppendLine("     ,hari_15='" & HARI_15 & "' ")
                                query.AppendLine("     ,hari_16='" & HARI_16 & "' ")
                                query.AppendLine("     ,hari_17='" & HARI_17 & "' ")
                                query.AppendLine("     ,hari_18='" & HARI_18 & "' ")
                                query.AppendLine("     ,hari_19='" & HARI_19 & "' ")
                                query.AppendLine("     ,hari_20='" & HARI_20 & "' ")
                                query.AppendLine("     ,hari_21='" & HARI_21 & "' ")
                                query.AppendLine("     ,hari_22='" & HARI_22 & "' ")
                                query.AppendLine("     ,hari_23='" & HARI_23 & "' ")
                                query.AppendLine("     ,hari_24='" & HARI_24 & "' ")
                                query.AppendLine("     ,hari_25='" & HARI_25 & "' ")
                                query.AppendLine("     ,hari_26='" & HARI_26 & "' ")
                                query.AppendLine("     ,hari_27='" & HARI_27 & "' ")
                                query.AppendLine("     ,hari_28='" & HARI_28 & "' ")
                                query.AppendLine("     ,hari_29='" & HARI_29 & "' ")
                                query.AppendLine("     ,hari_30='" & HARI_30 & "' ")
                                query.AppendLine("     ,hari_31='" & HARI_31 & "' ")
                                query.AppendLine(" where ")
                                query.AppendLine("     nik='" & NIK & "' ")
                                query.AppendLine("     and shift_date='" & str_shiftdate & "' ")

                                ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)

                                query.Length = 0
                                query.Capacity = 0

                                Console.WriteLine("Proses Update Data OPMJ By Manpower : " & EscapeQuote(NAMA) & "")
                            Else
                                query.AppendLine(" Insert Into ")
                                query.AppendLine("     ad_dis_rtjn_sum_qty_opmj ")
                                query.AppendLine("     ( ")
                                query.AppendLine("         shift_date")
                                query.AppendLine("         ,nik ")
                                query.AppendLine("         ,nama ")
                                query.AppendLine("         ,grp ")
                                query.AppendLine("         ,hari_01 ")
                                query.AppendLine("         ,hari_02 ")
                                query.AppendLine("         ,hari_03 ")
                                query.AppendLine("         ,hari_04 ")
                                query.AppendLine("         ,hari_05 ")
                                query.AppendLine("         ,hari_06 ")
                                query.AppendLine("         ,hari_07 ")
                                query.AppendLine("         ,hari_08 ")
                                query.AppendLine("         ,hari_09 ")
                                query.AppendLine("         ,hari_10 ")
                                query.AppendLine("         ,hari_11 ")
                                query.AppendLine("         ,hari_12 ")
                                query.AppendLine("         ,hari_13 ")
                                query.AppendLine("         ,hari_14 ")
                                query.AppendLine("         ,hari_15 ")
                                query.AppendLine("         ,hari_16 ")
                                query.AppendLine("         ,hari_17 ")
                                query.AppendLine("         ,hari_18 ")
                                query.AppendLine("         ,hari_19 ")
                                query.AppendLine("         ,hari_20 ")
                                query.AppendLine("         ,hari_21 ")
                                query.AppendLine("         ,hari_22 ")
                                query.AppendLine("         ,hari_23 ")
                                query.AppendLine("         ,hari_24 ")
                                query.AppendLine("         ,hari_25 ")
                                query.AppendLine("         ,hari_26 ")
                                query.AppendLine("         ,hari_27 ")
                                query.AppendLine("         ,hari_28 ")
                                query.AppendLine("         ,hari_29 ")
                                query.AppendLine("         ,hari_30 ")
                                query.AppendLine("         ,hari_31 ")
                                query.AppendLine("     ) ")
                                query.AppendLine(" values ")
                                query.AppendLine("     ( ")
                                query.AppendLine("     '" & str_shiftdate & "'")
                                query.AppendLine("     ,'" & NIK & "' ")
                                query.AppendLine("     ,'" & EscapeQuote(NAMA) & "' ")
                                query.AppendLine("     ,'" & GRP & "' ")
                                query.AppendLine("     ,'" & HARI_01 & "' ")
                                query.AppendLine("     ,'" & HARI_02 & "' ")
                                query.AppendLine("     ,'" & HARI_03 & "' ")
                                query.AppendLine("     ,'" & HARI_04 & "' ")
                                query.AppendLine("     ,'" & HARI_05 & "' ")
                                query.AppendLine("     ,'" & HARI_06 & "' ")
                                query.AppendLine("     ,'" & HARI_07 & "' ")
                                query.AppendLine("     ,'" & HARI_08 & "' ")
                                query.AppendLine("     ,'" & HARI_09 & "' ")
                                query.AppendLine("     ,'" & HARI_10 & "' ")
                                query.AppendLine("     ,'" & HARI_11 & "' ")
                                query.AppendLine("     ,'" & HARI_12 & "' ")
                                query.AppendLine("     ,'" & HARI_13 & "' ")
                                query.AppendLine("     ,'" & HARI_14 & "' ")
                                query.AppendLine("     ,'" & HARI_15 & "' ")
                                query.AppendLine("     ,'" & HARI_16 & "' ")
                                query.AppendLine("     ,'" & HARI_17 & "' ")
                                query.AppendLine("     ,'" & HARI_18 & "' ")
                                query.AppendLine("     ,'" & HARI_19 & "' ")
                                query.AppendLine("     ,'" & HARI_20 & "' ")
                                query.AppendLine("     ,'" & HARI_21 & "' ")
                                query.AppendLine("     ,'" & HARI_22 & "' ")
                                query.AppendLine("     ,'" & HARI_23 & "' ")
                                query.AppendLine("     ,'" & HARI_24 & "' ")
                                query.AppendLine("     ,'" & HARI_25 & "' ")
                                query.AppendLine("     ,'" & HARI_26 & "' ")
                                query.AppendLine("     ,'" & HARI_27 & "' ")
                                query.AppendLine("     ,'" & HARI_28 & "' ")
                                query.AppendLine("     ,'" & HARI_29 & "' ")
                                query.AppendLine("     ,'" & HARI_30 & "' ")
                                query.AppendLine("     ,'" & HARI_31 & "' ")
                                query.AppendLine("     ) ")
                                ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
                                query.Length = 0
                                query.Capacity = 0
                                Console.WriteLine("Proses Insert Data OPMJ By Manpower : " & EscapeQuote(NAMA) & "")
                            End If
                        End If
                    Next

                    day = day.AddDays(1)

                End While
            End If

        Catch ex As Exception
            ClsConfig.create_log_error("[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Calculation OPMJ By Manpower Error")
            Environment.Exit(0)
        End Try

        query.AppendLine(" select ")
        query.AppendLine("     id ")
        query.AppendLine("     ,shift_date ")
        query.AppendLine("     ,nik ")
        query.AppendLine("     ,nama ")
        query.AppendLine("     ,grp ")
        query.AppendLine("     ,hari_01 ")
        query.AppendLine("     ,hari_02 ")
        query.AppendLine("     ,hari_03 ")
        query.AppendLine("     ,hari_04 ")
        query.AppendLine("     ,hari_05 ")
        query.AppendLine("     ,hari_06 ")
        query.AppendLine("     ,hari_07 ")
        query.AppendLine("     ,hari_08 ")
        query.AppendLine("     ,hari_09 ")
        query.AppendLine("     ,hari_10 ")
        query.AppendLine("     ,hari_11 ")
        query.AppendLine("     ,hari_12 ")
        query.AppendLine("     ,hari_13 ")
        query.AppendLine("     ,hari_14 ")
        query.AppendLine("     ,hari_15 ")
        query.AppendLine("     ,hari_16 ")
        query.AppendLine("     ,hari_17 ")
        query.AppendLine("     ,hari_18 ")
        query.AppendLine("     ,hari_19 ")
        query.AppendLine("     ,hari_20 ")
        query.AppendLine("     ,hari_21 ")
        query.AppendLine("     ,hari_22 ")
        query.AppendLine("     ,hari_23 ")
        query.AppendLine("     ,hari_24 ")
        query.AppendLine("     ,hari_25 ")
        query.AppendLine("     ,hari_26 ")
        query.AppendLine("     ,hari_27 ")
        query.AppendLine("     ,hari_28 ")
        query.AppendLine("     ,hari_29 ")
        query.AppendLine("     ,hari_30 ")
        query.AppendLine("     ,hari_31 ")
        query.AppendLine(" from ")
        query.AppendLine("     ad_dis_rtjn_sum_qty_opmj ")
        query.AppendLine(" where ")
        query.AppendLine("     shift_date between '" + startDate.ToString("yyyyMM") + "' and '" + endDate.ToString("yyyyMM") + "' ")

        CalculateManpower = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)

    End Function

    Public Function CalculateType(ByVal startDate As Date, ByVal endDate As Date) As DataTable
        Dim query As New StringBuilder
        Dim shift_date As String

        Try

            shift_date = startDate.ToString("yyyyMM")

        query.AppendLine(" 
            Select
            (LEFT(MAX([SC].shift_date), 6)) as SHIFT_DATE, 
            [SC].DMC_CODE,
            ISNULL(SUM([SC].[HARI_01]), 0) As 'HARI_01', ISNULL(SUM([SC].[HARI_02]), 0) as 'HARI_02', ISNULL(SUM([SC].[HARI_03]), 0) as 'HARI_03', ISNULL(SUM([SC].[HARI_04]), 0) as 'HARI_04', ISNULL(SUM([SC].[HARI_05]), 0) as 'HARI_05',  
            ISNULL(SUM([SC].[HARI_06]), 0) As 'HARI_06', ISNULL(SUM([SC].[HARI_07]), 0) as 'HARI_07', ISNULL(SUM([SC].[HARI_08]), 0) as 'HARI_08', ISNULL(SUM([SC].[HARI_09]), 0) as 'HARI_09', ISNULL(SUM([SC].[HARI_10]), 0) as 'HARI_10',  
            ISNULL(SUM([SC].[HARI_11]), 0) As 'HARI_11', ISNULL(SUM([SC].[HARI_12]), 0) as 'HARI_12', ISNULL(SUM([SC].[HARI_13]), 0) as 'HARI_13', ISNULL(SUM([SC].[HARI_14]), 0) as 'HARI_14', ISNULL(SUM([SC].[HARI_15]), 0) as 'HARI_15',  
            ISNULL(SUM([SC].[HARI_16]), 0) As 'HARI_16', ISNULL(SUM([SC].[HARI_17]), 0) as 'HARI_17', ISNULL(SUM([SC].[HARI_18]), 0) as 'HARI_18', ISNULL(SUM([SC].[HARI_19]), 0) as 'HARI_19', ISNULL(SUM([SC].[HARI_20]), 0) as 'HARI_20',  
            ISNULL(SUM([SC].[HARI_21]), 0) As 'HARI_21', ISNULL(SUM([SC].[HARI_22]), 0) as 'HARI_22', ISNULL(SUM([SC].[HARI_23]), 0) as 'HARI_23', ISNULL(SUM([SC].[HARI_24]), 0) as 'HARI_24', ISNULL(SUM([SC].[HARI_25]), 0) as 'HARI_25',  
            ISNULL(SUM([SC].[HARI_26]), 0) As 'HARI_26', ISNULL(SUM([SC].[HARI_27]), 0) as 'HARI_27', ISNULL(SUM([SC].[HARI_28]), 0) as 'HARI_28', ISNULL(SUM([SC].[HARI_29]), 0) as 'HARI_29', ISNULL(SUM([SC].[HARI_30]), 0) as 'HARI_30',  
            ISNULL(SUM([SC].[HARI_31]), 0) As 'HARI_31'  
            FROM(
            SELECT 
            [J_KOTEI].shift_date as SHIFT_DATE,
            [J_KOTEI].DMC_CODE As DMC_CODE,
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '01') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_01', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '02') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_02', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '03') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_03', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '04') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_04', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '05') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_05', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '06') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_06', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '07') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_07', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '08') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_08', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '09') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_09', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '10') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_10', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '11') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_11', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '12') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_12', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '13') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_13', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '14') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_14', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '15') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_15', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '16') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_16', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '17') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_17', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '18') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_18', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '19') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_19', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '20') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_20', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '21') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_21', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '22') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_22', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '23') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_23', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '24') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_24', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '25') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_25', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '26') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_26', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '27') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_27', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '28') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_28', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '29') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_29', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '30') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_30', 
            Case WHEN (RIGHT([J_KOTEI].SHIFT_DATE, 2) = '31') THEN CAST((SUM([J_KOTEI].QTY_INPUT_PCS) / NULLIF((SUM([J_KOTEI].WORKING_TIME) / 3600 ), 0)) as numeric(36,1)) ELSE 0 end AS 'HARI_31'  
            FROM(
            SELECT  
            Z_RT_data_J_kotei.id_seihin AS DMC_CODE,
            Z_RT_data_J_kotei.shift_date As SHIFT_DATE, 
            SUM(amnt_OK) + SUM(amnt_NG) + SUM(amnt_PND) As QTY_INPUT, 
            Case WHEN Z_RT_data_J_kotei.id_hinmoku Like '%-S' then Z_PRTS.SIYOUW else 1 end  AS CAVITY, 
            (SUM(amnt_OK) + SUM(amnt_NG) + SUM(amnt_PND)) * CASE WHEN Z_RT_data_J_kotei.id_hinmoku Like '%-S' then Z_PRTS.SIYOUW else 1 end AS QTY_INPUT_PCS, 
            ((SUM(ttl_sagyo)) - (ISNULL(SUM(J_PAUSE_WORKING.ttl_pause_working),0)) - (ISNULL(SUM(J_SETTING_AFTER.ttl_setting_after), 0))) AS WORKING_TIME 
            FROM
            dbo.Z_RT_data_J_kotei
            INNER Join  
            dbo.Z_RT_data_J_sagyosha ON dbo.Z_RT_data_J_kotei.id_seisan = dbo.Z_RT_data_J_sagyosha.id_seisan  
            And dbo.Z_RT_data_J_kotei.id_kotei = dbo.Z_RT_data_J_sagyosha.id_kotei And dbo.Z_RT_data_J_kotei.id_kikai = dbo.Z_RT_data_J_sagyosha.id_kikai  
            And dbo.Z_RT_data_J_kotei.bunban = dbo.Z_RT_data_J_sagyosha.bunban  
            INNER Join  
            dbo.Z_RT_master_sagyosha ON dbo.Z_RT_data_J_sagyosha.id_sagyosha = dbo.Z_RT_master_sagyosha.id_sagyosha  
            INNER Join  
            dbo.Z_RT_master_kotei ON dbo.Z_RT_data_J_kotei.id_kotei = dbo.Z_RT_master_kotei.id_kotei  
            Left OUTER JOIN (Select id_seisan, id_kotei, id_kikai, id_maejotai, bunban, SUM(sbttl_jotai) AS ttl_pause_working  
            From dbo.Z_RT_data_J_kikai AS Z_RT_data_J_kikai_1  
            Where (id_Remarks <> 6) And (id_Remarks <> 2) And (id_Remarks <> 24)
            And (id_Remarks <> 23) And (id_Remarks <> 32) And (id_Remarks <> 111) And (id_Remarks <> 19)  
            And (id_Remarks <> 38) And (id_Remarks <> 9) And (id_Remarks <> 10)  
            And (id_Remarks <> 8) And (id_Remarks <> 11) And (id_Remarks <> 5) And (id_Remarks <> 44) And (id_Remarks <> 17)  
            And (id_Remarks <> 45) And (id_Remarks <> 31) And (id_Remarks <> 15) And (id_Remarks <> 129)  
            GROUP BY id_seisan, id_kotei, id_kikai, id_maejotai, bunban  
            HAVING(id_maejotai = 5)) As J_PAUSE_WORKING  
            On Z_RT_data_J_kotei.id_seisan = J_PAUSE_WORKING.id_seisan And Z_RT_data_J_kotei.id_kotei = J_PAUSE_WORKING.id_kotei And  
            Z_RT_data_J_kotei.id_kikai = J_PAUSE_WORKING.id_kikai And Z_RT_data_J_kotei.bunban = J_PAUSE_WORKING.bunban
            Left OUTER JOIN (Select id_seisan, id_kotei, id_kikai, id_maejotai, bunban, SUM(sbttl_jotai) AS ttl_setting_after  
            From dbo.Z_RT_data_J_kikai AS Z_RT_data_J_kikai_1  
            Where (id_Remarks = 6) And (id_Remarks = 2) And (id_Remarks = 24)
            Group By id_seisan, id_kotei, id_kikai, id_maejotai, bunban  
            HAVING(id_maejotai = 5)) As J_SETTING_AFTER  
            On Z_RT_data_J_kotei.id_seisan = J_SETTING_AFTER.id_seisan And Z_RT_data_J_kotei.id_kotei = J_SETTING_AFTER.id_kotei And Z_RT_data_J_kotei.id_kikai = J_SETTING_AFTER.id_kikai And  
            Z_RT_data_J_kotei.bunban = J_SETTING_AFTER.bunban
            Left OUTER JOIN Z_PRTS ON Z_RT_data_J_kotei.id_hinmoku = Z_PRTS.KCODE And (Z_PRTS.SDATE Like '000000001' OR Z_PRTS.SDATE LIKE '000000011') AND (EDATE like '999999991' OR EDATE like '999999999') 
            WHERE  1 = 1 
            And dbo.Z_RT_data_J_kotei.shift_date Like '%" & shift_date & "%' 
            And Z_RT_master_kotei.name_kotei in ( 
            'Double Sheet', 
            'Hariawase', 
            'Hariawase Awal', 
            'Hariawase Polycarbon', 
            'Pasang Ag Protection Sht', 
            'Pasang Anti Bacteri Film', 
            'Pasang EMI Shield', 
            'Pasang Overlay', 
            'Pasang Smoke Sheet', 
            'Pasang UV Cut Film') 
            And Z_RT_data_J_kotei.flg_sagyokanryo = 1 
            And dbo.Z_RT_master_sagyosha.flg_opmj = 1 
            GROUP BY  
            Z_RT_master_sagyosha.id_sagyosha, Z_RT_master_sagyosha.name_sagyosha, Z_RT_master_sagyosha.grp, shift_date, Z_RT_data_J_kotei.id_hinmoku, Z_PRTS.SIYOUW, Z_RT_data_J_kotei.id_seihin 

            ) as [J_KOTEI] 
            GROUP BY  [J_KOTEI].SHIFT_DATE, [J_KOTEI].DMC_CODE
            ) AS [SC] 
            WHERE 1 = 1 
            GROUP BY [SC].DMC_CODE
            ORDER BY [SC].DMC_CODE 
         ")

        Dim GetTipe As DataTable = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_RTJN_PRD)

        query.Length = 0
        query.Capacity = 0

        Console.WriteLine("## FINISH GET DATA")
        Console.WriteLine("")
        Console.WriteLine("Total Calculation Record : " & GetTipe.Rows.Count)
        Console.WriteLine("")

        Console.WriteLine("## START CALCULATION")


        If GetTipe.Rows.Count > 0 Then

            Dim day = startDate
            Dim end_day = endDate

            While day <= end_day
                For i = 0 To GetTipe.Rows.Count - 1
                    If day.ToString("yyyyMM") = GetTipe(i)("SHIFT_DATE") Then
                        Dim str_shiftdate As String = day.ToString("yyyyMM")
                        Dim DMC_CODE As String = GetTipe(i)("DMC_CODE")
                        Dim HARI_01 As Decimal = GetTipe(i)("HARI_01")
                        Dim HARI_02 As Decimal = GetTipe(i)("HARI_02")
                        Dim HARI_03 As Decimal = GetTipe(i)("HARI_03")
                        Dim HARI_04 As Decimal = GetTipe(i)("HARI_04")
                        Dim HARI_05 As Decimal = GetTipe(i)("HARI_05")
                        Dim HARI_06 As Decimal = GetTipe(i)("HARI_06")
                        Dim HARI_07 As Decimal = GetTipe(i)("HARI_07")
                        Dim HARI_08 As Decimal = GetTipe(i)("HARI_08")
                        Dim HARI_09 As Decimal = GetTipe(i)("HARI_09")
                        Dim HARI_10 As Decimal = GetTipe(i)("HARI_10")
                        Dim HARI_11 As Decimal = GetTipe(i)("HARI_11")
                        Dim HARI_12 As Decimal = GetTipe(i)("HARI_12")
                        Dim HARI_13 As Decimal = GetTipe(i)("HARI_13")
                        Dim HARI_14 As Decimal = GetTipe(i)("HARI_14")
                        Dim HARI_15 As Decimal = GetTipe(i)("HARI_15")
                        Dim HARI_16 As Decimal = GetTipe(i)("HARI_16")
                        Dim HARI_17 As Decimal = GetTipe(i)("HARI_17")
                        Dim HARI_18 As Decimal = GetTipe(i)("HARI_18")
                        Dim HARI_19 As Decimal = GetTipe(i)("HARI_19")
                        Dim HARI_20 As Decimal = GetTipe(i)("HARI_20")
                        Dim HARI_21 As Decimal = GetTipe(i)("HARI_21")
                        Dim HARI_22 As Decimal = GetTipe(i)("HARI_22")
                        Dim HARI_23 As Decimal = GetTipe(i)("HARI_23")
                        Dim HARI_24 As Decimal = GetTipe(i)("HARI_24")
                        Dim HARI_25 As Decimal = GetTipe(i)("HARI_25")
                        Dim HARI_26 As Decimal = GetTipe(i)("HARI_26")
                        Dim HARI_27 As Decimal = GetTipe(i)("HARI_27")
                        Dim HARI_28 As Decimal = GetTipe(i)("HARI_28")
                        Dim HARI_29 As Decimal = GetTipe(i)("HARI_29")
                        Dim HARI_30 As Decimal = GetTipe(i)("HARI_30")
                        Dim HARI_31 As Decimal = GetTipe(i)("HARI_31")

                        query.AppendLine("  Select ")
                        query.AppendLine("          id ")
                        query.AppendLine("         ,shift_date")
                        query.AppendLine("         ,dmc_code ")
                        query.AppendLine("         ,hari_01 ")
                        query.AppendLine("         ,hari_02 ")
                        query.AppendLine("         ,hari_03 ")
                        query.AppendLine("         ,hari_04 ")
                        query.AppendLine("         ,hari_05 ")
                        query.AppendLine("         ,hari_06 ")
                        query.AppendLine("         ,hari_07 ")
                        query.AppendLine("         ,hari_08 ")
                        query.AppendLine("         ,hari_09 ")
                        query.AppendLine("         ,hari_10 ")
                        query.AppendLine("         ,hari_11 ")
                        query.AppendLine("         ,hari_12 ")
                        query.AppendLine("         ,hari_13 ")
                        query.AppendLine("         ,hari_14 ")
                        query.AppendLine("         ,hari_15 ")
                        query.AppendLine("         ,hari_16 ")
                        query.AppendLine("         ,hari_17 ")
                        query.AppendLine("         ,hari_18 ")
                        query.AppendLine("         ,hari_19 ")
                        query.AppendLine("         ,hari_20 ")
                        query.AppendLine("         ,hari_21 ")
                        query.AppendLine("         ,hari_22 ")
                        query.AppendLine("         ,hari_23 ")
                        query.AppendLine("         ,hari_24 ")
                        query.AppendLine("         ,hari_25 ")
                        query.AppendLine("         ,hari_26 ")
                        query.AppendLine("         ,hari_27 ")
                        query.AppendLine("         ,hari_28 ")
                        query.AppendLine("         ,hari_29 ")
                        query.AppendLine("         ,hari_30 ")
                        query.AppendLine("         ,hari_31 ")
                        query.AppendLine("  From ")
                        query.AppendLine("      ad_dis_rtjn_sum_qty_opmj_tipe ")
                        query.AppendLine(" where ")
                        query.AppendLine("     dmc_code='" & DMC_CODE & "' ")
                        query.AppendLine("     and shift_date='" & str_shiftdate & "' ")

                        Dim GetTblManpower As DataTable = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)

                        query.Length = 0
                        query.Capacity = 0

                        If GetTblManpower.Rows.Count > 0 Then
                            query.AppendLine(" update ")
                            query.AppendLine("     ad_dis_rtjn_sum_qty_opmj_tipe ")
                            query.AppendLine(" set ")
                            query.AppendLine("     dmc_code='" & EscapeQuote(DMC_CODE) & "' ")
                            query.AppendLine("     ,hari_01='" & HARI_01 & "' ")
                            query.AppendLine("     ,hari_02='" & HARI_02 & "' ")
                            query.AppendLine("     ,hari_03='" & HARI_03 & "' ")
                            query.AppendLine("     ,hari_04='" & HARI_04 & "' ")
                            query.AppendLine("     ,hari_05='" & HARI_05 & "' ")
                            query.AppendLine("     ,hari_06='" & HARI_06 & "' ")
                            query.AppendLine("     ,hari_07='" & HARI_07 & "' ")
                            query.AppendLine("     ,hari_08='" & HARI_08 & "' ")
                            query.AppendLine("     ,hari_09='" & HARI_09 & "' ")
                            query.AppendLine("     ,hari_10='" & HARI_10 & "' ")
                            query.AppendLine("     ,hari_11='" & HARI_11 & "' ")
                            query.AppendLine("     ,hari_12='" & HARI_12 & "' ")
                            query.AppendLine("     ,hari_13='" & HARI_13 & "' ")
                            query.AppendLine("     ,hari_14='" & HARI_14 & "' ")
                            query.AppendLine("     ,hari_15='" & HARI_15 & "' ")
                            query.AppendLine("     ,hari_16='" & HARI_16 & "' ")
                            query.AppendLine("     ,hari_17='" & HARI_17 & "' ")
                            query.AppendLine("     ,hari_18='" & HARI_18 & "' ")
                            query.AppendLine("     ,hari_19='" & HARI_19 & "' ")
                            query.AppendLine("     ,hari_20='" & HARI_20 & "' ")
                            query.AppendLine("     ,hari_21='" & HARI_21 & "' ")
                            query.AppendLine("     ,hari_22='" & HARI_22 & "' ")
                            query.AppendLine("     ,hari_23='" & HARI_23 & "' ")
                            query.AppendLine("     ,hari_24='" & HARI_24 & "' ")
                            query.AppendLine("     ,hari_25='" & HARI_25 & "' ")
                            query.AppendLine("     ,hari_26='" & HARI_26 & "' ")
                            query.AppendLine("     ,hari_27='" & HARI_27 & "' ")
                            query.AppendLine("     ,hari_28='" & HARI_28 & "' ")
                            query.AppendLine("     ,hari_29='" & HARI_29 & "' ")
                            query.AppendLine("     ,hari_30='" & HARI_30 & "' ")
                            query.AppendLine("     ,hari_31='" & HARI_31 & "' ")
                            query.AppendLine(" where ")
                            query.AppendLine("     dmc_code='" & DMC_CODE & "' ")
                            query.AppendLine("     and shift_date='" & str_shiftdate & "' ")

                            ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)

                            query.Length = 0
                            query.Capacity = 0

                            Console.WriteLine("Proses Update Data OPMJ By Tipe : " & EscapeQuote(DMC_CODE) & "")
                        Else
                            query.AppendLine(" Insert Into ")
                            query.AppendLine("     ad_dis_rtjn_sum_qty_opmj_tipe ")
                            query.AppendLine("     ( ")
                            query.AppendLine("         shift_date")
                            query.AppendLine("         ,dmc_code ")
                            query.AppendLine("         ,hari_01 ")
                            query.AppendLine("         ,hari_02 ")
                            query.AppendLine("         ,hari_03 ")
                            query.AppendLine("         ,hari_04 ")
                            query.AppendLine("         ,hari_05 ")
                            query.AppendLine("         ,hari_06 ")
                            query.AppendLine("         ,hari_07 ")
                            query.AppendLine("         ,hari_08 ")
                            query.AppendLine("         ,hari_09 ")
                            query.AppendLine("         ,hari_10 ")
                            query.AppendLine("         ,hari_11 ")
                            query.AppendLine("         ,hari_12 ")
                            query.AppendLine("         ,hari_13 ")
                            query.AppendLine("         ,hari_14 ")
                            query.AppendLine("         ,hari_15 ")
                            query.AppendLine("         ,hari_16 ")
                            query.AppendLine("         ,hari_17 ")
                            query.AppendLine("         ,hari_18 ")
                            query.AppendLine("         ,hari_19 ")
                            query.AppendLine("         ,hari_20 ")
                            query.AppendLine("         ,hari_21 ")
                            query.AppendLine("         ,hari_22 ")
                            query.AppendLine("         ,hari_23 ")
                            query.AppendLine("         ,hari_24 ")
                            query.AppendLine("         ,hari_25 ")
                            query.AppendLine("         ,hari_26 ")
                            query.AppendLine("         ,hari_27 ")
                            query.AppendLine("         ,hari_28 ")
                            query.AppendLine("         ,hari_29 ")
                            query.AppendLine("         ,hari_30 ")
                            query.AppendLine("         ,hari_31 ")
                            query.AppendLine("     ) ")
                            query.AppendLine(" values ")
                            query.AppendLine("     ( ")
                            query.AppendLine("     '" & str_shiftdate & "'")
                            query.AppendLine("     ,'" & EscapeQuote(DMC_CODE) & "' ")
                            query.AppendLine("     ,'" & HARI_01 & "' ")
                            query.AppendLine("     ,'" & HARI_02 & "' ")
                            query.AppendLine("     ,'" & HARI_03 & "' ")
                            query.AppendLine("     ,'" & HARI_04 & "' ")
                            query.AppendLine("     ,'" & HARI_05 & "' ")
                            query.AppendLine("     ,'" & HARI_06 & "' ")
                            query.AppendLine("     ,'" & HARI_07 & "' ")
                            query.AppendLine("     ,'" & HARI_08 & "' ")
                            query.AppendLine("     ,'" & HARI_09 & "' ")
                            query.AppendLine("     ,'" & HARI_10 & "' ")
                            query.AppendLine("     ,'" & HARI_11 & "' ")
                            query.AppendLine("     ,'" & HARI_12 & "' ")
                            query.AppendLine("     ,'" & HARI_13 & "' ")
                            query.AppendLine("     ,'" & HARI_14 & "' ")
                            query.AppendLine("     ,'" & HARI_15 & "' ")
                            query.AppendLine("     ,'" & HARI_16 & "' ")
                            query.AppendLine("     ,'" & HARI_17 & "' ")
                            query.AppendLine("     ,'" & HARI_18 & "' ")
                            query.AppendLine("     ,'" & HARI_19 & "' ")
                            query.AppendLine("     ,'" & HARI_20 & "' ")
                            query.AppendLine("     ,'" & HARI_21 & "' ")
                            query.AppendLine("     ,'" & HARI_22 & "' ")
                            query.AppendLine("     ,'" & HARI_23 & "' ")
                            query.AppendLine("     ,'" & HARI_24 & "' ")
                            query.AppendLine("     ,'" & HARI_25 & "' ")
                            query.AppendLine("     ,'" & HARI_26 & "' ")
                            query.AppendLine("     ,'" & HARI_27 & "' ")
                            query.AppendLine("     ,'" & HARI_28 & "' ")
                            query.AppendLine("     ,'" & HARI_29 & "' ")
                            query.AppendLine("     ,'" & HARI_30 & "' ")
                            query.AppendLine("     ,'" & HARI_31 & "' ")
                            query.AppendLine("     ) ")
                            ClsConfig.ExecuteNonQuery(query.ToString(), ClsConfig.IPServer_ADDONS)
                            query.Length = 0
                            query.Capacity = 0
                            Console.WriteLine("Proses Insert Data OPMJ By Tipe : " & EscapeQuote(DMC_CODE) & " ")
                        End If
                    End If
                Next

                day = day.AddDays(1)

            End While
        End If

        Catch ex As Exception
            ClsConfig.create_log_error("[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Calculation OPMJ By Tipe Error")
            Environment.Exit(0)
        End Try

        query.AppendLine(" select ")
        query.AppendLine("     id ")
        query.AppendLine("     ,shift_date ")
        query.AppendLine("     ,dmc_code ")
        query.AppendLine("     ,hari_01 ")
        query.AppendLine("     ,hari_02 ")
        query.AppendLine("     ,hari_03 ")
        query.AppendLine("     ,hari_04 ")
        query.AppendLine("     ,hari_05 ")
        query.AppendLine("     ,hari_06 ")
        query.AppendLine("     ,hari_07 ")
        query.AppendLine("     ,hari_08 ")
        query.AppendLine("     ,hari_09 ")
        query.AppendLine("     ,hari_10 ")
        query.AppendLine("     ,hari_11 ")
        query.AppendLine("     ,hari_12 ")
        query.AppendLine("     ,hari_13 ")
        query.AppendLine("     ,hari_14 ")
        query.AppendLine("     ,hari_15 ")
        query.AppendLine("     ,hari_16 ")
        query.AppendLine("     ,hari_17 ")
        query.AppendLine("     ,hari_18 ")
        query.AppendLine("     ,hari_19 ")
        query.AppendLine("     ,hari_20 ")
        query.AppendLine("     ,hari_21 ")
        query.AppendLine("     ,hari_22 ")
        query.AppendLine("     ,hari_23 ")
        query.AppendLine("     ,hari_24 ")
        query.AppendLine("     ,hari_25 ")
        query.AppendLine("     ,hari_26 ")
        query.AppendLine("     ,hari_27 ")
        query.AppendLine("     ,hari_28 ")
        query.AppendLine("     ,hari_29 ")
        query.AppendLine("     ,hari_30 ")
        query.AppendLine("     ,hari_31 ")
        query.AppendLine(" from ")
        query.AppendLine("     ad_dis_rtjn_sum_qty_opmj_tipe ")
        query.AppendLine(" where ")
        query.AppendLine("     shift_date between '" + startDate.ToString("yyyyMM") + "' and '" + endDate.ToString("yyyyMM") + "' ")

        CalculateType = ClsConfig.ExecuteQuery(query.ToString(), ClsConfig.IPServer_ADDONS)

    End Function

End Class
