Imports Microsoft.VisualBasic
Imports System.Data.SqlClient
Imports System.Text
Imports System.Runtime.InteropServices
Imports System.IO

Public Class ClsConfig
    'execute query
    Public Shared SQL As String
    Public Shared Cn As SqlConnection
    Public Shared Cmd As SqlCommand
    Public Shared Da As SqlDataAdapter
    Public Shared Ds As DataSet
    Public Shared Dt As DataTable

    'server database
    Public Shared DATABASE_TYPE As String
    Public Shared IPServer_RTJN_PRD As String
    Public Shared IPServer_TxDTIPRD As String
    Public Shared IPServer_ADDONS As String
    Public Shared IPServer_ADDONS_Intranet As String

    'email
    Public Shared email_from_alias As String
    Public Shared email_nama As String
    Public Shared email_password As String
    Public Shared email_server_smtp As String
    Public Shared email_server_port As String
    Public Shared subject_email As String
    Public Shared tls As String

    'log history error
    Public Shared nama_folder_log_error As String
    Public Shared nama_file_txt_log_error As String

    'export excel
    Public Shared nama_file_template As String
    Public Shared nama_file_lampiran_email As String
    Public Shared lokasi_simpan_file As String

    'monitoring OPMJ
    Public Shared nama_file_template_monitoring As String
    Public Shared nama_file_lampiran_email_monitoring As String
    Public Shared lokasi_simpan_file_monitoring As String
    Public Shared subject_email_monitoring As String
    Public Shared email_monitoring_mail_sender As String

    <DllImport("kernel32.dll")>
    Private Shared Function GetPrivateProfileString(ByVal lpApplicationName As String,
                                                    ByVal lpKeyName As String,
                                                    ByVal lpDefault As String,
                                                    ByVal lpReturnedString As StringBuilder,
                                                    ByVal nSize As UInt32,
                                                    ByVal lpFileName As String) As UInt32
    End Function

    Private Shared Function GetIniString(ByVal iniFileName As String,
                                 ByVal section As String,
                                 ByVal key As String,
                                 Optional ByVal defaultValue As String = "") As String
        Dim nSize As Integer = 1024
        Dim sb As StringBuilder = New StringBuilder(nSize)
        Dim ret As UInt32 = GetPrivateProfileString(section, key, defaultValue, sb, Convert.ToUInt32(sb.Capacity), iniFileName)

        Return sb.ToString
    End Function

    Public Shared Sub get_variable_setting()
        Try

            Dim EXE_PATH As String

            'server database
            EXE_PATH = System.AppDomain.CurrentDomain.BaseDirectory
            DATABASE_TYPE = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "DATABASE", "TYPE")
            IPServer_RTJN_PRD = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "DATABASE", "RTJN")
            IPServer_TxDTIPRD = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "DATABASE", "TPICS")
            IPServer_ADDONS = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "DATABASE", "ADDONS")
            Console.WriteLine("---> Setup database done")

            'email
            email_from_alias = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "EMAIL", "email_from_alias")
            email_nama = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "EMAIL", "email_nama")
            email_password = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "EMAIL", "email_password")
            email_server_smtp = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "EMAIL", "email_server_smtp")
            email_server_port = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "EMAIL", "email_server_port")
            subject_email = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "EMAIL", "subject_email")
            tls = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "EMAIL", "tls")
            Console.WriteLine("---> Setup email done")

            'log history error
            nama_folder_log_error = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "FILE", "nama_folder_log_error")
            nama_file_txt_log_error = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "FILE", "nama_file_txt_log_error")
            Console.WriteLine("---> Setup log error history done")

            'export excel
            nama_file_template = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "FILE", "nama_file_template")
            nama_file_lampiran_email = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "FILE", "nama_file_lampiran_email")
            lokasi_simpan_file = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "FILE", "lokasi_simpan_file")
            Console.WriteLine("---> Setup template file excel done")

            'monitoring OPMJ
            nama_file_template_monitoring = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "MONITORING", "nama_file_template_monitoring")
            nama_file_lampiran_email_monitoring = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "MONITORING", "nama_file_lampiran_email_monitoring")
            lokasi_simpan_file_monitoring = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "MONITORING", "lokasi_simpan_file_monitoring")
            subject_email_monitoring = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "MONITORING", "subject_email_monitoring")
            email_monitoring_mail_sender = GetIniString(EXE_PATH & "\Auto_RTJN_Opmj.ini", "MONITORING", "email_monitoring_mail_sender")
            Console.WriteLine("---> Setup monitoring OPMJ done")

        Catch ex As Exception
            ClsConfig.create_log_error("[" + DateTime.Now.ToString("dd-MM-yyyy hh:mm:ss") + "] -- [ " + ex.Message + " ] -- Setup config file .ini error")
            Environment.Exit(0)
        End Try


    End Sub

    Public Shared Function OpenConn(ByVal IPServer As String) As Boolean
        Cn = New SqlConnection(IPServer)
        Cn.Open()

        If Cn.State <> ConnectionState.Open Then
            Return False
        Else
            Return True
        End If
    End Function

    Public Shared Sub CloseConn()
        If Not IsNothing(Cn) Then
            Cn.Close()
            Cn = Nothing
        End If
    End Sub

    Public Shared Function ExecuteQuery(ByVal Query As String, ByVal IPServer As String) As DataTable
        If Not OpenConn(IPServer) Then
            MsgBox("Koneksi Gagal..!!", MsgBoxStyle.Critical, "Access Failed")
            Return Nothing
            Exit Function
        End If

        Cmd = New SqlCommand(Query, Cn)
        Da = New SqlDataAdapter
        Da.SelectCommand = Cmd

        Ds = New Data.DataSet
        Cmd.CommandTimeout = 1000
        Da.Fill(Ds)
        Dt = Ds.Tables(0)

        Ds = Nothing
        Da = Nothing
        Cmd = Nothing

        CloseConn()

        Return Dt

        Dt = Nothing
    End Function

    Public Shared Sub ExecuteNonQuery(ByVal Query As String, ByVal IPServer As String)
        If Not OpenConn(IPServer) Then
            MsgBox("Koneksi Gagal..!!", MsgBoxStyle.Critical, "Access Failed..!!")
            Exit Sub
        End If

        Cmd = New SqlCommand
        Cmd.Connection = Cn
        Cmd.CommandTimeout = 600
        Cmd.CommandType = CommandType.Text
        Cmd.CommandText = Query
        Cmd.ExecuteNonQuery()
        Cmd = Nothing
        CloseConn()
    End Sub

    Public Shared Sub create_log_error(ByVal pesan_error As String)
        Dim PathFile As String = ClsConfig.nama_folder_log_error
        If Not System.IO.Directory.Exists(PathFile) Then
            System.IO.Directory.CreateDirectory(PathFile)
        End If

        Dim nama_file_txt_log_error_n_path As String
        nama_file_txt_log_error_n_path = PathFile & "\" & ClsConfig.nama_file_txt_log_error & ".txt"

        If Not File.Exists(nama_file_txt_log_error_n_path) Then
            Using writer As New StreamWriter(nama_file_txt_log_error_n_path, True)
                writer.Write(pesan_error)
            End Using
        Else
            File.AppendAllText(nama_file_txt_log_error_n_path, Environment.NewLine + pesan_error)
        End If
    End Sub

End Class
