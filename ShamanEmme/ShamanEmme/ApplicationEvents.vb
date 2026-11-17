Imports ShamanExpressDLL
Imports System.Data
Imports System.Data.SqlClient

Namespace My

    ' The following events are available for MyApplication:
    ' 
    ' Startup: Raised when the application starts, before the startup form is created.
    ' Shutdown: Raised after all application forms are closed.  This event is not raised if the application terminates abnormally.
    ' UnhandledException: Raised if the application encounters an unhandled exception.
    ' StartupNextInstance: Raised when launching a single-instance application and the application is already active. 
    ' NetworkAvailabilityChanged: Raised when the network connection is connected or disconnected.
    Partial Friend Class MyApplication

        Private Sub MyApplication_Shutdown(sender As Object, e As EventArgs) Handles Me.Shutdown
            Try
                Dim objSession As New conUsuariosLogueados
                Dim cnnCloseNET As SqlConnection

                If cnnsNET.Count > 0 Then

                    objSession.Eliminar(logID)

                    '----> Desconecto toda las conexiones
                    For Each cnnCloseNET In cnnsNET.Values
                        cnnCloseNET.Close()
                    Next

                End If

                objSession = Nothing

                End

            Catch ex As Exception
                HandleError(Me.GetType.Name, "MyApplication_Shutdown", ex)
            End Try

        End Sub

        Private Sub MyApplication_Startup(sender As Object, e As ApplicationServices.StartupEventArgs) Handles Me.Startup

            Try

                If shamanStartUp.GetValoresHardkey() Then

                    If shamanStartUp.GetVariablesConexion() Then

                        If shamanStartUp.AbrirConexion(cnnDefault) Then

                            InitDateVars()
                            InitSepDecimal()

                            shamanConfig = New conConfiguracion

                            shamanConfig.UpConfig()

                        Else

                            End

                        End If

                    Else

                        '-----> corto por licencia vencida en frmLog

                    End If

                Else

                    End

                End If

            Catch ex As Exception
                HandleError(Me.GetType.Name, "MyApplication_Startup", ex)
            End Try

        End Sub

    End Class


End Namespace

