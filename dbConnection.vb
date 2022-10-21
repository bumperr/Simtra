Imports System.Data.OleDb
Imports System.IO

Module dbConnection
    'PUBLIC FINAL booking info to be inserted to database 
    Public finalTransportID, finalSeatID, finalRoomID As String
    Public finalCheckIn, finalCheckout As Date
    Public finalRoomTotalPrice As Double
    'set up global variable to be used in all project
    Public MyPath As String = CurDir()
    Public adminPath As String = MyPath & "\admin.txt"
    Public dbPath As String = MyPath & "\Database\Simtra.mdb"
    Public adminName As String
    Public dbConnectionString As String = "Provider=Microsoft.Jet.OLEDB.4.0; Data Source=" & dbPath
    Public dbConnect As OleDbConnection = New OleDbConnection()
    Public session_email As String
    Public admin_viewOrder_type As String
    Public session_name As String
    Public session_userType As String
    Public session_transportID, session_AccommodationID As String
    Public session_accommodation As String
    Public session_typeofservice As String
    Public session_totalPayment As Double
    Public session_transport_endDate As Date
    Public travel_endDate As Date
    Public session_seat_ID As Integer

    Public session_picture As String = MyPath & "\Database\noPic.jpg"

    Public outerState_items As String() = {"Indonesia", "Malaysia", "Singapore", "Thailand"}

    Public interState_items As String() = {"Johor", "Kedah", "Kelantan", "Melaka", "Negeri Sembilan",
                                           "Pahang", "Penang", "Perak", "Perlis", "Sabah", "Sarawak",
                                           "Selangor", "Terengganu", "Kuala Lumpur", "Labuan", "Putrajaya"}

    Public customerImage As String

    'custom function to create database connection using ole db framework
    Public Function dbOpenConnection() As Boolean
        Try
            dbConnect.ConnectionString = dbConnectionString
            dbConnect.Open()
            Return True
        Catch ex As Exception
            Return False
        End Try
    End Function

    Sub dbConnectionTest()
        If dbOpenConnection() Then
            dbConnect.Close()
        Else
            MsgBox("connection to database error")
        End If
    End Sub

    Sub ending()
        'clear temporary files
        Dim strDirectory As String = MyPath & "\temp"
        For Each filepath As String In Directory.GetFiles(strDirectory)
            File.Delete(filepath)
        Next
        End
    End Sub
End Module

Module CustomFunction

    Public Sub customMessage(ByVal type As Boolean, ByVal message As String)
        Dim messagePrompt As New customMessageBox
        messagePrompt.lblMessage.Text = message

        If type Then
            messagePrompt.picTypeOfNoti.Image = SimtraUser.My.Resources.Resources.blueSuccessGIF

        Else
            messagePrompt.picTypeOfNoti.Image = SimtraUser.My.Resources.Resources.errorGIF1


            'default image
        End If

        messagePrompt.ShowDialog()
    End Sub
    'function that return path  or the file Name of the uploaded document 
    Public Function uploadFile(ByVal originalName As String) As String
        Dim file As New OpenFileDialog
        file.Filter = "PDF files|*.pdf|All files|*.*;"

        If file.ShowDialog() = DialogResult.OK Then
            Return file.FileName()
        Else
            Return originalName
        End If

    End Function
    'wrapper function to open file 
    Public Sub viewFile(ByVal path As String)

        If path = Nothing Then
            customMessage(False, "The user does not upload the File")
        Else
            System.Diagnostics.Process.Start(path)

        End If

    End Sub

    Public Function uploadPicture() As String
        Dim photo As New OpenFileDialog
        photo.Filter = ("Picture File|*.jpg;*.gif;*.png;*.bmp;*.jpeg")
        photo.ShowDialog()
        Return photo.FileName()
    End Function

    Public Function updateFileToDB(ByVal tableName As String, ByVal dataColumn As String, ByVal keycolumn As String, ByVal keyValue As String, ByVal data As Byte()) As Boolean

        If dbOpenConnection() Then
            Dim sqlStatement = "UPDATE " & tableName & " SET " & dataColumn & "=@data " & "WHERE " & keycolumn & "=@keyValue"

            Dim cmd As New OleDbCommand(sqlStatement, dbConnect)
            cmd.Parameters.AddWithValue("@data", data)
            cmd.Parameters.AddWithValue("@keyValue", keyValue)
            cmd.ExecuteNonQuery()
            dbConnect.Close()
            Return True
        Else
            Return False
        End If
    End Function

    Public Function retrieveFileFromDB(ByVal tableName As String, ByVal targetedColumn As String, ByVal keyColumn As String, ByVal keyValue As String, ByVal fileName As String) As String
        If dbOpenConnection() Then
            Dim sqlStatement = "SELECT " & targetedColumn & " FROM " & tableName & " WHERE " & keyColumn & "=@keyValue"

            Dim cmd As New OleDbCommand(sqlStatement, dbConnect)
            cmd.Parameters.AddWithValue("@keyValue", keyValue)
            Try
                Dim fileData As Byte() = CType(cmd.ExecuteScalar(), Byte())
                Dim filePath As String = MyPath & "\temp\" & fileName
                File.WriteAllBytes(filePath, fileData)
                cmd.Dispose()
                dbConnect.Close()
                Return filePath
            Catch ex As Exception
                cmd.Dispose()
                dbConnect.Close()
                Return Nothing
            End Try





        Else
            Return MsgBox("error")
        End If
    End Function

    'function that validate formatting of email
    Public Function validateEmail(ByVal input As String) As Boolean

        If input.Length = 0 OrElse input.IndexOf("@") = -1 OrElse input.IndexOf("@") > input.IndexOf(".") Then
            Return True
        End If


        Return False

    End Function


    'function that validate password is same as re-confirm field
    Public Function validatePassword(newPass As String, confirmPass As String)
        If newPass.Length = 0 Or confirmPass.Length = 0 Then
            customMessage(False, "Password is empty")
            Return True
        End If

        If newPass = confirmPass Then
            Return False
        Else
            customMessage(False, "Password does not match")
            Return True
        End If

    End Function

    'custome function to validatePhone Number


    Public Function validatePhone(telno As String)
        If telno.Length = 0 Or Not IsNumeric(telno) Then
            customMessage(False, "Phone number in wrong format")
            Return True
        End If
        Return False
    End Function


End Module
