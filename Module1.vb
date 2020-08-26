Imports System.IO
Imports System.Security.Cryptography

Module Module1

    Public Sub Main() 'ByVal Method As String, Optional ByVal ConfigureFile As String = "", Optional ByVal ExtensionPatterns As String = "")
        Dim nameServerList As New ArrayList
        Dim Method As String = "Modified"
        Dim ExtensionPatterns As String = "*.*"

        nameServerList.Add("\\svr-tdc-ss01.cpf.co.th\smartsoft_app$")
        nameServerList.Add("\\svr-tdc-ss02.cpf.co.th\smartsoft_app$")
        nameServerList.Add("\\svr-tdc-ss03.cpf.co.th\smartsoft_app$")

        Dim sourceDS As New DataSet
        Dim sourceTable As New DataTable
        Dim sourceRow As DataRow
        sourceDS.Tables.Add(sourceTable)
        sourceDS.Tables(0).Columns.Add("Name", GetType(String))
        sourceDS.Tables(0).Columns.Add("DirectoryName", GetType(String))
        sourceDS.Tables(0).Columns.Add("FullName", GetType(String))
        sourceDS.Tables(0).Columns.Add("Extension", GetType(String))
        sourceDS.Tables(0).Columns.Add("CreationTime", GetType(DateTime))
        sourceDS.Tables(0).Columns.Add("LastAccessTime", GetType(DateTime))
        sourceDS.Tables(0).Columns.Add("LastWriteTime", GetType(DateTime))
        sourceDS.Tables(0).Columns.Add("Status", GetType(String))

        Dim destinationDS As New DataSet
        Dim destinationTable As New DataTable
        Dim destinationRow As DataRow
        destinationDS.Tables.Add(destinationTable)
        destinationDS.Tables(0).Columns.Add("Name", GetType(String))
        destinationDS.Tables(0).Columns.Add("DirectoryName", GetType(String))
        destinationDS.Tables(0).Columns.Add("FullName", GetType(String))
        destinationDS.Tables(0).Columns.Add("Extension", GetType(String))
        destinationDS.Tables(0).Columns.Add("CreationTime", GetType(DateTime))
        destinationDS.Tables(0).Columns.Add("LastAccessTime", GetType(DateTime))
        destinationDS.Tables(0).Columns.Add("LastWriteTime", GetType(DateTime))
        destinationDS.Tables(0).Columns.Add("Status", GetType(String))

        Dim resultDS As New DataSet
        Dim resultTable As New DataTable
        Dim resultRow As DataRow
        resultDS.Tables.Add(resultTable)
        resultDS.Tables(0).Columns.Add("Name", GetType(String))
        resultDS.Tables(0).Columns.Add("ShowDirectoryName", GetType(String))
        For createColumn As Integer = 0 To nameServerList.Count - 1
            resultDS.Tables(0).Columns.Add(nameServerList(createColumn), GetType(String))
        Next

        If nameServerList.Count > 0 Then
            Dim di1 As New IO.DirectoryInfo(nameServerList(0))

            Dim files As New List(Of FileInfo)
            Dim searchPatterns As String() = Replace(ExtensionPatterns, " ", "").Split(",")
            For Each searchPattern As String In searchPatterns
                files.AddRange(di1.GetFiles(searchPattern, SearchOption.AllDirectories))
            Next

            Dim diar1 As IO.FileInfo() = files.ToArray
            Dim dra1 As IO.FileInfo

            'list the names of all files in the specified directory
            For Each dra1 In diar1
                sourceRow = sourceDS.Tables(0).NewRow
                sourceRow("Name") = dra1.Name
                sourceRow("DirectoryName") = dra1.DirectoryName
                sourceRow("FullName") = dra1.FullName
                sourceRow("Extension") = dra1.Extension
                sourceRow("CreationTime") = CDate(dra1.CreationTime)
                sourceRow("LastAccessTime") = CDate(dra1.LastAccessTime)
                sourceRow("LastWriteTime") = CDate(dra1.LastWriteTime)
                sourceDS.Tables(0).Rows.Add(sourceRow)
            Next

            For Each onlyOnSource As DataRow In sourceDS.Tables(0).Select("Status Is null or Status = ''")
                resultRow = resultDS.Tables(0).NewRow
                resultRow("Name") = Replace(onlyOnSource("FullName"), nameServerList(0), "")
                resultRow("ShowDirectoryName") = Replace(onlyOnSource("DirectoryName"), nameServerList(0), "")
                resultRow(nameServerList(0)) = "Found"
                resultDS.Tables(0).Rows.Add(resultRow)
            Next

            Dim sourceFile, destinationFile As String
            Dim desResultRow As DataRow

            Dim di2 As IO.DirectoryInfo
            Dim diar2 As IO.FileInfo()
            Dim dra2 As IO.FileInfo


            For i As Integer = 1 To nameServerList.Count - 1
                di2 = New IO.DirectoryInfo(nameServerList(i))
                files = New List(Of FileInfo)

                For Each searchPattern As String In searchPatterns
                    files.AddRange(di2.GetFiles(searchPattern, SearchOption.AllDirectories))
                Next
                diar2 = files.ToArray
                destinationTable.Rows.Clear()

                'list the names of all files in the specified directory
                For Each dra2 In diar2
                    destinationRow = destinationDS.Tables(0).NewRow
                    destinationRow("Name") = dra2.Name
                    destinationRow("DirectoryName") = dra2.DirectoryName
                    destinationRow("FullName") = dra2.FullName
                    destinationRow("Extension") = dra2.Extension
                    destinationRow("CreationTime") = CDate(dra2.CreationTime)
                    destinationRow("LastAccessTime") = CDate(dra2.LastAccessTime)
                    destinationRow("LastWriteTime") = CDate(dra2.LastWriteTime)
                    destinationDS.Tables(0).Rows.Add(destinationRow)
                Next

                Dim compareResult As Boolean
                For Each result As DataRow In sourceDS.Tables(0).Rows
                    sourceFile = result("FullName")
                    Try
                        desResultRow = destinationDS.Tables(0).Select("FullName = '" & Replace(result("DirectoryName"), nameServerList(0), nameServerList(i)) & "\" & result("Name") & "'")(0)
                    Catch
                        desResultRow = Nothing
                    End Try
                    If desResultRow IsNot Nothing Then
                        destinationFile = desResultRow("FullName")

                        resultRow = resultDS.Tables(0).Select("Name = '" & Replace(result("FullName"), nameServerList(0), "") & "'")(0)
                        If resultRow IsNot Nothing Then
                            If Method = "MD5" Then
                                compareResult = FileCompareWithMD5(sourceFile, destinationFile)
                            ElseIf Method = "Binary" Then
                                compareResult = FileCompare(sourceFile, destinationFile)
                            ElseIf Method = "Modified" Then
                                compareResult = result("LastWriteTime") = desResultRow("LastWriteTime")
                            End If

                            If compareResult Then
                                resultRow(nameServerList(i)) = "Match"
                            Else
                                resultRow(nameServerList(i)) = "Unmatch"
                            End If
                        Else
                            resultRow = resultDS.Tables(0).NewRow
                            resultRow("Name") = Replace(result("FullName"), nameServerList(i), "")
                            resultRow("ShowDirectoryName") = Replace(result("DirectoryName"), nameServerList(i), "")
                            resultRow(nameServerList(i)) = ""
                            resultDS.Tables(0).Rows.Add(resultRow)
                        End If
                        destinationTable.Rows.Remove(desResultRow)
                    Else
                        resultRow = resultDS.Tables(0).Select("Name = '" & Replace(result("FullName"), nameServerList(0), "") & "'")(0)
                        resultRow(nameServerList(i)) = ""
                    End If
                Next

                Dim existingRow As DataRow

                For Each onlyOnDestination As DataRow In destinationDS.Tables(0).Rows
                    Try
                        existingRow = resultDS.Tables(0).Select("Name = '" & Replace(onlyOnDestination("FullName"), nameServerList(i), "") & "'")(0)
                    Catch
                        existingRow = Nothing
                    End Try
                    If existingRow Is Nothing Then
                        existingRow = resultDS.Tables(0).NewRow

                        For clearColumn As Integer = 0 To nameServerList.Count - 1
                            existingRow(nameServerList(clearColumn)) = ""
                        Next

                        existingRow("Name") = Replace(onlyOnDestination("FullName"), nameServerList(i), "")
                        existingRow("ShowDirectoryName") = Replace(onlyOnDestination("DirectoryName"), nameServerList(i), "")
                        existingRow(nameServerList(i)) = "Found"
                        resultDS.Tables(0).Rows.Add(existingRow)
                    Else
                        existingRow(nameServerList(i)) = "Found"
                    End If
                Next
            Next

            Dim view As New DataView(resultDS.Tables(0))
            view.Sort = "ShowDirectoryName"

            resultDS.WriteXml("c:\\Result.xml")
        End If

    End Sub

    Private Function FileCompareWithMD5(File1 As String, File2 As String) As Boolean
        Dim objMD5 As New MD5CryptoServiceProvider()
        Dim objEncoding As New System.Text.ASCIIEncoding()

        Dim aFile1() As Byte, aFile2() As Byte
        Dim strContents1, strContents2 As String
        Dim objReader As StreamReader
        Dim objFS As FileStream
        Dim bAns As Boolean
        If Not File.Exists(File1) Then
            Return False
        End If
        If Not File.Exists(File2) Then
            Return False
        End If

        Try
            objFS = New FileStream(File1, FileMode.Open, FileAccess.Read)
            objReader = New StreamReader(objFS)
            aFile1 = objEncoding.GetBytes(objReader.ReadToEnd)
            strContents1 = objEncoding.GetString(objMD5.ComputeHash(aFile1))
            objReader.Close()
            objFS.Close()


            objFS = New FileStream(File2, FileMode.Open, FileAccess.Read)
            objReader = New StreamReader(objFS)
            aFile2 = objEncoding.GetBytes(objReader.ReadToEnd)
            strContents2 = objEncoding.GetString(objMD5.ComputeHash(aFile2))

            bAns = strContents1 = strContents2
            objReader.Close()
            objFS.Close()
            aFile1 = Nothing
            aFile2 = Nothing
            Return bAns
        Catch ex As Exception
            Return False
        End Try
    End Function

    Private Function FileCompare(ByVal file1 As String, ByVal file2 As String) As Boolean
        Dim file1byte As Integer
        Dim file2byte As Integer
        Dim fs1 As FileStream
        Dim fs2 As FileStream

        ' Determine if the same file was referenced two times.
        If (file1 = file2) Then
            ' Return 0 to indicate that the files are the same.
            Return True
        End If

        ' Open the two files.
        fs1 = New FileStream(file1, FileMode.Open, FileAccess.Read)
        fs2 = New FileStream(file2, FileMode.Open, FileAccess.Read)

        ' Check the file sizes. If they are not the same, the files
        ' are not equal.
        If (fs1.Length <> fs2.Length) Then
            ' Close the file
            fs1.Close()
            fs2.Close()

            ' Return a non-zero value to indicate that the files are different.
            Return False
        End If

        ' Read and compare a byte from each file until either a
        ' non-matching set of bytes is found or until the end of
        ' file1 is reached.
        Do
            ' Read one byte from each file.
            file1byte = fs1.ReadByte()
            file2byte = fs2.ReadByte()
        Loop While ((file1byte = file2byte) And (file1byte <> -1))

        ' Close the files.
        fs1.Close()
        fs2.Close()

        ' Return the success of the comparison. "file1byte" is
        ' equal to "file2byte" at this point only if the files are 
        ' the same.


        Return ((file1byte - file2byte) = 0)
    End Function

End Module
