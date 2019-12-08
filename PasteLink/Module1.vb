Imports System.Windows.Forms
Module Module1
    Public sfile, sdir, efile, edir As Integer
    Public err As Boolean = False
    Public LinkDIR As Boolean = True
    Public Class ConsoleOutput
        Private Shared gWorkingDirectory As String = Environment.GetFolderPath(Environment.SpecialFolder.Personal)
        Public Shared Property WorkingDirectory() As String
            Get
                Return gWorkingDirectory
            End Get
            Set(ByVal Value As String)
                gWorkingDirectory = Value
            End Set
        End Property
        Public Shared Function ExcuteCmd(ByVal command As String) As String
            Dim mResult As String = ""
            Dim tmpProcess As New Process
            With tmpProcess
                With .StartInfo
                    .CreateNoWindow = True
                    .FileName = .EnvironmentVariables.Item("ComSpec")
                    .RedirectStandardOutput = True
                    .UseShellExecute = False
                    .Arguments = String.Format("/C {0}", command)
                    .WorkingDirectory = gWorkingDirectory
                End With
                Try
                    .Start()
                    .WaitForExit(5000)
                    mResult = .StandardOutput.ReadToEnd
                Catch e As System.ComponentModel.Win32Exception
                    mResult = e.ToString
                End Try
            End With
            Return mResult
        End Function
    End Class
    Public Sub CopyDIR(src() As String, tar As String)
        If Not tar.EndsWith("\") Then tar &= "\"
        If My.Computer.FileSystem.DirectoryExists(tar) Then
            For i As Integer = 0 To src.Count - 1
                Console.WriteLine()
                Try
                    If My.Computer.FileSystem.DirectoryExists(src(i)) Then
                        Dim di As IO.DirectoryInfo = New IO.DirectoryInfo(src(i))
                        If LinkDIR Then
                            If My.Computer.FileSystem.DirectoryExists(tar & di.Name) Or My.Computer.FileSystem.FileExists(tar & di.Name) Then
                                Console.WriteLine(">>错误：已存在" & tar & di.Name & "")
                                err = True
                                edir += 1
                            Else
                                Console.WriteLine("*" & di.Attributes.ToString & "")
                                Console.WriteLine("mklink" & " /j """ & tar & di.Name & """ """ & di.FullName & """")
                                Console.WriteLine(">>" & ConsoleOutput.ExcuteCmd("mklink" & " /j """ & tar & di.Name & """ """ & di.FullName & """"))
                                sdir += 1
                            End If
                        Else
                            If My.Computer.FileSystem.DirectoryExists(tar & di.Name) Or My.Computer.FileSystem.FileExists(tar & di.Name) Then
                                Console.WriteLine(">>错误：已存在" & tar & di.Name & "")
                                err = True
                                edir += 1
                            Else
                                Console.WriteLine("*" & di.Attributes.ToString & "")
                                Console.WriteLine("md """ & tar & di.Name & """")
                                Console.WriteLine(">>" & ConsoleOutput.ExcuteCmd("md """ & tar & di.Name & """"))
                                Dim s As New List(Of String)
                                For Each fn As String In My.Computer.FileSystem.GetDirectories(src(i))
                                    s.Add(fn)
                                Next
                                For Each fn As String In My.Computer.FileSystem.GetFiles(src(i))
                                    s.Add(fn)
                                Next
                                CopyDIR(s.ToArray, tar & di.Name)
                                sdir += 1
                            End If
                        End If

                    ElseIf My.Computer.FileSystem.FileExists(src(i)) Then
                        Dim fl As IO.FileInfo = New IO.FileInfo(src(i))
                        If My.Computer.FileSystem.DirectoryExists(tar & fl.Name) Or My.Computer.FileSystem.FileExists(tar & fl.Name) Then
                            Console.WriteLine(">>错误：已存在" & tar & fl.Name & "")
                            err = True
                            efile += 1
                        Else
                            Console.WriteLine("*" & fl.Attributes.ToString & "")
                            Console.WriteLine("mklink" & " """ & tar & fl.Name & """ """ & fl.FullName & """")
                            Console.WriteLine(">>" & ConsoleOutput.ExcuteCmd("mklink" & " """ & tar & fl.Name & """ """ & fl.FullName & """"))
                            sfile += 1
                        End If
                    Else
                        Console.WriteLine(">>" & src(i) & "不是文件或目录，跳过...")
                        err = True
                    End If
                Catch ex As Exception
                    Console.WriteLine(">>" & ex.ToString)
                    err = True
                End Try
            Next
        Else
            err = True
            Console.WriteLine()
            Console.WriteLine(">>错误：当前文件夹已被删除")
        End If
    End Sub

    Sub Main()

        Dim tar As String = Microsoft.VisualBasic.Command.Replace("""", "")
        If tar.EndsWith(">") Then
            tar = tar.Replace(">", "")
            LinkDIR = False
        End If
        Dim src() As String = Clipboard.GetFileDropList.Cast(Of String).ToArray
        If tar = "" Then tar = My.Computer.FileSystem.CurrentDirectory
        If Not tar.EndsWith("\") Then tar &= "\"
        Console.WriteLine("目标文件夹：" & tar)
        Console.WriteLine("选中的对象：")
        For i As Integer = 0 To src.Count - 1
            Console.WriteLine("    " & src(i))
        Next
        Console.WriteLine()

        If src.Count = 0 Then
            err = True
            Console.WriteLine()
            Console.WriteLine(">>错误：未复制任何文件/目录")
        End If
        CopyDIR(src, tar)
        Dim canc As Boolean = False
        Console.WriteLine()
        Console.WriteLine("==========================")
        Console.WriteLine("已处理 " & sfile & "个文件、" & sdir & "个目录")
        If Not err Then
            Console.WriteLine("运行完毕，1秒后自动退出，按任意键取消")
        Else
            Console.WriteLine("未处理 " & efile & "个文件、" & edir & "个目录")
            Console.WriteLine()
            Console.WriteLine("出错，请检查错误信息！")
            canc = True
        End If
        Dim th As New Threading.Thread(
                Sub()
                    Threading.Thread.Sleep(1000)
                    If Not canc Then End
                End Sub)
        th.Start()
        Console.ReadKey()
        canc = True
        While canc
            Console.ReadKey()
        End While
    End Sub

End Module
