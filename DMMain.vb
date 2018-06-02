
Imports System.Runtime.InteropServices
Imports System.Reflection


Module MDMain

    Private oOneInstance As SingleInstanceApplication = Nothing
    Sub Main()
        oOneInstance = New SingleInstanceApplication()
        oOneInstance.Run()
        If oOneInstance.OK Then Application.Run(Form1)
    End Sub

End Module
Public Class SingleInstanceApplication
    <DllImport("coredll.dll", SetLastError:=True)> _
    Private Shared Function CreateMutex(ByVal Attr As IntPtr, ByVal Own As Boolean, ByVal Name As String) As Integer ''''http://tech.sina.com.cn/s/2009-12-14/00481172103.shtml
    End Function

    <DllImport("coredll.dll", SetLastError:=True)> _
    Private Shared Function ReleaseMutex(ByVal hMutex As IntPtr) As Boolean
    End Function

    Const ERROR_ALREADY_EXISTS As Long = 183
    Public OK As Boolean

    Public Sub Run()
        Dim name As String = Assembly.GetExecutingAssembly().GetName().Name
        Dim mutexHandle As IntPtr = CreateMutex(IntPtr.Zero, True, name)
        Dim xError As Long = Marshal.GetLastWin32Error()



        If xError <> ERROR_ALREADY_EXISTS Then
            'frm.Show()
            'Application.Run(frm)
            OK = True
        Else
            'MsgBox("程序已经开启，请勿重复运行!")
            'frm.Close()
        End If
        ReleaseMutex(mutexHandle)
    End Sub
End Class