Imports System.Collections.Generic
Imports System.Text

Imports System.Runtime.InteropServices

'Namespace NetworkShare
Class WNETClass
    <StructLayout(LayoutKind.Sequential)> _
    Friend Structure NETRESOURCE
        Public dwScope As Integer
        Public dwType As resType
        Public dwDisplayType As Integer
        Public dwUsage As Integer
        <MarshalAs(UnmanagedType.LPWStr)> _
        Public lpLocalName As String
        <MarshalAs(UnmanagedType.LPWStr)> _
        Public lpRemoteName As String
        <MarshalAs(UnmanagedType.LPWStr)> _
        Public lpComment As String
        <MarshalAs(UnmanagedType.LPWStr)> _
        Public lpProvider As String
    End Structure

    <DllImport("coredll.dll", EntryPoint:="WNetAddConnection3", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function WNetAddConnection3(ByVal hWnd As IntPtr, ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUsername As String, ByVal dwFlag As Int32) As Integer
    End Function

    <DllImport("coredll.dll", EntryPoint:="WNetAddConnection3", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function WNetAddConnection3(ByVal hWnd As IntPtr, ByRef lpNetResource As NETRESOURCE, ByVal lpPassword As String, ByVal lpUsername As String, ByVal dwFlag As dwFlags) As Integer
    End Function

    <DllImport("coredll.dll", EntryPoint:="WNetCancelConnection2", CharSet:=System.Runtime.InteropServices.CharSet.Unicode, SetLastError:=True)> _
    Private Shared Function WNetCancelConnection2(ByVal lpName As String, ByVal dwFlags As dwFlags, ByVal fForce As Boolean) As Integer
    End Function

    <FlagsAttribute()> _
    Public Enum resType As Integer
        RESOURCETYPE_ANY = &H0
        RESOURCETYPE_DISK = &H1
        RESOURCETYPE_PRINT = &H2
    End Enum
    <FlagsAttribute()> _
    Public Enum dwFlags As UInteger
        CONNECT_UPDATE_PROFILE = &H1
        CONNECT_UPDATE_RECENT = &H2
        CONNECT_TEMPORARY = &H4
        CONNECT_INTERACTIVE = &H8
        CONNECT_PROMPT = &H10
        CONNECT_NEED_DRIVE = &H20
        CONNECT_REFCOUNT = &H40
        CONNECT_REDIRECT = &H80
        CONNECT_LOCALDRIVE = &H100
        CONNECT_CURRENT_MEDIA = &H200
    End Enum


    Private Const RESOURCETYPE_ANY As Integer = &H0
    Private Const CONNECT_INTERACTIVE As Integer = &H8
    Private Const CONNECT_PROMPT As Integer = &H10

#Region "Errors"
    '
    ' MessageId: ERROR_NOT_SUPPORTED
    '
    ' MessageText:
    '
    '  The request is not supported.
    '
    Private Const ERROR_NOT_SUPPORTED As Integer = 50

    '
    ' MessageId: ERROR_REM_NOT_LIST
    '
    ' MessageText:
    '
    '  Windows cannot find the network path. Verify that the network path is correct and the destination computer is not busy or turned off. If Windows still cannot find the network path, contact your network administrator.
    '
    Private Const ERROR_REM_NOT_LIST As Integer = 51

    '
    ' MessageId: ERROR_DUP_NAME
    '
    ' MessageText:
    '
    '  You were not connected because a duplicate name exists on the network. Go to System in Control Panel to change the computer name and try again.
    '
    Private Const ERROR_DUP_NAME As Integer = 52

    '
    ' MessageId: ERROR_BAD_NETPATH
    '
    ' MessageText:
    '
    '  The network path was not found.
    '
    Private Const ERROR_BAD_NETPATH As Integer = 53

    '
    ' MessageId: ERROR_NETWORK_BUSY
    '
    ' MessageText:
    '
    '  The network is busy.
    '
    Private Const ERROR_NETWORK_BUSY As Integer = 54

    '
    ' MessageId: ERROR_DEV_NOT_EXIST
    '
    ' MessageText:
    '
    '  The specified network resource or device is no longer available.
    '
    Private Const ERROR_DEV_NOT_EXIST As Integer = 55
    ' dderror
    Private Const ERROR_INVALID_PARAMETER As Integer = 87

#End Region

    Public Function doConnect(ByVal inShare As String) As Boolean
        Dim bRet As Boolean = False
        Dim ConnInf As New NETRESOURCE()

        ConnInf.dwScope = 0
        'ConnInf.dwType = RESOURCETYPE_ANY;
        ConnInf.dwType = resType.RESOURCETYPE_ANY
        ' resType.RESOURCETYPE_DISK;
        ConnInf.dwDisplayType = 0
        ConnInf.dwUsage = 0
        ConnInf.lpLocalName = "share"
        ' null;
        ConnInf.lpRemoteName = inShare & vbNullChar
        ConnInf.lpComment = Nothing
        ConnInf.lpProvider = Nothing

        WNetAddConnection3(IntPtr.Zero, ConnInf, Nothing, Nothing, dwFlags.CONNECT_REDIRECT Or dwFlags.CONNECT_UPDATE_PROFILE)
        ' CONNECT_INTERACTIVE | CONNECT_PROMPT);
        Return bRet
    End Function

    Public Function doConnect(ByVal inShare As String, ByVal sPassword As String, ByVal sUser As String, ByVal sFileName As String) As Boolean
        Dim bRET As Boolean
        Dim ConnInf As New NETRESOURCE()
        ConnInf.dwScope = 0
        ConnInf.dwType = resType.RESOURCETYPE_ANY
        ConnInf.dwDisplayType = 0
        ConnInf.dwUsage = 0
        ConnInf.lpLocalName = "share"
        ConnInf.lpRemoteName = inShare
        ConnInf.lpComment = Nothing
        ConnInf.lpProvider = Nothing

        If WNetAddConnection3(IntPtr.Zero, ConnInf, sPassword, sUser, dwFlags.CONNECT_UPDATE_PROFILE Or dwFlags.CONNECT_LOCALDRIVE) = 0 Then
            Dim iErr As Integer = Marshal.GetLastWin32Error()
            Return iErr
        Else
            Dim AppPath As String
            AppPath = System.IO.Path.GetDirectoryName(System.Reflection.Assembly.GetExecutingAssembly().GetName().CodeBase)
            If Not AppPath.EndsWith("\") Then
                AppPath += "\"
            End If
            Try
                System.IO.File.Copy(inShare & sFileName, AppPath & sFileName, True)
                WNetCancelConnection2(ConnInf.lpLocalName, dwFlags.CONNECT_UPDATE_PROFILE, True)
                bRET = True
            Catch ex As Exception
                MessageBox.Show(ex.Message)
            End Try
        End If
        Return bRET
    End Function


End Class
'End Namespace

