'Option Strict Off
'Option Explicit On=
Imports Microsoft.VisualBasic
Imports System.Text
Imports System.Threading
Imports HomeSeerAPI


Public Class PIHSPI

    Public link As String       ' actual link string
    Public linktext As String   ' display text for link
    Public page_title As String ' title of web page




    Public Sub New()
    End Sub

    Private Sub SetupEnvironment()
        Try
            Dim sver As String = ""
            Dim ourver As String = ""
            Dim EXEPath As String = System.IO.Path.GetDirectoryName(System.AppDomain.CurrentDomain.BaseDirectory)

            If System.IO.File.Exists(EXEPath & "\html\bin\HSPI_OCELOT.DLL") Then
                Try
                    sver = GetVersionInfo(EXEPath & "\html\bin\HSPI_OCELOT.DLL")
                    ourver = GetVersionInfo(EXEPath & "\HSPI_OCELOT.DLL")
                Catch ex As Exception
                End Try
                If sver <> ourver Then
                    Try
                        System.IO.File.Copy(EXEPath & "\HSPI_OCELOT.DLL", EXEPath & "\html\bin\HSPI_OCELOT.DLL", True)
                    Catch ex As Exception
                        hs.WriteLog("Ocelot", "Error copying HSPI_OCELOT.DLL to html\bin folder, configuration may not be accessible. " & ex.Message)
                    End Try
                End If
            Else
                Try
                    System.IO.File.Copy(EXEPath & "\HSPI_OCELOT.DLL", EXEPath & "\html\bin\HSPI_OCELOT.DLL", True)
                Catch ex As Exception
                    hs.WriteLog("Ocelot", "Error copying HSPI_OCELOT.DLL to html\bin folder, configuration may not be accessible. " & ex.Message)
                End Try
            End If

            Dim rval As String
            Try
                rval = GetResourceToFile("OcelotConfig.aspx", EXEPath & "\html\OcelotConfig.aspx")
                If rval <> "" Then
                    hs.WriteLog("Ocelot", "Error, unable to save Ocelot configuration page to HTML folder, extraction error: " & rval)
                End If
            Catch ex As Exception
                hs.WriteLog("Ocelot", "Error, unable to save Ocelot configuration page to HTML folder: " & ex.Message)
            End Try
        Catch ex As Exception
            hs.WriteLog("Ocelot", "Error setting up environment: " & ex.Message)
        End Try
    End Sub


    '
    Public Sub RegisterCallback(ByRef frm As Object)
        On Error Resume Next
        callback = frm

        hs = frm.GetHSIface

        InterfaceVersion = hs.InterfaceVersion
        SetupEnvironment()
    End Sub

    ' return status of interface
    ' see util.bas for constants
    Public Function InterfaceStatus() As Integer
        InterfaceStatus = Status
    End Function

    ' ******************************** X10 Interface ********************

    Public Sub ConfigX10()
    End Sub

    ' return the status of an individual device
    ' return -1 if the status is not available
    ' if the interface keeps track of X10 devices, the actual status of the
    ' device should be returned (X10 command like ON,OFF,DIM)
    Public Function GetDeviceStatus(ByRef housecode As String, ByRef devicecode As String) As Integer

    End Function


    ' ******************************** Infrared Interface ********************

    ' return the correct port type for your device
    ' this allows HS to display the proper selections in the
    ' interfaces tab
    Public Function IRPortType() As Integer
        IRPortType = IRC_SERIAL
    End Function


    ' max keys per device
    ' must be in the range 1-36
    Public ReadOnly Property MaxKeys() As Short
        Get
            MaxKeys = 36
        End Get
    End Property

    Public ReadOnly Property MaxDevices() As Integer
        Get
            If gIRSize <> 0 Then
                MaxDevices = gIRSize / 36
            Else
                MaxDevices = 14
            End If

        End Get
    End Property

    ' return TRUE if this device has an options dialog
    ' the dialog will be called with a call to Config
    Public ReadOnly Property HasOptions() As Boolean
        Get
            HasOptions = False
        End Get
    End Property


    ' return TRUE if device supports IR matching
    ' this will allow HS to display special screens to handle
    ' IR events when a match is detected
    Public ReadOnly Property SupportsIRMatch() As Boolean
        Get
            SupportsIRMatch = True
        End Get
    End Property



    ' return max # of zones supported by the interface
    Public ReadOnly Property MaxZones() As Integer
        Get
            If gSecu16IR_present Then
                MaxZones = 17
            End If
        End Get
    End Property

    ' return true if the given IR location is matchable on IR recieve
    ' rjh 11/1/2010 changed max match locations from 80 to 1024, the default is 80, and 1024 can be set with one of the parameters
    Public Function IsMatchable(ByRef loc_Renamed As Integer) As Boolean
        If loc_Renamed < 1024 Then
            Return True
        Else
            Return False
        End If
    End Function

    ' display config dialog for any special options this device supports
    ' such as special device codes for equipment
    Public Sub ConfigIR()
    End Sub


    ' return the color of the key for given IR location
    ' this is used to display different colors for keys
    ' use black for learned, and red for unlearned keys
    'UPGRADE_NOTE: loc was upgraded to loc_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Function KeyColor(ByRef loc_Renamed As Integer) As Integer

        'KeyColor = &HFF&    ' red=learned

        KeyColor = 0 ' black unknown/not learned
    End Function

    ' return true if the IR location is learnable
    ' this will disable the learn key in the GUI
    'UPGRADE_NOTE: loc was upgraded to loc_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Function KeyLearnable(ByRef loc_Renamed As Integer) As Boolean
        KeyLearnable = True
    End Function





End Class

