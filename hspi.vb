Imports HSCF
Imports HomeSeerAPI
Imports Scheduler
Imports HSCF.Communication.ScsServices.Service
Imports System.Reflection
Imports System.Text
Imports System.Threading
Imports System.Web

Public Class HSPI
    Inherits ScsService
    Implements IPlugInAPI

    Public OurInstanceFriendlyName As String = ""
    Public Shared bShutDown As Boolean = False
    'Private plugin As New PIHSPI

    Private WithEvents MSComm1 As System.IO.Ports.SerialPort 'rs232
    Public rbuffer(500) As Byte
    Private PollThread As Thread
    Private GotRcvData As Boolean
    Private Structure q_entry
        Dim hcc As String
        Dim dc As String
        Dim cm As Short
        Dim br As Integer
        Dim d1 As Short
        Dim d2 As Short
        Dim unit As Integer
        Dim point As Integer
        Dim variable As Integer
        Dim dvref As Integer
    End Structure
    Dim gSendBusy As Short
    Dim grefs As Short
    Dim gAunit As Short
    Dim gApoint As Short
    Dim gAinit As Boolean
    Dim gAlast As Date
    Dim gLoVar As Short
    Dim gHiVar As Short
    Dim gLearnWait As Boolean
    Dim gCommTOCount As Short ' track comm timeouts
    Dim gCommCRCCount As Short ' track CRC errors
    Dim gSetZone As Short ' current zone to use
    Const QMAX As Short = 1000 ' max number of commands to queue
    Dim gCmdQueue(QMAX) As q_entry
    Dim gQtail As Short
    Dim gQhead As Short
    Dim gQsize As Short










    Public Shared Function ConfigUnits(unit_num As Integer) As Byte
        Return gUnits(unit_num)
    End Function
    Public Shared Function ConfigUnits() As Byte()
        Return gUnits
    End Function

#If 0 Then
    Public Property ConfigComPort() As String
        Get
            ConfigComPort = gComPort
        End Get
        Set(ByVal value As String)
            If gComPort <> value Then
                gComPort = value
                SaveSettingLocal("hspi_ocelot", "Settings", "comport", gComPort.Trim)
            End If
        End Set
    End Property
#End If

    Public Shared Property ConfigPollInterval() As Integer
        Get
            ConfigPollInterval = gPollInterval
        End Get
        Set(ByVal value As Integer)
            gPollInterval = value
            SaveSettingLocal("hspi_ocelot", "Settings", "poll_interval", gPollInterval.ToString)
        End Set
    End Property

    Public Shared Property ConfigInvertInput() As Boolean
        Get
            ConfigInvertInput = gInvert
        End Get
        Set(ByVal value As Boolean)
            gInvert = value
            If gInvert Then
                SaveSettingLocal("hspi_ocelot", "Settings", "invert", 1)
            Else
                SaveSettingLocal("hspi_ocelot", "Settings", "invert", 0)
            End If
        End Set
    End Property

    Public Shared Property ConfigLogErrors() As Boolean
        Get
            ConfigLogErrors = gLogErrors
        End Get
        Set(ByVal value As Boolean)
            gLogErrors = value
            If gLogErrors Then
                SaveSettingLocal("hspi_ocelot", "Settings", "log_errors", 1)
            Else
                SaveSettingLocal("hspi_ocelot", "Settings", "log_errors", 0)
            End If
        End Set
    End Property

    Public Shared Property ConfigLogComms() As Boolean
        Get
            ConfigLogComms = gLogComms
        End Get
        Set(ByVal value As Boolean)
            gLogComms = value
            If gLogComms Then
                SaveSettingLocal("hspi_ocelot", "Settings", "log_comms", 1)
            Else
                SaveSettingLocal("hspi_ocelot", "Settings", "log_comms", 0)
            End If
        End Set
    End Property

    Public Shared Property ConfigPollVars() As Boolean
        Get
            ConfigPollVars = gConfigPollVars
        End Get
        Set(ByVal value As Boolean)
            gConfigPollVars = value
            SaveSettingLocal("hspi_ocelot", "Settings", "gConfigPollVars", gConfigPollVars.ToString)
        End Set
    End Property

    Public Sub ConfigCreateVarDevices()
        CreateVarDevices()
    End Sub

    Public Sub ConfigCreateIODevices(ByVal unit As Integer)
        CreateIODevices(unit)
    End Sub

    Public Sub ConfigDeleteAllDevices()
        DeleteAllDevices()
    End Sub

    Public Sub ConfigCreateLeopardDevice()
        CreateLeopardDevice()
    End Sub

    Public Function AccessLevel() As Integer Implements HomeSeerAPI.IPlugInAPI.AccessLevel
        Return 1
    End Function

    Public Property ActionAdvancedMode As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionAdvancedMode
        Get

        End Get
        Set(value As Boolean)

        End Set
    End Property

    Public Function ActionBuildUI(sUnique As String, ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionBuildUI
        Return ""
    End Function

    Public Function ActionConfigured(ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionConfigured
        Return False
    End Function

    Public Function ActionCount() As Integer Implements HomeSeerAPI.IPlugInAPI.ActionCount
        Return 0
    End Function

    Public Function ActionFormatUI(ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionFormatUI
        Return ""
    End Function

    Public ReadOnly Property ActionName(ActionNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.ActionName
        Get
            Return ""
        End Get
    End Property

    Public Function ActionProcessPostUI(PostData As System.Collections.Specialized.NameValueCollection, TrigInfoIN As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.ActionProcessPostUI
        Return Nothing
    End Function

    Public Function ActionReferencesDevice(ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo, dvRef As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionReferencesDevice
        Return False
    End Function

    Public Function Capabilities() As Integer Implements HomeSeerAPI.IPlugInAPI.Capabilities
        Return HomeSeerAPI.Enums.eCapabilities.CA_IO
    End Function

    Public Property Condition(TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.Condition
        Get

        End Get
        Set(value As Boolean)

        End Set
    End Property

    Public Function ConfigDevice(ref As Integer, user As String, userRights As Integer, newDevice As Boolean) As String Implements HomeSeerAPI.IPlugInAPI.ConfigDevice
        Dim dv As Scheduler.Classes.DeviceClass = Nothing
        Dim stb As New StringBuilder

        dv = hs.GetDeviceByRef(ref)
        If dv Is Nothing Then
            stb.Append(HTML_NewLine & HTML_NewLine)
            stb.Append(HTML_StartBold & HTML_StartFont(COLOR_RED))
            stb.Append("Error, device was not successfully retrieved from HomeSeer - Cannot Continue.")
            stb.Append(HTML_EndFont & HTML_EndBold)
            Return stb.ToString
        End If

        If dv.Interface(Nothing) <> IFACE_NAME Then
            stb.Append(HTML_NewLine & HTML_NewLine)
            stb.Append(HTML_StartBold & HTML_StartFont(COLOR_RED))
            stb.Append("Error, the device is not owned by this plug-in; Will not Continue.")
            stb.Append(HTML_EndFont & HTML_EndBold)
            Return stb.ToString
        End If

        'Dim PED As clsPlugExtraData = dv.PlugExtraData_Get(hs)
        Dim Find_Unit_Type As CPUXA_UnitTypes
        Dim Find_Unit As Short
        Dim Find_Point As Short
        Dim Find_IO As enumIO



        Try
            stb.Append(clsPageBuilder.FormStart("OcelotDeviceForm_" & ref.ToString, "OcelotDeviceFormName_" & ref.ToString, "post"))
            stb.Append(HTML_StartTable(0, , 50, , , HTML_TableAlign.LEFT) & vbCrLf)

            If Not GetPEDInfo(dv, Find_Unit_Type, Find_Unit, Find_Point, Find_IO) Then
                Dim PED As HomeSeerAPI.clsPlugExtraData = Nothing
                PED = dv.PlugExtraData_Get(hs)
                If PED IsNot Nothing AndAlso PED.NamedCount > 0 Then
                    stb.Append(HTML_StartRow() & vbCrLf)
                    stb.Append(HTML_StartCell("", 1, HTML_Align.RIGHT, True) & vbCrLf)
                    stb.Append("Diagnostic:")
                    stb.Append(HTML_EndCell & vbCrLf)
                    stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True) & vbCrLf)
                    For Each s As String In PED.GetNamedKeys
                        Try
                            stb.Append(s & " = " & PED.GetNamed(s).ToString & HTML_NewLine)
                        Catch ex As Exception
                            stb.Append(s & " = Error Retrieving Value" & HTML_NewLine)
                        End Try
                    Next
                    stb.Append(HTML_EndCell & vbCrLf)
                    stb.Append(HTML_EndRow & vbCrLf)
                End If

            Else

                stb.Append(HTML_StartRow() & vbCrLf)
                stb.Append(HTML_StartCell("", 1, HTML_Align.RIGHT, True) & vbCrLf)
                stb.Append("Unit Type:")
                stb.Append(HTML_EndCell & vbCrLf)
                stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True) & vbCrLf)
                stb.Append(Unit_Type_String(Find_Unit_Type))
                stb.Append(HTML_EndCell & vbCrLf)
                stb.Append(HTML_EndRow & vbCrLf)

                If Find_Unit_Type = CPUXA_UnitTypes.Leopard Or Find_Unit_Type = CPUXA_UnitTypes.Variable Then
                    ' Unit is not applicable.
                Else
                    stb.Append(HTML_StartRow() & vbCrLf)
                    stb.Append(HTML_StartCell("", 1, HTML_Align.RIGHT, True) & vbCrLf)
                    stb.Append("Unit ID:")
                    stb.Append(HTML_EndCell & vbCrLf)
                    stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True) & vbCrLf)
                    stb.Append(Find_Unit.ToString)
                    stb.Append(HTML_EndCell & vbCrLf)
                    stb.Append(HTML_EndRow & vbCrLf)
                End If

                If Find_Unit_Type = CPUXA_UnitTypes.Leopard Then
                Else
                    stb.Append(HTML_StartRow() & vbCrLf)
                    stb.Append(HTML_StartCell("", 1, HTML_Align.RIGHT, True) & vbCrLf)
                    If Find_Unit_Type = CPUXA_UnitTypes.Variable Then
                        stb.Append("Variable Number:")
                    Else
                        stb.Append("I/O Point on Unit:")
                    End If
                    stb.Append(HTML_EndCell & vbCrLf)
                    stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True) & vbCrLf)
                    stb.Append(Find_Point.ToString)
                    stb.Append(HTML_EndCell & vbCrLf)
                    stb.Append(HTML_EndRow & vbCrLf)
                End If

                Select Case Find_Unit_Type
                    Case CPUXA_UnitTypes.Leopard, CPUXA_UnitTypes.Variable
                    Case Else
                        stb.Append(HTML_StartRow() & vbCrLf)
                        stb.Append(HTML_StartCell("", 1, HTML_Align.RIGHT, True) & vbCrLf)
                        stb.Append("I/O Point Type:")
                        stb.Append(HTML_EndCell & vbCrLf)
                        stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True) & vbCrLf)
                        If Find_IO = enumIO._Input Then
                            stb.Append("Input")
                        ElseIf Find_IO = enumIO._Output Then
                            stb.Append("Output")
                        Else
                            stb.Append(HTML_StartFont(COLOR_RED) & "Uh oh. Not Input or Output" & HTML_EndFont)
                        End If
                        stb.Append(HTML_EndCell & vbCrLf)
                        stb.Append(HTML_EndRow & vbCrLf)
                End Select

                stb.Append(HTML_StartRow() & vbCrLf)
                stb.Append(HTML_StartCell("", 1, HTML_Align.RIGHT, True) & vbCrLf)
                stb.Append("Click to set the device 'Address' property to the plug-in default.")
                stb.Append(HTML_NewLine)
                stb.Append("(Does not change the device 'Code' value if set.)")
                stb.Append(HTML_EndCell & vbCrLf)
                stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True) & vbCrLf)
                Dim bSave As New clsJQuery.jqButton("ReAddress", "Reset Address", "DeviceUtility", True)
                stb.Append(bSave.Build)
                stb.Append(HTML_EndCell & vbCrLf)
                stb.Append(HTML_EndRow & vbCrLf)


            End If

            stb.Append(HTML_StartRow() & vbCrLf)
            stb.Append(HTML_StartCell("", 2, HTML_Align.LEFT, True) & vbCrLf)
            Dim butt As New clsJQuery.jqButton("Ref_" & ref.ToString & "_Cancel", "Done", "deviceutility", True)
            stb.Append(butt.Build)
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)

            stb.Append(HTML_EndTable & vbCrLf)

            stb.Append(clsPageBuilder.FormEnd)

            Return stb.ToString

        Catch ex As Exception
            stb = New StringBuilder
            stb.Append(HTML_NewLine & HTML_NewLine)
            stb.Append(HTML_StartBold & HTML_StartFont(COLOR_RED))
            stb.Append("Error, building device properties for device with ref=" & ref.ToString & " encountered an exception: ")
            stb.Append(HTML_EndFont & HTML_EndBold)
            stb.Append(HTML_NewLine)
            stb.Append(ex.Message)
            stb.Append(HTML_NewLine)
            Return stb.ToString
        End Try

    End Function

    Public Function ConfigDevicePost(ref As Integer, data As String, user As String, userRights As Integer) As HomeSeerAPI.Enums.ConfigDevicePostReturn Implements HomeSeerAPI.IPlugInAPI.ConfigDevicePost
        Dim dv As Scheduler.Classes.DeviceClass = Nothing
        Dim parts As Collections.Specialized.NameValueCollection
        Dim ReturnValue As Integer = Enums.ConfigDevicePostReturn.DoneAndCancel

        Dim Find_Unit_Type As CPUXA_UnitTypes
        Dim Find_Unit As Short
        Dim Find_Point As Short
        Dim Find_IO As enumIO

        Try
            parts = HttpUtility.ParseQueryString(data)

            If parts("id").Trim.ToLower.EndsWith("_cancel") Then
                Return Enums.ConfigDevicePostReturn.DoneAndCancel
            End If

            dv = hs.GetDeviceByRef(ref)
            If dv IsNot Nothing AndAlso dv.Interface(Nothing) = IFACE_NAME Then
                If parts("id") = "ReAddress" Then
                    If GetPEDInfo(dv, Find_Unit_Type, Find_Unit, Find_Point, Find_IO) Then
                        hs.WriteLogEx(IFACE_OTHER, "Resetting Address on " & hs.DeviceName(ref), COLOR_ORANGE)
                        SetHSDeviceAddress(dv, Find_Unit_Type, Find_Unit, Find_Point, Find_IO)
                        hs.SaveEventsDevices()
                        ReturnValue = Enums.ConfigDevicePostReturn.DoneAndSave
                    End If
                End If
            End If

            Return ReturnValue
        Catch ex As Exception
            hs.WriteLogEx(IFACE_OTHER, "Exception processing POST data for ConfigDevice: " & ex.Message, COLOR_RED)
        End Try
        Return ReturnValue
    End Function

    Public Function GenPage(link As String) As String Implements HomeSeerAPI.IPlugInAPI.GenPage
        Return ""
    End Function

    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String Implements HomeSeerAPI.IPlugInAPI.GetPagePlugin
        If ConfigPage IsNot Nothing AndAlso pageName = ConfigPage.PageName Then
            Return (ConfigPage.GetPagePlugin(pageName, user, userRights, queryString))
        End If
        Return "Error, " & pageName & " (GET) is invalid for this plug-in, has not been initialized yet, or has terminated in an error."
    End Function
    Public Function PostBackProc(ByVal pageName As String, ByVal data As String, ByVal user As String, ByVal userRights As Integer) As String Implements HomeSeerAPI.IPlugInAPI.PostBackProc
        If ConfigPage IsNot Nothing AndAlso pageName = ConfigPage.PageName Then
            Return ConfigPage.postBackProc(pageName, data, user, userRights)
        End If
        Return "Error, " & pageName & " (POST) is invalid for this plug-in, has not been initialized yet, or has terminated in an error."
    End Function

    Public Function HandleAction(ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.HandleAction
        Return False
    End Function

    Public ReadOnly Property HasConditions(TriggerNumber As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.HasConditions
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property HasTriggers As Boolean Implements HomeSeerAPI.IPlugInAPI.HasTriggers
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property HSCOMPort As Boolean Implements HomeSeerAPI.IPlugInAPI.HSCOMPort
        Get
            Return True
        End Get
    End Property

    Public Sub HSEvent(EventType As HomeSeerAPI.Enums.HSEvent, parms() As Object) Implements HomeSeerAPI.IPlugInAPI.HSEvent

    End Sub

    Public Function InitIO(port As String) As String Implements HomeSeerAPI.IPlugInAPI.InitIO
        ' just exit already initialized, this function maybe called multiple times
        Dim rval As String = ""

        If gIOEnabled Then Return "" 'gIOInitialized Then Return ""

        If gIREnabled Or gX10Enabled Then
            SetAutoIO(True)
            'PollVars()
            gIOEnabled = True
        End If

        LoadAllSettings()

        If Not gmoduleOK Then
            ' get our comport (4/13/2014 Changing to using HS COM Port given to us.)
            gComPort = port 'GetSettingLocal("hspi_ocelot", "Settings", "comport", "COM1")
        End If

        rval = InitCPUXA()
        If rval <> "" Then
            Return ""
        End If

        'gIOInitialized = True
        gIOEnabled = True
        InitFirstRun()
        SetAutoIO(True)
        'Wait 2
        ' do a single poll to get initial IO before HS is fully running
        ' any changes detected here will not trigger events
        'PollVars()
        Return ""
    End Function


    Private Sub InitFirstRun()
        Dim ft As Integer

        gPollInterval = GetSettingLocal("hspi_ocelot", "Settings", "poll_interval", "500")

        ft = Convert.ToInt32(GetSettingLocal("hspi_ocelot", "Settings", "first_run", "0"))
        If ft = 0 Then
            hs.WriteLog(IFACE_OTHER, "This is the first time this plugin has been loaded, use the Config button on the Interfaces page to configure")
            Try
                HSPI.ConfigInvertInput = False
                HSPI.ConfigLogComms = False
                HSPI.ConfigLogErrors = False
                HSPI.ConfigPollInterval = 500
                HSPI.ConfigPollVars = False
            Catch ex As Exception
            End Try
        End If

        SaveSettingLocal("hspi_ocelot", "Settings", "first_run", "1")
        CreateUnitMappings()
    End Sub

    Public Function InstanceFriendlyName() As String Implements HomeSeerAPI.IPlugInAPI.InstanceFriendlyName
        Return ""
    End Function

    Public Function InterfaceStatus() As HomeSeerAPI.IPlugInAPI.strInterfaceStatus Implements HomeSeerAPI.IPlugInAPI.InterfaceStatus
        Dim es As New IPlugInAPI.strInterfaceStatus
        es.intStatus = IPlugInAPI.enumInterfaceStatus.OK
        Return es
    End Function

    Public ReadOnly Property Name As String Implements HomeSeerAPI.IPlugInAPI.Name
        Get
            Return IFACE_NAME
        End Get
    End Property

    Public Function PagePut(data As String) As String Implements HomeSeerAPI.IPlugInAPI.PagePut
        Return "" 'plugin.PagePut(data)
    End Function

    ' a custom call to call a specific procedure in the plugin
    Public Function PluginFunction(ByVal proc As String, ByVal parms() As Object) As Object Implements IPlugInAPI.PluginFunction
        Try
            Dim ty As Type = Me.GetType
            Dim mi As MethodInfo = ty.GetMethod(proc)
            If mi Is Nothing Then
                hs.WriteLog(IFACE_OTHER, "Method " & proc & " does not exist in this plugin.")
                Return Nothing
            End If
            Return (mi.Invoke(Me, parms))
        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error in PluginProc: " & ex.Message)
        End Try

        Return Nothing
    End Function

    Public Function PluginPropertyGet(ByVal proc As String, parms() As Object) As Object Implements IPlugInAPI.PluginPropertyGet
        Try
            Dim ty As Type = Me.GetType
            Dim mi As PropertyInfo = ty.GetProperty(proc)
            If mi Is Nothing Then
                hs.WriteLog(IFACE_OTHER, "Method " & proc & " does not exist in this plugin.")
                Return Nothing
            End If
            Return mi.GetValue(Me, Nothing)
        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error in PluginProc: " & ex.Message)
        End Try

        Return Nothing
    End Function

    Public Sub PluginPropertySet(ByVal proc As String, value As Object) Implements IPlugInAPI.PluginPropertySet
        Try
            Dim ty As Type = Me.GetType
            Dim mi As PropertyInfo = ty.GetProperty(proc)
            If mi Is Nothing Then
                hs.WriteLog(IFACE_OTHER, "Property " & proc & " does not exist in this plugin.")
            End If
            mi.SetValue(Me, value, Nothing)
        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error in PluginPropertySet: " & ex.Message)
        End Try
    End Sub

    Public Function PollDevice(dvref As Integer) As HomeSeerAPI.IPlugInAPI.PollResultInfo Implements HomeSeerAPI.IPlugInAPI.PollDevice

    End Function

    Public Function RaisesGenericCallbacks() As Boolean Implements HomeSeerAPI.IPlugInAPI.RaisesGenericCallbacks
        Return False
    End Function

    Public Function Search(SearchString As String, RegEx As Boolean) As HomeSeerAPI.SearchReturn() Implements HomeSeerAPI.IPlugInAPI.Search
        Dim col As New Collections.Generic.List(Of SearchReturn)
        Return col.ToArray
    End Function


    ' send an X10 command over the powerline
    ' the command should be queued as this function can be called
    ' many times
    '
    ' data1 and data2 are the two extended bytes for the X10 Extended command
    Public Sub QExec(ByVal housecode As String, ByVal devicecode As String, ByVal command As Integer, ByVal brightness As Integer, ByRef data1 As Integer, ByRef data2 As Integer, unit As Integer, point As Integer, variable As Integer, dvref As Integer)
        'hs.writelog("Ocelot Debug", "Queuing command")

        If gQsize < QMAX Then
            gCmdQueue(gQhead).hcc = housecode
            gCmdQueue(gQhead).dc = devicecode
            gCmdQueue(gQhead).cm = command
            gCmdQueue(gQhead).br = brightness
            gCmdQueue(gQhead).d1 = data1
            gCmdQueue(gQhead).d2 = data2
            gCmdQueue(gQhead).unit = unit
            gCmdQueue(gQhead).point = point
            gCmdQueue(gQhead).variable = variable
            gCmdQueue(gQhead).dvref = dvref

            gQhead = gQhead + 1
            If gQhead >= QMAX Then
                gQhead = 0
            End If
            gQsize = gQsize + 1
        Else
            hs.WriteLog(IFACE_OTHER, "Warning, Ocelot command queue overflow, size is " & gQsize.ToString)
        End If
        GotRcvData = True   ' force quick handling of queues
        Do
            Thread.Sleep(10)
        Loop While gQsize > 0
    End Sub

    Public Sub SetIOMulti(colSend As System.Collections.Generic.List(Of HomeSeerAPI.CAPI.CAPIControl)) Implements HomeSeerAPI.IPlugInAPI.SetIOMulti
        Try
            Dim CC As CAPIControl
            Dim PED As HomeSeerAPI.PlugExtraData.clsPlugExtraData = Nothing

            Dim Find_Unit_Type As CPUXA_UnitTypes
            Dim Find_Unit As Short
            Dim Find_Point As Short
            Dim Find_IO As enumIO

            For Each CC In colSend
                Console.WriteLine("SetIOMulti set value: " & CC.ControlValue.ToString & "->ref:" & CC.Ref.ToString)
                ' get the device we are controlling
                Dim dv As Classes.DeviceClass = hs.GetDeviceByRef(CC.Ref)
                If dv Is Nothing Then Continue For
                PED = Nothing
                PED = dv.PlugExtraData_Get(hs)
                If PED Is Nothing Then Continue For
                If Not GetPEDInfo(dv, Find_Unit_Type, Find_Unit, Find_Point, Find_IO) Then Continue For
                Select Case Find_Unit_Type
                    Case CPUXA_UnitTypes.Leopard

                    Case CPUXA_UnitTypes.Variable
                        QExec("", 0, VALUE_SET, 0, 0, 0, 0, 0, CC.ControlValue, CC.Ref)

                    Case CPUXA_UnitTypes.Rly08XA, CPUXA_UnitTypes.Secu16
                        If Find_IO <> enumIO._Output Then Continue For
                        QExec("", 0, SET_RELAY, CC.ControlValue, 0, 0, Find_Unit, Find_Point, 0, CC.Ref)

                    Case CPUXA_UnitTypes.Secu16IR

                End Select
            Next
        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error in SetIOMult: " & ex.Message)
        End Try
    End Sub


    Public Sub ShutdownIO() Implements HomeSeerAPI.IPlugInAPI.ShutdownIO
        If Not gX10Enabled And Not gIREnabled Then
            SetAutoIO(False)
            gIOEnabled = False
            Shutdown()
        Else
            SetAutoIO(False)
            gIOEnabled = False
        End If
    End Sub

    Public Sub ShutDownX10()
        If Not gIOEnabled And Not gIREnabled Then
            SetAutoX10(False)
            gX10Enabled = False
            Shutdown()
        Else
            SetAutoX10(False)
            gX10Enabled = False
        End If
    End Sub

    Public Sub ShutDownIR()
        If Not gIOEnabled And Not gX10Enabled Then
            SetAutoIR(False)
            gIREnabled = False
            Shutdown()
        Else
            SetAutoIR(False)
            gIREnabled = False
        End If
    End Sub

    Private Sub Shutdown()
        On Error Resume Next
        PollThread.Abort()
        'MSComm1.PortOpen = False
        FlushCommInput()
        MSComm1.Close()
    End Sub

    ' init function, called when HS starts up
    ' return error string or empty string if ok
    Public Function InitX10(ByVal port As String) As String
        On Error Resume Next
        Dim s As String
        If gX10Enabled Then Return ""
        If gIOEnabled Or gIREnabled Then
            SetAutoX10(True)
            gX10Enabled = True
            Return ""
        End If

        LoadAllSettings()

        gComPort = port
        s = InitCPUXA()
        If s = "" Then
            gX10Enabled = True
            InitFirstRun()
            SetAutoX10(True)
        End If
        Return s
    End Function


    ' init function, called when HS starts up
    ' all the work may have already been done in the InitX10 function
    ' return error string or empty string if ok
    Public Function InitIR(ByRef port As String) As String
        ' just exit already initialized, this function maybe called multiple times
        Dim s As String

        If gIREnabled Then Return "" 'gIRInitialized Then Return ""

        If gIOEnabled Or gX10Enabled Then 'gIOInitialized Or gX10Enabled Then
            'gIREnabled = True
            SetAutoIR(True)
            gIREnabled = True 'gIRInitialized = True
            Return ""
        End If

        LoadAllSettings()

        If Not gmoduleOK Then
            ' get our comport (4/13/2014 Changing to using HS COM Port given to us.)
            gComPort = port 'GetSettingLocal("hspi_ocelot", "Settings", "comport", "COM1")
        End If

        s = InitCPUXA()
        If s <> "" Then Return s

        InitFirstRun()
        'gIRInitialized = True
        gIREnabled = True
        SetAutoIR(True)

        InitIR = s

    End Function

    Public Sub SpeakIn(device As Integer, txt As String, w As Boolean, host As String) Implements HomeSeerAPI.IPlugInAPI.SpeakIn

    End Sub

    Public ReadOnly Property SubTriggerCount(TriggerNumber As Integer) As Integer Implements HomeSeerAPI.IPlugInAPI.SubTriggerCount
        Get
            Return 0
        End Get
    End Property

    Public ReadOnly Property SubTriggerName(TriggerNumber As Integer, SubTriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.SubTriggerName
        Get
            Return ""
        End Get
    End Property

    Public Function SupportsAddDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsAddDevice
        Return False
    End Function

    Public Function SupportsConfigDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDevice
        Return True
    End Function

    Public Function SupportsConfigDeviceAll() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDeviceAll
        Return False
    End Function

    Public Function SupportsMultipleInstances() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstances
        Return False
    End Function

    Public Function SupportsMultipleInstancesSingleEXE() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstancesSingleEXE
        Return False
    End Function

    Public Function TriggerBuildUI(sUnique As String, TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerBuildUI
        Return ""
    End Function

    Public ReadOnly Property TriggerConfigured(TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerConfigured
        Get
            Return False
        End Get
    End Property

    Public ReadOnly Property TriggerCount As Integer Implements HomeSeerAPI.IPlugInAPI.TriggerCount
        Get
            Return 0
        End Get
    End Property

    Public Function TriggerFormatUI(TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerFormatUI
        Return ""
    End Function

    Public ReadOnly Property TriggerName(TriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.TriggerName
        Get
            Return ""
        End Get
    End Property

    Public Function TriggerProcessPostUI(PostData As System.Collections.Specialized.NameValueCollection, TrigInfoIN As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.TriggerProcessPostUI
        Return Nothing
    End Function

    Public Function TriggerReferencesDevice(TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo, dvRef As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerReferencesDevice
        Return False
    End Function

    Public Function TriggerTrue(TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerTrue
        Return False
    End Function





#Region "----- from old HSPI object -----"

    ' start a learn of an IR command
    Private Sub DoLearnIR(ByRef loc As Integer)
        'Dim r As MsgBoxResult
        gbusy = 1
        LearnIRStart(loc)
        'r = MsgBox("Point orignal remote at input window on the Ocelot and press the key you wish to learn." & vbCrLf & "Key will be learned in Ocelot location: " & Str(loc_Renamed) & vbCrLf & "Press OK to start learning.", MsgBoxStyle.Information + MsgBoxStyle.MsgBoxSetForeground)
        LearnIREnd()
        'TODO: callback.LearnIRUpdate(loc, True)
        'Beep()
        gbusy = 0

    End Sub


    'Public Sub SetPoint(ByRef unit As Short, ByRef point As Short, ByRef Value As Short)
    '    'QExec("", "", "", SET_RELAY, Value, unit, point)
    'End Sub

    ' get realtime I/O status and read in buffer
    Public Function PollIO() As Boolean
        'Dim snd(7) As Byte
        Dim i As Short
        Dim count As Short
        Dim Y As Short
        Dim crc As Integer
        Dim rcrc As Integer
        Dim retry_counter As Short
        Dim debugstr As String = ""
        On Error Resume Next

        If Not gIOEnabled Then
            PollIO = False
            Exit Function
        End If

        If gbusy = 1 Then Exit Function


retry:
        ' set busy flag so comm does not treat any data bytes as notifications
        gbusy = 1
        'snd(0) = 42
        'snd(1) = 0
        'snd(2) = 0
        'snd(3) = 141
        'snd(4) = 181
        'snd(5) = 0
        'snd(6) = 37
        'snd(7) = 233
        'SendToCPUXA(snd)

        SendToCPUXA(FrameGenerator.GetRealtimeIO)

        If gLogComms Then
            WriteLog("Waiting for IO info from Ocelot")
        End If
        If (WaitForChar(gTimeOut, 264, 42)) Then
            If gLogComms Then
                For i = 0 To (255 + 6)
                    debugstr = debugstr & rbuffer(i).ToString("x2").ToUpper & "-"
                Next
                WriteLog(debugstr)
            End If
            crc = CalcCRC(rbuffer, 262)
            rcrc = char_to_int(CShort(rbuffer(263)), CShort(rbuffer(262)))
            If crc = rcrc Then
                If rbuffer(3) = 141 And rbuffer(4) = 181 Then
                    For i = 0 To 255
                        gIO(i) = rbuffer(i + 6)
                    Next
                Else
                    If gLogComms Then
                        WriteLog("Warning, Ocelot returned wrong buffer for IO buffer request, ignoring data")
                    End If
                End If
                gCommCRCCount = 0
            Else
                gCommCRCCount = gCommCRCCount + 1
                If gLogErrors And gCommCRCCount > 1 Then
                    hs.WriteLog(IFACE_OTHER, "Warning, Ocelot Plug-in, Checksum error reading io points: " & Hex(crc) & " " & Hex(rcrc))
                    gCommCRCCount = 0
                End If
                PollIO = False
                gbusy = 0
                GoTo done
            End If
            'Debug.Print "RCV: " + Str(rcv(0)) + " " + Str(rcv(1))
            'Debug.Print "IO: " + Str(gIO(0)) + " " + Str(gIO(1))
            gCommTOCount = 0
            Status = ERR_NONE
        Else
            retry_counter = retry_counter + 1
            If retry_counter = 2 Then
                'Debug.Print "POLL IO TIMEOUT"
                gCommTOCount = gCommTOCount + 1
                If gLogErrors And gCommTOCount > 1 Then
                    hs.WriteLog(IFACE_OTHER, "Warning, Ocelot Plug-in, Timeout getting realtime I/O")
                    gCommTOCount = 0
                End If
            Else
                GoTo retry
            End If
            Status = ERR_SEND
        End If
        gbusy = 0

        If Status = ERR_NONE Then
            PollIO = True
        Else
            PollIO = False
        End If
done:

    End Function

    Public Function GetVar(ByRef varnum As Short) As Integer
        GetVar = gVars(varnum)
    End Function

    Public Function GetVarsFromHW() As Boolean
        'Dim snd(7) As Byte
        Dim i As Short
        Dim count As Short
        Dim Y As Short
        Dim crc As Integer
        Dim rcrc As Integer
        Dim t As Integer
        Dim retry_counter As Short
        Dim debugstr As String = ""
        On Error Resume Next

        If gbusy = 1 Then Return False

retry:
        gbusy = 1

        ' changed request for cmax 1.70 (not sure if older executives will still work)
        'snd(0) = 42
        'snd(1) = 0
        'snd(2) = 0
        'snd(3) = gLoVar
        'snd(4) = gHiVar
        'snd(5) = 0
        'Err.Clear()
        't = CalcCRC(snd, 6)
        ''If Err.Number <> 0 Then
        ''hs.WriteLog("Error", "Ocelot, missing DLL ahutil.dll")
        ''End If
        'snd(7) = CByte(t And &HFFS)
        'snd(6) = CByte(Int(t / 256))
        'SendToCPUXA(snd)
        SendToCPUXA(FrameGenerator.UNDOCUMENTED_GetMemoryData(gLoVar, gHiVar))

        If gLogComms Then
            WriteLog("Waiting for var info from Ocelot")
        End If
        If (WaitForChar(gTimeOut, 264, 42)) Then
            If gLogComms Then
                For Y = 0 To (127 + 6)
                    debugstr = debugstr & Hex(rbuffer(Y)) & " "
                Next
                WriteLog(debugstr)
            End If
            crc = CalcCRC(rbuffer, 262)
            rcrc = char_to_int(CShort(rbuffer(263)), CShort(rbuffer(262)))
            If crc = rcrc Then
                If (rbuffer(3) = gLoVar) And (rbuffer(4) = gHiVar) Then
                    For Y = 0 To 127
                        t = rbuffer(Y * 2 + 7)
                        t = t * 256
                        t = t + rbuffer(Y * 2 + 6)
                        gVars(Y) = t
                    Next Y
                Else
                    If gLogComms Then
                        WriteLog("Warning, Ocelot returned wrong buffer for variable buffer request, ignoring data")
                    End If
                End If
                gCommCRCCount = 0
                GetVarsFromHW = True
            Else
                gCommCRCCount = gCommCRCCount + 1
                If gLogErrors And gCommCRCCount > 1 Then
                    hs.WriteLog(IFACE_OTHER, "Error, Ocelot Plug-in, Checksum error reading vars")
                    'hs.WriteLog "DEBUG", "Expected: 2A 00 00"
                    'hs.WriteLog "DEBUG", "Got: " + Hex(rbuffer(0)) + " " + Hex(rbuffer(1)) + " " + Hex(rbuffer(2))
                    gCommCRCCount = 0
                End If
                GetVarsFromHW = False
            End If
            gbusy = 0
            gCommTOCount = 0
            GoTo done
        Else
            retry_counter = retry_counter + 1
            If retry_counter = 2 Then
                'Debug.Print "POLL IO TIMEOUT"
                gCommTOCount = gCommTOCount + 1
                If gLogErrors And gCommTOCount > 1 Then
                    hs.WriteLog(IFACE_OTHER, "Warning, Ocelot Plug-in, Timeout getting Vars")
                    gCommTOCount = 0
                End If
            Else
                GoTo retry
            End If
            Status = ERR_SEND
        End If
        gbusy = 0
        GetVarsFromHW = False
done:

        Exit Function
error1:
        gbusy = 0
        hs.WriteLog(IFACE_OTHER, "Error, Ocelot Plug-in, In PollVars: " & Err.Description)
        GetVarsFromHW = False
    End Function

    ' tell unit to get analog input values
    Public Function StartGetAnalog(ByRef unit As Byte, ByRef point As Byte) As Short
        'Dim snd(7) As Byte
        Dim i As Short
        Dim rcv(6) As Byte
        Dim buf(256) As Byte
        On Error Resume Next

        If Not gIOEnabled Then Exit Function

        gbusy = 1
        'snd(0) = 200
        'snd(1) = 31
        'snd(2) = point + 10
        'snd(3) = 0
        'snd(4) = 0
        'snd(5) = 0
        'snd(6) = 0
        'snd(7) = CheckSum(snd)
        'SendToCPUXA(snd)
        SendToCPUXA(FrameGenerator.InitiateGetUnitParameters(point + 10))
        WaitForChar(gTimeOut, 3, 0)

        gbusy = 0
    End Function

    Public Function GetAnalogInputValue(ByRef unit As Short, ByRef point As Short) As Short
        'Dim snd(7) As Byte
        Dim i As Short
        Dim rcv(6) As Byte
        Dim buf(256) As Byte
        On Error Resume Next

        If Not gIOEnabled Then Exit Function

        gbusy = 1

        ' get the parameter
        'snd(0) = 42
        'snd(1) = 0
        'snd(2) = 0
        'snd(3) = 139
        'snd(4) = 180
        'snd(5) = 0
        'snd(6) = 164
        'snd(7) = 120
        'SendToCPUXA(snd)
        SendToCPUXA(FrameGenerator.GetUnitParameters)

        If (WaitForChar(gTimeOut, 6 + 256 + 2, 42)) Then
            For i = 0 To 5
                rcv(i) = rbuffer(i)
            Next
            For i = 0 To 255
                buf(i) = rbuffer(i + 6)
            Next
            'Debug.Print "RCV: " + str(rcv(0)) + " " + str(rcv(1))
            'Debug.Print "IO: " + str(gIO(0)) + " " + str(gIO(1))
            Status = ERR_NONE
            GetAnalogInputValue = buf((unit - 1) * 2)
        Else
            If gLogErrors Then
                hs.WriteLog(IFACE_OTHER, "Warning, Ocelot Plug-in, Timeout getting analog I/O")
            End If
            Status = ERR_SEND
            GetAnalogInputValue = -1
        End If
        gbusy = 0
    End Function

    Public Function IsBusy() As Boolean
        If gbusy = 1 Then
            Return True
        Else
            Return False
        End If
    End Function

    Public Function GetAnalogInput(ByRef unit As Byte, ByRef point As Byte) As Short
        'Dim snd(7) As Byte
        Dim i As Short
        Dim rcv(6) As Byte
        Dim buf(256) As Byte
        On Error Resume Next

        If Not gIOEnabled Then Exit Function

        gbusy = 1

        'snd(0) = 200
        'snd(1) = 31
        'snd(2) = point + 10
        'snd(3) = 0
        'snd(4) = 0
        'snd(5) = 0
        'snd(6) = 0
        'snd(7) = CheckSum(snd)
        'SendToCPUXA(snd)
        SendToCPUXA(FrameGenerator.InitiateGetUnitParameters(point + 10))
        WaitForChar(gTimeOut, 3, 0)
        ' per app dig, wait 3 seconds before data is available
        ' seems to work ok with 1
        Wait(1)

        gbusy = 1
        ' get the parameter
        'snd(0) = 42
        'snd(1) = 0
        'snd(2) = 0
        'snd(3) = 139
        'snd(4) = 180
        'snd(5) = 0
        'snd(6) = 164
        'snd(7) = 120
        'SendToCPUXA(snd)
        SendToCPUXA(FrameGenerator.GetUnitParameters)

        If (WaitForChar(gTimeOut, 6 + 256 + 2, 42)) Then
            For i = 0 To 5
                rcv(i) = rbuffer(i)
            Next
            For i = 0 To 255
                buf(i) = rbuffer(i + 6)
            Next
            'Debug.Print "RCV: " + str(rcv(0)) + " " + str(rcv(1))
            'Debug.Print "IO: " + str(gIO(0)) + " " + str(gIO(1))
            Status = ERR_NONE
        Else
            If gLogErrors Then
                hs.WriteLog(IFACE_OTHER, "Warning, Ocelot Plug-in, Timeout getting analog I/O")
            End If
            Status = ERR_SEND

        End If
        gbusy = 0
        GetAnalogInput = buf((unit - 1) * 2)
    End Function

    Private Sub SetInternalPoint(ByVal unit As Short, ByVal point As Short, ByVal Value As Short)
        Dim u As Short
        On Error Resume Next

        u = unit
        If u > 128 Then Exit Sub

        u = u * 2
        u = u - 2

        'lo = gIO(u)
        'hi = gIO(u + 1)


        If point < 8 Then
            WriteLog("SetInternalPoint, unit " & unit.ToString & ", point " & point.ToString & ", to Value " & Value.ToString & " converts to index " & u.ToString & " bit value " & Hex(2 ^ point))
            If Value = 1 Then
                gIO(u) = gIO(u) Or (2 ^ point)
            Else
                gIO(u) = gIO(u) And Not (2 ^ point)
            End If
        Else
            WriteLog("SetInternalPoint, unit " & unit.ToString & ", point " & point.ToString & ", to Value " & Value.ToString & " converts to index " & (u + 1).ToString & " bit value " & Hex(2 ^ (point - 8)))
            If Value = 1 Then
                gIO(u + 1) = gIO(u + 1) Or (2 ^ (point - 8))
            Else
                gIO(u + 1) = gIO(u + 1) And Not (2 ^ (point - 8))
            End If
        End If

        'Select Case point
        '    Case 0
        '        If Value = 1 Then
        '            lo = lo Or &H1S
        '        Else
        '            lo = lo And Not &H1S
        '        End If
        '    Case 1
        '        If Value = 1 Then
        '            lo = lo Or &H2S
        '        Else
        '            lo = lo And Not &H2S
        '        End If
        '    Case 2
        '        If Value = 1 Then
        '            lo = lo Or &H4S
        '        Else
        '            lo = lo And Not &H4S
        '        End If
        '    Case 3
        '        If Value = 1 Then
        '            lo = lo Or &H8S
        '        Else
        '            lo = lo And Not &H8S
        '        End If
        '    Case 4
        '        If Value = 1 Then
        '            lo = lo Or &H10S
        '        Else
        '            lo = lo And Not &H10S
        '        End If
        '    Case 5
        '        If Value = 1 Then
        '            lo = lo Or &H20S
        '        Else
        '            lo = lo And Not &H20S
        '        End If
        '    Case 6
        '        If Value = 1 Then
        '            lo = lo Or &H40S
        '        Else
        '            lo = lo And Not &H40S
        '        End If
        '    Case 7
        '        If Value = 1 Then
        '            lo = lo Or &H80S
        '        Else
        '            lo = lo And Not &H80S
        '        End If
        '    Case 8
        '        If Value = 1 Then
        '            hi = hi Or &H1S
        '        Else
        '            hi = hi And Not &H1S
        '        End If
        '    Case 9
        '        If Value = 1 Then
        '            hi = hi Or &H2S
        '        Else
        '            hi = hi And Not &H2S
        '        End If
        '    Case 10
        '        If Value = 1 Then
        '            hi = hi Or &H4S
        '        Else
        '            hi = hi And Not &H4S
        '        End If
        '    Case 11
        '        If Value = 1 Then
        '            hi = hi Or &H8S
        '        Else
        '            hi = hi And Not &H8S
        '        End If
        '    Case 12
        '        If Value = 1 Then
        '            hi = hi Or &H10S
        '        Else
        '            hi = hi And Not &H10S
        '        End If
        '    Case 13
        '        If Value = 1 Then
        '            hi = hi Or &H20S
        '        Else
        '            hi = hi And Not &H20S
        '        End If
        '    Case 14
        '        If Value = 1 Then
        '            hi = hi Or &H40S
        '        Else
        '            hi = hi And Not &H40S
        '        End If
        '    Case 15
        '        If Value = 1 Then
        '            hi = hi Or &H80S
        '        Else
        '            hi = hi And Not &H80S
        '        End If
        'End Select


        'gIO(u) = lo
        'gIO(u + 1) = hi


    End Sub

    Public Function GetPointValue(ByVal unit As Short, ByVal point As Short) As Short
        Dim j As Short
        Dim u As Short
        On Error Resume Next

        u = unit
        If u > 128 Then Exit Function

        u = u * 2
        u = u - 2

        'lo = gIO(u)
        'hi = gIO(u + 1)

        If point < 8 Then
            j = gIO(u) And (2 ^ point)
        Else
            j = gIO(u + 1) And (2 ^ (point - 8))
        End If

        If j = 0 Then
            j = 1
        Else
            j = 0
        End If

        Return j

        'Select Case point
        '    Case 0
        '        j = lo And &H1
        '    Case 1
        '        j = lo And &H2
        '    Case 2
        '        j = lo And &H4
        '    Case 3
        '        j = lo And &H8
        '    Case 4
        '        j = lo And &H10
        '    Case 5
        '        j = lo And &H20
        '    Case 6
        '        j = lo And &H40
        '    Case 7
        '        j = lo And &H80
        '    Case 8
        '        j = hi And &H1
        '    Case 9
        '        j = hi And &H2
        '    Case 10
        '        j = hi And &H4
        '    Case 11
        '        j = hi And &H8
        '    Case 12
        '        j = hi And &H10
        '    Case 13
        '        j = hi And &H20
        '    Case 14
        '        j = hi And &H40
        '    Case 15
        '        j = hi And &H80
        'End Select

        'If j = 0 Then
        '    j = 1
        'Else
        '    j = 0
        'End If

        'GetPoint = j

    End Function

    Private Sub DoSetPoint(ByRef unit As Byte, ByRef point As Byte, ByRef Value As Byte)
        Dim rcv(6) As Byte
        On Error Resume Next

        point -= 1

        gbusy = 1

        Dim bte() As Byte = FrameGenerator.WriteRelayOutput(unit, point, Value)
        WriteLog("DoSetPoint for Unit " & unit.ToString & ", Point " & point.ToString & ", to Value " & Value.ToString & " = " & FrameGenerator.Debug(bte))

        SendToCPUXA(bte)

        WaitForChar(gTimeOut, 3, 0)
        gbusy = 0


    End Sub

    Public Function GetParameter(ByRef unit As Byte, ByRef pnum As Byte) As Integer
        'Dim snd(7) As Byte
        Dim i As Short
        Dim rcv(6) As Byte
        Dim buf(256) As Byte
        Dim count As Short
        Dim lo As Integer
        Dim hi As Integer
        Dim lpval As Integer


        On Error Resume Next
        gbusy = 1

        count = 3
start:
        count = count - 1
        If count = 0 Then GoTo timeout

        'snd(0) = 200
        'snd(1) = 31
        'snd(2) = pnum - 1
        'snd(3) = 0
        'snd(4) = 0
        'snd(5) = 0
        'snd(6) = 0
        'snd(7) = CheckSum(snd)
        'SendToCPUXA(snd)
        SendToCPUXA(FrameGenerator.InitiateGetUnitParameters(pnum - 1))

        WaitForChar(gTimeOut, 3, 0)

        ' per app dig, wait 3 seconds before data is available
        ' seems to work ok with 1
        Wait(3)

        ' get the parameter
        'snd(0) = 42
        'snd(1) = 0
        'snd(2) = 0
        'snd(3) = 139
        'snd(4) = 180
        'snd(5) = 0
        'snd(6) = 164
        'snd(7) = 120
        'SendToCPUXA(snd)
        SendToCPUXA(FrameGenerator.GetUnitParameters)


        If (WaitForChar(gTimeOut, 6 + 256 + 2, 42)) Then
            For i = 0 To 5
                rcv(i) = rbuffer(i)
            Next
            For i = 0 To 255
                buf(i) = rbuffer(i + 6)
            Next
            'Debug.Print "RCV: " + str(buf(0)) + " " + str(buf(1))
            Status = ERR_NONE
        Else
            hs.WriteLog(IFACE_OTHER, "Warning, Ocelot Plug-in, Timeout getting parameter, retrying ...")
            Status = ERR_SEND
            GoTo start
        End If
        gbusy = 0
        lo = buf((unit - 1) * 2)
        hi = buf(((unit - 1) * 2) + 1)
        lpval = (hi * 256) + lo


        GetParameter = lpval
        Exit Function
timeout:
        gbusy = 0
        hs.WriteLog(IFACE_OTHER, "Error, Ocelot Plug-in, Timeout getting parameter, retried 3 times")
        GetParameter = -1
    End Function

    Public Function InitCPUXA() As String
        Dim st As String
        'Dim irsize As Integer

        Try
            'FrmDebug.Show
            If gmoduleOK Then
                ' already init, can use for IR or IO
                Return ""
            End If

            hs.WriteLog(IFACE_OTHER, "Ocelot Plug-in, Version " & My.Application.Info.Version.Major & "." & My.Application.Info.Version.Minor & "." & My.Application.Info.Version.Revision)

            Try
                ConfigPage = New ConfigWebPag("ocelot_config")
                hs.RegisterPage("ocelot_config", IFACE_NAME, "")
            Catch ex As Exception
                Return "Error registering web pages with the HomeSeer server."
            End Try

            Try
                ' register this page as a link in the HomeSeer setup/Interfaces page
                Dim wpd As New WebPageDesc
                wpd.link = "ocelot_config"                ' we add the instance so it goes to the proper plugin instance when selected
                wpd.linktext = "Ocelot Plugin Config"
                wpd.page_title = "Ocelot Plugin Config"
                wpd.plugInName = IFACE_NAME
                wpd.plugInInstance = Instance
                callback.RegisterConfigLink(wpd)
            Catch ex As Exception
                Return "Error registering configuration web page."
            End Try


            gTimeOut = 4
            gReverse = New Object() {12, 13, 14, 15, 2, 3, 0, 1, 4, 5, 6, 7, 10, 11, 8, 9}
            irtable = New Object() {&H5S, &H3S, &H1BS, &H2BS, &HBS, &H13S, &H3BS, &HDS, &H15S, &H2DS, &H1DS, &H11S, &H19S, &H29S, &H39S, &H1S, &H9S, &HFS, &H10S, &H4S, &H1CS, &HCS, &H1FS, &H21S, &H24S, &H2CS, &H2S, &H1AS, &H3AS, &H32S, &HAS, &H33S, &H23S, &H25S, &H35S, &H31S}
            irdevtable = New Object() {&H1ES, &H17S, &HES, &H6S, &H16S, &H34S, &H2FS, &H3FS}

            If MSComm1 Is Nothing Then
                MSComm1 = New System.IO.Ports.SerialPort 'rs232
            Else
                Try
                    MSComm1.Close()
                Catch ex As Exception
                End Try
            End If
            hs.WriteLog(IFACE_OTHER, "Initializing Ocelot on port " & gComPort)

            Dim ports() As String = IO.Ports.SerialPort.GetPortNames
            Dim FoundPort As Boolean = False
            If ports IsNot Nothing AndAlso ports.Length > 0 Then
                For Each p As String In ports
                    If p.Trim.ToUpper = gComPort.Trim.ToUpper Then
                        gComPort = p
                        FoundPort = True
                        Exit For
                    End If
                Next
            End If

            If FoundPort Then
                MSComm1.PortName = gComPort
                MSComm1.BaudRate = 9600
                MSComm1.Parity = IO.Ports.Parity.None
                MSComm1.DataBits = 8
                MSComm1.StopBits = IO.Ports.StopBits.One
                MSComm1.ReceivedBytesThreshold = 1
                MSComm1.DiscardNull = False

                MSComm1.Open()
            Else
                Dim sErr As String = "Error, port " & gComPort & " was not found in the system.  Please check your port setting on the Interface Management page and try again."
                hs.WriteLog(IFACE_OTHER, sErr)
                Return sErr
            End If

            'Wait 2
            gmoduleOK = True
            InitCPUXA = GetRev()
            If InitCPUXA <> "" Then
                Status = ERR_INIT
                Exit Function
            End If
            If gIRSize = 0 Then
                ' bad read from the Ocelot try again
                InitCPUXA = GetRev()
            End If

            SetAutoIO(False)
            SetAutoX10(False)
            SetAutoIR(False)
            DisableRescan()

            gConfigPollVars = Convert.ToBoolean(GetSettingLocal("hspi_ocelot", "Settings", "gConfigPollVars", True))

            ' initial polls
            If gIOEnabled Then
                PollVars()
                PollInputs()
            End If

            ' start polling thread
            PollThread = New Thread(AddressOf PollThreadProc)
            PollThread.Name = "Ocelot Polling"
            PollThread.Start()
            Exit Function
        Catch ex As Exception
            Status = ERR_INIT
            st = "Ocelot init error: " & Err.Description
            hs.WriteLog(IFACE_OTHER, "Error, Ocelot Plug-in, " & st)
            Return st
        End Try

        Return ""

    End Function

    Private Sub PollThreadProc()
        'Dim stopWatch As New Stopwatch
        Dim LastQPoll As Date = Date.MinValue
        Dim LastInputPoll As Date = Date.MinValue
        Dim DidOne As Boolean = False

        Try
            Try
                'stopWatch.Start()
            Catch ex As Exception
                'hs.WriteLog("Ocelot", "Unable to start poll timer: " & ex.Message)
            End Try

            Do
                DidOne = False

                If LastQPoll = Date.MinValue OrElse Now.Subtract(LastQPoll).TotalMilliseconds > 200 Then
                    If Not IsBusy() Then
                        GotRcvData = False
                        PollQueues()
                        LastQPoll = Now
                        DidOne = True
                    End If
                End If
                'If (Stopwatch.ElapsedMilliseconds > 200 And ((Stopwatch.ElapsedMilliseconds Mod 200) = 0)) Or GotRcvData Then
                '    GotRcvData = False
                '    PollQueues()
                'End If
                If LastInputPoll = Date.MinValue OrElse Now.Subtract(LastInputPoll).TotalMilliseconds >= gPollInterval Then
                    If gIOEnabled AndAlso Not IsBusy() Then
                        PollInputs()
                        If gConfigPollVars Then PollVars()
                        LastInputPoll = Now
                        DidOne = True
                    End If
                End If
                'If Stopwatch.ElapsedMilliseconds >= gPollInterval Then
                '    If gIOEnabled Then
                '        PollInputs()
                '        If gConfigPollVars Then
                '            PollVars()
                '        End If
                '    End If
                '    Try
                '        Stopwatch.Reset()
                '        Stopwatch.Start()
                '    Catch ex As Exception
                '    End Try
                'End If
                If DidOne Then
                    Thread.Sleep(10)
                End If
            Loop
        Catch ex As Exception
            'hs.WriteLog(IFACE_OTHER, "Polling thread shutdown")
        End Try
    End Sub

    Private Sub SetAutoIO(ByRef Enabled As Boolean)
        'Dim snd(7) As Byte
        On Error Resume Next

        gbusy = 1
        ' set parameters to get I/O change notifications
        'snd(0) = 200
        'snd(1) = 40
        'snd(2) = 16
        'snd(3) = 0
        'If Enabled Then
        '    snd(4) = 2 ' 1=get FF if I/O point has changed, 2=get F0-mod#-point#-new_state
        'Else
        '    snd(4) = 0
        'End If
        'snd(5) = 0
        'snd(6) = 0
        'snd(7) = CheckSum(snd)
        'SendToCPUXA(snd)

        Dim bte() As Byte = FrameGenerator.WriteCPUXAParameterData(16, IIf(Enabled, 2, 0))
        WriteLog("SetAutoIO to " & IIf(Enabled, "Enabled", "Disabled") & " = " & FrameGenerator.Debug(bte))

        SendToCPUXA(bte)

        WaitForChar(gTimeOut, 3, 0)

        gbusy = 0
    End Sub

    Public Sub EnableTouch(ByRef mode As UInt16)
        'Dim snd(7) As Byte
        On Error Resume Next

        gbusy = 1
        ' set parameters to get I/O change notifications
        'snd(0) = 200
        'snd(1) = 40
        'snd(2) = 6
        'snd(3) = 0
        'snd(4) = mode
        'snd(5) = 0
        'snd(6) = 0
        'snd(7) = CheckSum(snd)
        'SendToCPUXA(snd)
        SendToCPUXA(FrameGenerator.WriteCPUXAParameterData(6, mode))
        WaitForChar(gTimeOut, 3, 0)

        gbusy = 0
    End Sub

    ' Appdig recommends disabling this as it could cause false io point events
    Private Sub DisableRescan()
        'Dim snd(7) As Byte
        On Error Resume Next

        ' set parameters to get X10 notifications
        gbusy = 1
        If gmoduleOK Then
            'snd(0) = 200
            'snd(1) = 40
            'snd(2) = 7 ' param 7
            'snd(3) = 0 ' set data to 0
            'snd(4) = 0 ' set data to 0
            'snd(5) = 0
            'snd(6) = 0
            'snd(7) = CheckSum(snd)
            'SendToCPUXA(snd)
            SendToCPUXA(FrameGenerator.WriteCPUXAParameterData(7, 0))
            WaitForChar(4, 3, 0)
        End If
        gbusy = 0
    End Sub

    Private Sub SetAutoX10(ByRef Enabled As Boolean)
        'Dim snd(7) As Byte
        On Error Resume Next

        ' set parameters to get X10 notifications
        gbusy = 1
        If gmoduleOK Then
            'snd(0) = 200
            'snd(1) = 40
            'snd(2) = 15
            'snd(3) = 0
            'If Enabled = True Then
            '    snd(4) = 1
            'Else
            '    snd(4) = 0
            'End If
            'snd(5) = 0
            'snd(6) = 0
            'snd(7) = CheckSum(snd)
            'SendToCPUXA(snd)
            SendToCPUXA(FrameGenerator.WriteCPUXAParameterData(15, IIf(Enabled, 1, 0)))

            WaitForChar(4, 3, 0)
        End If
        gbusy = 0
    End Sub

    ' enable cpuxa to send ir match notifications
    Private Sub SetAutoIR(ByRef Enabled As Boolean)
        'Dim snd(7) As Byte
        On Error Resume Next

        gbusy = 1
        If gmoduleOK Then
            'snd(0) = 200
            'snd(1) = 40
            'snd(2) = 17
            'snd(3) = 0
            'If Enabled Then
            '    snd(4) = 1
            'Else
            '    snd(4) = 0
            'End If
            'snd(5) = 0
            'snd(6) = 0
            'snd(7) = CheckSum(snd)
            'SendToCPUXA(snd)
            SendToCPUXA(FrameGenerator.WriteCPUXAParameterData(17, IIf(Enabled, 1, 0)))

            WaitForChar(gTimeOut, 3, 0)

            ' ascii reporting of IR is disabled
            'snd(0) = 200
            'snd(1) = 40
            'snd(2) = 18
            'snd(3) = 0
            'snd(4) = 0
            'snd(5) = 0
            'snd(6) = 0
            'snd(7) = CheckSum(snd)
            'SendToCPUXA(snd)
            SendToCPUXA(FrameGenerator.WriteCPUXAParameterData(18, 0))

            WaitForChar(gTimeOut, 3, 0)
        End If
        gbusy = 0
    End Sub

    Public Sub PlayIR(ByRef loc_Renamed As Integer)
        'Debug.Print("PlayIR: " & Str(loc_Renamed))
        'hs.WriteLog "DEBUG", "queueing PLAY IR: " + Str(loc) + " Zone: " + Str(gSetZone)
        'TODO: QExec(0, "", "", SEND_IR, CShort(loc_Renamed), gSetZone, 0)
    End Sub

    Private Sub DoPlayIR(ByRef IRLocation As UInt16)
        'Debug.Print("DoPlayIR: " & Str(loc_Renamed))
        'Dim snd(7) As Byte
        On Error Resume Next

        'hs.WriteLog "DEBUG", "PlayIR-Sending IR to unit: " & "location: " & Str(loc)

        gbusy = 1
        If gmoduleOK Then
            'snd(0) = 200
            'snd(1) = 90
            'snd(2) = IRLocation And &HFFS
            'If IRLocation <> 0 Then
            '    snd(3) = Int(IRLocation / 256)
            'Else
            '    snd(3) = 0
            'End If
            'snd(4) = 0
            'snd(5) = 0
            'snd(6) = 0
            'snd(7) = CheckSum(snd)
            'SendToCPUXA(snd)
            SendToCPUXA(FrameGenerator.SendResidentIR(IRLocation))

            WaitForChar(gTimeOut, 3, 0)
        End If
        gbusy = 0
    End Sub

    Public Sub PlayRemoteIR(ByRef addr As Short, ByRef zone As Short, ByRef loc_Renamed As Short)
        'hs.WriteLog "DEBUG", "PlayRemoteIR " & "Addr: " & Str(addr) & " Zone: " & Str(zone) & " Loc: " & Str(loc)
        'QExec(0, "", "", SEND_REMOTE_IR, addr, zone, CShort(loc_Renamed))
        GotRcvData = True   ' force quick handling of queues
    End Sub

    'UPGRADE_NOTE: loc was upgraded to loc_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Private Sub DoPlayRemoteIR(ByRef addr As Byte, ByRef zone As Byte, ByRef IRLocation As UInt16)
        'Dim snd(7) As Byte
        On Error Resume Next

        'hs.WriteLog "DEBUG", "RemoteIR-Sending IR to unit: " & Str(addr) & " Zone: " & Str(zone) & " Loc: " & Str(loc)

        gbusy = 1
        If gmoduleOK Then
            'snd(0) = 200
            'snd(1) = 92
            'snd(2) = addr
            'snd(3) = zone
            'snd(4) = IRLocation And &HFFS
            'If IRLocation <> 0 Then
            '    If IRLocation > 255 Then
            '        snd(5) = Int(IRLocation / 256)
            '    Else
            '        snd(5) = 0
            '    End If
            'Else
            '    snd(5) = 0
            'End If
            'snd(6) = 0
            'snd(7) = CheckSum(snd)
            'SendToCPUXA(snd)
            SendToCPUXA(FrameGenerator.SendRemoteIRCommand(addr, zone, IRLocation))

            WaitForChar(gTimeOut, 3, 0)
        End If
        gbusy = 0
    End Sub

    ' learn an ir key into given location
    Public Sub LearnIRStart(ByRef IRLocation As UInt16)
        'Dim snd(7) As Byte
        On Error Resume Next

        If gmoduleOK Then
            'hs.WriteLog "DEBUG", "LEARN: " + Str(loc)
            'Debug.Print "LEARN: " + str(loc)

            'snd(0) = 200
            'snd(1) = 91
            'snd(2) = IRLocation And &HFFS
            'If IRLocation <> 0 Then
            '    snd(3) = Int(IRLocation / 256)
            'Else
            '    snd(3) = 0
            'End If
            ''Debug.Print "LSB: " + str(snd(2)) + " MSB: " + str(snd(3))
            'snd(4) = 120
            'snd(5) = 0
            'snd(6) = 0
            'snd(7) = CheckSum(snd)
            'SendToCPUXA(snd)
            SendToCPUXA(FrameGenerator.InitiateIRLearn(IRLocation, 120))

            WaitForChar(gTimeOut, 3, 0)
        End If

    End Sub

    Public Sub LearnIREnd()
        On Error Resume Next

        If gmoduleOK Then
            gLearnWait = True
            WaitForChar(15, 3, 0)
            gLearnWait = False
        End If
    End Sub

    ' get the firmware revision of cpuxa
    Private Function GetRev() As String
        'Dim snd(7) As Byte
        Dim rcv(6) As Byte
        Dim data(256) As Byte
        Dim i As Short
        Dim lo As Integer
        Dim hi As Integer
        Dim lpval As Integer
        Dim retry_count As Short

        On Error Resume Next

        GetRev = ""

        gbusy = 1

        ' get versions

        'snd(0) = 42
        'snd(1) = 0
        'snd(2) = 0
        'snd(3) = 143
        'snd(4) = 183
        'snd(5) = 0
        'snd(6) = 45
        'snd(7) = 235

again:
        'SendToCPUXA(snd)
        SendToCPUXA(FrameGenerator.UNDOCUMENTED_GetMemoryPointer)

        If (WaitForChar(2, 6 + 256 + 2, 42)) Then
            For i = 0 To 5
                rcv(i) = rbuffer(i)
            Next
            For i = 0 To 255
                data(i) = rbuffer(i + 6)
            Next


            ' save variables base address
            gLoVar = data(12)
            gHiVar = data(13)

            'Debug.Print "ver " + str(data(6)) + str(data(7)) + str(data(8)) + str(data(9)) + str(data(14))
            'fMainForm.WriteMon "Info", "CPU-XA Firmware version: " + str(data(6)) + str(data(7)) + str(data(8)) + str(data(9))
            hs.WriteLog(IFACE_OTHER, "Ocelot Plug-in, Found CPU-XA/Ocelot")

            ' get ir size
            ' this sequence is sent by CMAX to get controller parameters
            'snd(0) = &H2AS
            'snd(1) = 0
            'snd(2) = 0
            'snd(3) = 0
            'snd(4) = &H7FS
            'snd(5) = 0
            'snd(6) = &HA5S
            'snd(7) = &H7DS

            'snd(0) = 42
            'snd(1) = 0
            'snd(2) = 0
            'snd(3) = 0
            'snd(4) = 127
            'snd(5) = 0
            'snd(6) = 165
            'snd(7) = 125
            'SendToCPUXA(snd)
            SendToCPUXA(FrameGenerator.UNDOCUMENTED_GetIRSize)

            If (WaitForChar(3, 6 + 256 + 2, 42)) Then
                For i = 0 To 5
                    rcv(i) = rbuffer(i)
                Next
                For i = 0 To 255
                    data(i) = rbuffer(i + 6)
                Next

                lo = data(122 * 2)
                hi = data((122 * 2) + 1)
                lpval = (hi * 256) + lo
                gIRSize = lpval
                hs.WriteLog(IFACE_OTHER, "Ocelot Infrared locations size: " & gIRSize.ToString)
                Status = ERR_NONE
            Else
                Status = ERR_SEND
                hs.WriteLog(IFACE_OTHER, "Error, Ocelot Plug-in, Timeout getting controller parameters")
                GetRev = "Timeout from CPU-XA/Ocelot"
                retry_count = retry_count + 1
                If retry_count < 2 Then GoTo again
            End If
        Else
            Status = ERR_SEND
            hs.WriteLog(IFACE_OTHER, "Error, Ocelot Plug-in, Timeout looking for CPU-XA/Ocelot")
            GetRev = "Timeout from CPU-XA/Ocelot"
            retry_count = retry_count + 1
            If retry_count < 2 Then GoTo again
        End If

        'Text1 = Text1 + "Get RS485 units" + vbCrLf

        ' get units attached to RS485
        'snd(0) = 42
        'snd(1) = 0
        'snd(2) = 0
        'snd(3) = 1
        'snd(4) = 176
        'snd(5) = 0
        'snd(6) = 148
        'snd(7) = 39
        'SendToCPUXA(snd)
        SendToCPUXA(FrameGenerator.GetUnitType)

        Wait(2)
        If (WaitForChar(gTimeOut, 6 + 256 + 2, 42)) Then
            For i = 0 To 5
                rcv(i) = rbuffer(i)
            Next
            For i = 0 To 127
                gUnits(i) = rbuffer(i + 6)
            Next

            For i = 0 To 127
                Select Case gUnits(i)
                    Case CPUXA_UnitTypes.Secu16IR '6
                        hs.WriteLog(IFACE_OTHER, "Ocelot Plug-in startup, Unit found: SECU16-IR Addr: " & (i + 1).ToString)
                        gSecu16IR_present = True
                        gSecu16IR_Unit = i + 1
                        If gLogComms Then
                            WriteLog("UNIT SECU16-IR")
                        End If
                    Case CPUXA_UnitTypes.Secu16 '11
                        hs.WriteLog(IFACE_OTHER, "Ocelot Plug-in startup, Unit found: SECU16 Addr: " & (i + 1).ToString)
                        If gLogComms Then
                            WriteLog("UNIT SECU16")
                        End If
                    Case CPUXA_UnitTypes.Secu16i '12
                        hs.WriteLog(IFACE_OTHER, "Ocelot Plug-in startup, Unit found: SECU16I Addr: " & (i + 1).ToString)
                        If gLogComms Then
                            WriteLog("UNIT SECU16I")
                        End If
                    Case CPUXA_UnitTypes.Rly08XA '13
                        hs.WriteLog(IFACE_OTHER, "Ocelot Plug-in startup, Unit found: RLY08-XA Addr: " & (i + 1).ToString)
                        If gLogComms Then
                            WriteLog("UNIT RLY08")
                        End If
                End Select
            Next
        End If


        'PollIO
        gbusy = 0

    End Function

    ' command = -1 just address device
    ' dv="" then just send command
    Private Function SendCommand(ByRef hcc As String, ByRef dv As String, _
                                 ByRef Command As Short, ByRef DimPercent As Byte, _
                                 ByRef data1 As Byte, ByRef data2 As Byte) As Short
        Dim hcv As Byte
        Dim dimv As Byte
        Dim dvstr As String
        Dim index As Short
        Dim key As Byte
        Dim buf(7) As Byte
        Dim i As Short
        Dim j As Integer
        Dim s As String
        On Error Resume Next


        hcv = (Asc(UCase(hcc)) - &H41S) ' makes "A" = 0
        If hcv > 16 Then Exit Function

        'Debug.Print("Send Command")

        gSendBusy = 1

        Select Case Command
            Case -1
                ' just send address
            Case PRESET_DIM_1
                If DimPercent > 31 Then DimPercent = 31
                dimv = DimPercent
                If dimv > 15 Then
                    dimv = dimv - 16
                    Command = 31
                Else
                    Command = 26
                End If
                dimv = gReverse(dimv) ' reverse bits

                j = 1
                Do
                    s = MidString(dv, j, "+")
                    If s = "" Then Exit Do
                    ' address device first
                    i = Val(s)
                    key = i - 1
                    'buf(0) = 200
                    'buf(1) = 55
                    'buf(2) = hcv
                    'buf(3) = key
                    'buf(4) = 1
                    'buf(5) = 0
                    'buf(6) = 0
                    'buf(7) = CheckSum(buf)
                    'SendToCPUXA(buf)
                    SendToCPUXA(FrameGenerator.SendX10(hcv, key, 1))
                    WaitForChar(gTimeOut, 3, 0)
                    j = j + 1
                Loop

                'buf(0) = 200
                'buf(1) = 55
                'buf(2) = dimv
                'buf(3) = Command
                'buf(4) = 1
                'buf(5) = 0
                'buf(6) = 0
                'buf(7) = CheckSum(buf)
                'SendToCPUXA(buf)
                SendToCPUXA(FrameGenerator.SendX10(dimv, Command, 1))
                WaitForChar(gTimeOut, 3, 0)
            Case EXTENDED_CODE
                j = 1
                Do
                    s = MidString(dv, j, "+")
                    If s = "" Then Exit Do
                    'buf(0) = 200
                    'buf(1) = 97
                    'buf(2) = hcv
                    'buf(3) = Val(s) - 1
                    'buf(4) = data1
                    'buf(5) = data2
                    'buf(6) = 0
                    'buf(7) = CheckSum(buf)
                    'SendToCPUXA(buf)
                    SendToCPUXA(FrameGenerator.SendLevitonX10PresetDim(hcv, Convert.ToByte(s) - 1, data1))
                    WaitForChar(gTimeOut, 3, 0)
                    j = j + 1
                Loop

            Case UON, UOFF, UDIM, UBRIGHT
                If Command = UDIM Or Command = UBRIGHT Then
                    dimv = 22 * (DimPercent / 100)
                    'dimv = dimv * 8  ' shift left by 3
                Else
                    dimv = 0
                End If
                ' address all devices
                ' if no devices, just send command
                If Len(dv) <> 0 Then
                    index = 1
                    Do
                        dvstr = Mid(dv, index)
                        'Debug.Print "dvstr: "; dvstr
                        i = Val(dvstr)
                        If (i >= 1) And (i <= 16) Then
                            key = i - 1
                            'buf(0) = 200
                            'buf(1) = 55
                            'buf(2) = hcv
                            'buf(3) = key
                            'buf(4) = 1
                            'buf(5) = 0
                            'buf(6) = 0
                            'buf(7) = CheckSum(buf)
                            'SendToCPUXA(buf)
                            SendToCPUXA(FrameGenerator.SendX10(hcv, key, 1))
                            WaitForChar(gTimeOut, 3, 0)
                        End If
                        index = InStr(index, dv, "+")
                        If index <> 0 Then
                            index = index + 1
                        End If
                    Loop While index <> 0
                End If
                ' send command
                'buf(0) = 200
                'buf(1) = 55
                'buf(2) = hcv
                'If Command = UON Then
                '    buf(3) = 18
                '    buf(4) = 1
                'ElseIf Command = UOFF Then
                '    buf(3) = 19
                '    buf(4) = 1
                'ElseIf Command = UDIM Then
                '    buf(3) = 20
                '    buf(4) = dimv
                'ElseIf Command = UBRIGHT Then
                '    buf(3) = 21
                '    buf(4) = dimv
                'End If
                'buf(5) = 0
                'buf(6) = 0
                'buf(7) = CheckSum(buf)
                'SendToCPUXA(buf)
                Select Case Command
                    Case UON
                        SendToCPUXA(FrameGenerator.SendX10(hcv, clsFrameGenerator.X10KeyCode._On, 1))
                    Case UOFF
                        SendToCPUXA(FrameGenerator.SendX10(hcv, clsFrameGenerator.X10KeyCode._Off, 1))
                    Case UDIM
                        SendToCPUXA(FrameGenerator.SendX10(hcv, clsFrameGenerator.X10KeyCode._Dim, dimv))
                    Case UBRIGHT
                        SendToCPUXA(FrameGenerator.SendX10(hcv, clsFrameGenerator.X10KeyCode._Bright, dimv))
                End Select

                WaitForChar(gTimeOut, 3, 0)
            Case Else
                ' address device first
                If Len(dv) > 0 Then
                    i = Val(dv)

                    key = i - 1
                    'buf(0) = 200
                    'buf(1) = 55
                    'buf(2) = hcv
                    'buf(3) = key
                    'buf(4) = 1
                    'buf(5) = 0
                    'buf(6) = 0
                    'buf(7) = CheckSum(buf)
                    'SendToCPUXA(buf)
                    SendToCPUXA(FrameGenerator.SendX10(hcv, key, 1))

                    WaitForChar(gTimeOut, 3, 0)
                End If

                Command = Command + 16
                If Command > 26 Then
                    Command = Command - 1
                End If

                'buf(0) = 200
                'buf(1) = 55
                'buf(2) = hcv
                'buf(3) = Command
                'buf(4) = 1
                'buf(5) = 0
                'buf(6) = 0
                'buf(7) = CheckSum(buf)
                'SendToCPUXA(buf)
                SendToCPUXA(FrameGenerator.SendX10(hcv, Command, 1))

                WaitForChar(gTimeOut, 3, 0)

        End Select

        gSendBusy = 0
        'Debug.Print hcc + dv + "CMD: " + str(command) + "DIM: " + str(dimpercent)


    End Function

    Public Function SetVar(ByRef vnum As Byte, ByRef data As UInt16) As Short
        Dim buf(7) As Byte

        gbusy = 1

        ' set our local copy
        'gVarsCopy(vnum + 1) = data

        'buf(0) = 200
        'buf(1) = 41
        'buf(2) = CByte(vnum)
        'buf(3) = data And &HFFS
        'If data <> 0 Then
        '    buf(4) = Int(data / 256)
        'Else
        '    buf(4) = 0
        'End If
        'buf(5) = 0
        'buf(6) = 0
        'buf(7) = CheckSum(buf)
        'SendToCPUXA(buf)
        SendToCPUXA(FrameGenerator.WriteCPUXAVariableData(vnum, data))
        WaitForChar(gTimeOut, 3, 0)

        SetVar = 0
        gbusy = 0


    End Function

    'UPGRADE_NOTE: beep was upgraded to beep_Renamed. Click for more: 'ms-help://MS.VSCC.v80/dv_commoner/local/redirect.htm?keyword="A9E4979A-37FA-4718-9994-97DD76ED70A7"'
    Public Function WriteScreen(ByRef X As Short, ByRef Y As Short, ByRef txt As String, ByRef backlight As Boolean, ByRef font As Short, ByRef clear As Boolean, ByRef beep_Renamed As Boolean) As Short
        Dim i As Short
        Dim j As Short
        Dim attr As Byte
        Dim buf(50) As Byte
        On Error Resume Next


        If backlight Then attr = attr Or &H1S
        If font = 1 Then attr = attr Or &H2S
        If clear Then attr = attr Or &H4S
        If beep_Renamed Then attr = attr Or &H8S

        gbusy = 1

        SendToCPUXA(FrameGenerator.WriteTexttoLeopardTouchScreen(attr, X, Y, txt))
        WaitForChar(gTimeOut, 3, 0)

        WriteScreen = 0
        gbusy = 0


    End Function

    Public Sub SendToCPUXA(ByRef buf() As Byte) 'As Short
        If buf Is Nothing Then Exit Sub
        If buf.Length < 1 Then Exit Sub
        If Not gmoduleOK Then Exit Sub
        If MSComm1 Is Nothing Then Exit Sub
        If Not MSComm1.IsOpen Then Exit Sub
        Try
            MSComm1.Write(buf, 0, buf.Length)
        Catch ex As Exception
        End Try
    End Sub

    ' buf(0) = ignored
    ' buf(1) = house
    ' buf(2) = code
    Private Sub ProcessX10(ByRef buf() As Byte)
        Dim dimv As Short
        Dim i As Short
        Dim data1 As String
        Dim cmd As Short
        On Error Resume Next

        If Not gX10Enabled Then Exit Sub


        'Debug.Print("PROCESS X10")
        'If buf(3) = &HFF And buf(4) = &HFF Then Exit Sub

        'Debug.Print "buf(3): " + Hex(buf(3)) + " buf(4): " + Hex(buf(4)) + Str(Timer)
        If buf(2) < 16 Then
            'Debug.Print "GOT ADDRESS"
            ' its a device address, add to list
            If gDevices = "" Then
                gDevices = Chr(Asc("A") + buf(1)) & Trim(Str(buf(2) + 1))
            Else
                gDevices = gDevices & "+" & Chr(Asc("A") + buf(1)) & Trim(Str(buf(2) + 1))
            End If
        Else
            'Debug.Print "GOT COMMAND"
            ' must be a command, call event
            If buf(2) = 26 Or buf(2) = 31 Then
                ' handle preset dim differently
                For i = 0 To 15
                    If gReverse(i) = buf(1) Then
                        Exit For
                    End If
                Next
                If buf(2) = 31 Then
                    i = i + 16
                    cmd = PRESET_DIM_2
                Else
                    cmd = PRESET_DIM_1
                End If
                data1 = Trim(Str(i))
                'TODO: callback.X10Event(gDevices, Mid(gDevices, 1, 1), cmd, data1, "")
            Else
                Select Case buf(2)
                    Case 16
                        cmd = 0
                    Case 17
                        cmd = 1
                    Case 18
                        cmd = 2
                    Case 19
                        cmd = 3
                    Case 20
                        cmd = 4
                    Case 21
                        cmd = 5
                    Case 22
                        cmd = 6
                    Case 23
                        cmd = 7
                    Case 24
                        cmd = 8
                    Case 25
                        cmd = 9
                    Case 27
                        cmd = 12
                    Case 28
                        cmd = 13
                    Case 29
                        cmd = 14
                    Case 30
                        cmd = 15
                    Case Else
                        ' invalid command, ignore
                        GoTo done
                End Select
                'TODO: callback.X10Event(gDevices, Chr(Asc("A") + buf(1)), cmd, "", "")
            End If
            ' clear devices
            gDevices = ""
        End If
done:


    End Sub

    Private Sub FlushCommInput()
        Dim buf(3) As Byte
        Dim cnt As Integer

        Try
            cnt = MSComm1.BytesToRead
            If cnt = 0 Then Exit Sub

            Dim s As String = ""
            s = MSComm1.ReadExisting()
        Catch ex As Exception
        End Try

    End Sub

    Private Function WaitForChar(ByRef secs As Integer, ByRef count As Integer, ByRef first_byte As Integer) As Boolean
        'Private Function WaitForChar(ByRef secs As Short, ByRef count As Short, ByRef first_byte As Short) As Boolean
        Dim j As Integer 'Short
        Dim i As Integer 'Short
        Dim buf(3) As Byte
        Dim b As Byte
        Dim cnt As Integer 'Short

        On Error Resume Next

        j = 0
        Do
check:
            System.Windows.Forms.Application.DoEvents()
            'cnt = MSComm1.InBufferCount
            cnt = MSComm1.BytesToRead
            If cnt >= count Then
                'b = MSComm1.Input
                b = MSComm1.ReadByte
                buf(0) = b '(0)
                If buf(0) = &HFE Or buf(0) = &HFB Then
                    b = MSComm1.ReadByte 'MSComm1.Input
                    buf(1) = b '(0)
                    b = MSComm1.ReadByte 'MSComm1.Input
                    buf(2) = b '(0)
                    If buf(0) = &HFE Then
                        If gLogComms Then
                            WriteLog("******************* Ocelot says X10 rec (1)")
                        End If
                        ProcessX10(buf)
                    End If
                    GoTo check
                ElseIf buf(0) = &HFFS Then
                    'If gLogComms Then
                    'WriteLog("******************* Ocelot says IO has changed (3)")
                    'End If
                    'gDoPoll = True
                    GoTo check
                ElseIf buf(0) = &HFDS Then
                    b = MSComm1.ReadByte '.Input
                    buf(1) = b '(0)
                    b = MSComm1.ReadByte '.Input
                    buf(2) = b '(0)
                    If gLearnWait Then
                        WaitForChar = True
                        Exit Function
                    End If
                    If gIREnabled Then
                        'hs.WriteLog "DEBUG", "RCV IR MATCH " + Str(buf(1))
                        'TODO:callback.IREvent(CShort(buf(2)))
                    End If
                    GoTo check
                ElseIf buf(0) = &HFCS Then
                    ' IR sent
                    b = MSComm1.ReadByte '.Input
                    b = MSComm1.ReadByte '.Input
                    GoTo check
                ElseIf buf(0) = &HF0S Then
                    ' io point change, next 3 bytes are status
                    'Debug.Print("IO point change")
                    b = MSComm1.ReadByte '.Input
                    buf(1) = b '(0)
                    b = MSComm1.ReadByte '.Input
                    buf(2) = b '(0)
                    b = MSComm1.ReadByte '.Input
                    buf(3) = b '(0)
                    HandleIOPointChange(buf)
                    'hs.WriteLog "IO DEBUG", "Input change detected"
                ElseIf buf(0) = &HF3S Or buf(0) = &HF2S Then
                    ' leopard touch coordinates
                    b = MSComm1.ReadByte '.Input
                    buf(1) = b '(0)
                    b = MSComm1.ReadByte '.Input
                    buf(2) = b '(0)
                    b = MSComm1.ReadByte '.Input
                    buf(3) = b '(0)
                    b = MSComm1.ReadByte '.Input

                    UpdateHSDeviceValue(0, 0, -1, True, False, Convert.ToInt32(buf(2)), CPUXA_UnitTypes.Leopard)
                    'Dim dv As Scheduler.Classes.DeviceClass = FindDevice(0, 0, -1, True, False)
                    'If dv IsNot Nothing Then
                    '    hs.SetDeviceValueByRef(dv.Ref(Nothing), buf(2), True)
                    '    If gLogComms Then
                    '        WriteLog("LEOPARD Touch detected: " & Str(buf(2)))
                    '    End If
                    'End If
                ElseIf buf(0) <> first_byte And first_byte <> 0 Then
                    ' might be junk characters at beginnig of buffer
                    GoTo check
                Else
                    ' must have the correct data
                    rbuffer(0) = buf(0)
                    For i = 1 To count - 1
                        b = MSComm1.ReadByte '.Input
                        rbuffer(i) = b '(0)
                    Next
                End If
                Return True
            End If
            waitms(200)
            j = j + 200
        Loop While j < (secs * 1000)
        If gLogComms Then
            WriteLog("Ocelot Timeout " & "Expected: " & count.ToString & " Got: " & cnt.ToString)
        End If
        Return False
    End Function

    Public Sub SetTimeOut(ByRef timeout As Short)
        gTimeOut = timeout
    End Sub

    Private Sub HandleIOPointChange(ByRef buf() As Byte)
        Dim um As umap
        'Dim h As String = ""
        Dim stat As Short
        Dim unit As Short
        Dim p As Short
        Dim mess As String = ""


        ' we know we have at least 3 chars in the buffer, but we need 4
        ' check for last char
        If buf Is Nothing Then Exit Sub
        If buf.Length < 3 Then Exit Sub

        If gLogComms Then
            WriteLog("IO point change mod= " & buf(1).ToString & " point= " & buf(2).ToString & " newval= " & buf(3).ToString)
        End If

        Try
            'hs.WriteLog "IO DEBUG", "IO Input change, point-val " + Str(buf(2)) + " " + Str(buf(3))
            SetInternalPoint(buf(1), buf(2), buf(3))
            For Each um In gUnitMappings
                ' hibyte = real time inputs
                ' lobyte = relay outputs
                unit = um.unit
                If unit = buf(1) Then
                    ' found our unit
                    ' must be a secu16 or secu16i
                    p = GetPointValue(buf(1), buf(2))
                    If gInvert Then
                        If p = 1 Then
                            p = 0
                        Else
                            p = 1
                        End If
                    End If
                    If p = 1 Then
                        stat = UON
                        mess = "ON"
                    Else
                        stat = UOFF
                        mess = "OFF"
                    End If
                    If gLogComms Then
                        WriteLog("Setting HS device (HandleIOPointChange): Unit=" & unit.ToString & ", Point=" & (buf(2) + 1).ToString & "(Real=" & buf(2).ToString & ") to " & mess)
                    End If
                    'Debug.Print("Calling SetDeviceStatus")
                    Dim UnitType As CPUXA_UnitTypes
                    Try
                        UnitType = gUnits(unit)
                    Catch ex As Exception
                        UnitType = 0
                    End Try
                    If UnitType = CPUXA_UnitTypes.Secu16IR Then Exit Sub
                    Dim point As Short = buf(2)
                    Select Case UnitType
                        Case CPUXA_UnitTypes.Rly08XA, CPUXA_UnitTypes.Secu16i, CPUXA_UnitTypes.Secu16IR
                            point += 1  ' HomeSeer devices are 1's based! (What a PITA Rich!!)
                        Case CPUXA_UnitTypes.Secu16
                            point += 1  ' This handles the input only, so no need to add 9 for the outputs.
                    End Select
                    UpdateHSDeviceValue(unit, point, -1, True, True, p, UnitType)
                    Exit For
                End If
            Next um
        Catch ex As Exception
        End Try

    End Sub

    Private Sub PollQueues()
        Dim i As Integer 'Short
        Dim c As Integer 'Short
        Dim h As String
        Dim unit As Integer 'Short
        Dim j As Integer 'Short
        Dim p As Integer 'Short
        Dim stat As Integer 'Short
        Dim cs As Integer 'Short
        Dim buf(4) As Byte
        Dim b As Byte 'Object
        'Dim tm As Integer 'String

        On Error Resume Next
        If gSendBusy = 1 Then Exit Sub
        If gbusy = 1 Then Exit Sub


        ' handle X10 commands
        ' get all x10 commands in buffer
        ' problem with polling, we get back all commands we send
        'If PollX10 = False Then Exit Do
        Do
            c = MSComm1.BytesToRead '.InBufferCount
            If gLogComms AndAlso c > 0 Then
                WriteLog("Comm byte count: " & c.ToString)
            End If

            If c < 3 Then Exit Do

            MSComm1.Read(buf, 0, 3)
            If gLogComms Then
                WriteLog("Rcv data: b0=" & buf(0).ToString & " b1=" & buf(1).ToString & " b2=" & buf(2).ToString)
            End If
            Select Case buf(0)
                Case &HF0
                    ' get last char
                    If c = 3 Then
                        Dim dt As Date = Now
                        Do
                            If MSComm1.BytesToRead > 0 Then
                                b = MSComm1.ReadByte
                                Exit Do
                            End If
                            Thread.Sleep(3)
                            Windows.Forms.Application.DoEvents()
                        Loop Until Now.Subtract(dt).TotalMilliseconds > 999
                    Else
                        b = MSComm1.ReadByte '.Input
                    End If
                    buf(3) = b '(0) ' new value
                    HandleIOPointChange(buf)
                Case &HFE ' x10 received
                    If gLogComms Then
                        WriteLog("******************* Ocelot says x10 rec")
                    End If
                    ProcessX10(buf)
                Case &HFB ' xmit x10
                Case &HFF
                    'Debug.Print("IO Change (2)")
                  
                Case &HFD ' rcv IR match
                    If gIREnabled Then
                        'hs.WriteLog "DEBUG", "RCV IR MATCH " + Str(buf(1))
                        'TODO: callback.IREvent(CShort(buf(1)))
                    End If
                Case &HFC ' xmit IR
                Case &HF3, &HF2 ' Leopard touch grid # (F3) button # (F2)
                    ' get last byte
                    b = MSComm1.ReadByte '.Input

                    UpdateHSDeviceValue(0, 0, -1, True, False, Convert.ToInt32(buf(2)), CPUXA_UnitTypes.Leopard)

            End Select

        Loop


check_q:
        ' handle sending of X10 commands and other commands
        Dim zone As Short
        Dim sName As String = ""
        If gQsize > 0 Then
            Do

                sName = GetName(gCmdQueue(gQtail).dvref)

                If gCmdQueue(gQtail).cm < VALUE_SET Then
                    SendCommand(gCmdQueue(gQtail).hcc, gCmdQueue(gQtail).dc, gCmdQueue(gQtail).cm, CShort(gCmdQueue(gQtail).br), gCmdQueue(gQtail).d1, gCmdQueue(gQtail).d2)
                Else
                    Select Case gCmdQueue(gQtail).cm
                        Case VALUE_SET
                            SetVar(Val(gCmdQueue(gQtail).variable), gCmdQueue(gQtail).br)
                            hs.WriteLog(IFACE_NAME, "Setting Variable " & sName & " to " & gCmdQueue(gQtail).br.ToString)
                            hs.SetDeviceValueByRef(gCmdQueue(gQtail).dvref, CInt(gCmdQueue(gQtail).br), True)
                        Case SEND_IR
                            'hs.WriteLog("DEBUG", "UnQueuing PLAY IR: Location: " + Str(gCmdQueue(gQtail).br) + " Zone: " + Str(gCmdQueue(gQtail).d1))
                            If gSecu16IR_present Then
                                zone = gCmdQueue(gQtail).d1
                                'hs.WriteLog "DEBUG", "Sending IR to zone " + Str(zone - 1)
                            Else
                                zone = 0
                            End If
                            If zone > 0 And zone <> 17 Then
                                DoPlayRemoteIR(gSecu16IR_Unit, zone - 1, gCmdQueue(gQtail).br)
                            Else
                                DoPlayIR(gCmdQueue(gQtail).br)
                            End If
                        Case SET_RELAY
                            DoSetPoint(gCmdQueue(gQtail).unit, gCmdQueue(gQtail).point, CShort(gCmdQueue(gQtail).br))
                            hs.WriteLog(IFACE_NAME, "Setting Relay " & sName & " to " & gCmdQueue(gQtail).br.ToString)
                            WriteLog("Setting Relay " & sName & " to " & gCmdQueue(gQtail).br.ToString)
                            hs.SetDeviceValueByRef(gCmdQueue(gQtail).dvref, CInt(gCmdQueue(gQtail).br), True)
                            PollInputs()
                        Case SEND_REMOTE_IR
                            DoPlayRemoteIR(CShort(gCmdQueue(gQtail).br), gCmdQueue(gQtail).d1, CInt(gCmdQueue(gQtail).d2))
                        Case LEARN_IR
                            DoLearnIR(gCmdQueue(gQtail).br)
                    End Select
                End If
                gQtail = gQtail + 1
                If gQtail >= QMAX Then
                    gQtail = 0
                End If
                gQsize = gQsize - 1
            Loop While gQsize > 0
        End If
done:
    End Sub

    Friend Function GetName(ByVal dvRef As Integer) As String
        Dim sName As String = ""
        Try
            sName = hs.DeviceName(dvRef)
        Catch ex As Exception
            sName = "(Unknown)"
        End Try
        Return sName
    End Function


    Private Sub PollInputs()
        Dim um As umap
        Dim h As String = ""
        Dim unit As Short
        Dim j As Short
        Dim p As Short
        Dim stat As Integer
        Dim cs As Integer
        Dim buf(4) As Byte
        Dim dv As Scheduler.Classes.DeviceClass

        On Error Resume Next

        ' only poll the dig inputs if we got an FF input change notification
        If PollIO() = False Then Return

        'hs.writelog("OCELOT", "Poll Inputs")

        For Each um In gUnitMappings
            If gQsize > 0 Then Exit For
            Select Case um.utype
                Case 11 ' secu16
                    ' hibyte = real time inputs
                    ' lobyte = relay outputs

                    unit = um.unit

                    ' digital IO are first 8 codes
                    For j = 1 To 8
                        p = GetPointValue(unit, j - 1)
                        If gInvert Then
                            p = p Xor 1
                            'If p = 1 Then
                            '    p = 0
                            'Else
                            '    p = 1
                            'End If
                        End If
                        UpdateHSDeviceValue(unit, j, -1, False, True, p, CPUXA_UnitTypes.Secu16)
                        'dv = FindDevice(unit, j, -1, False, True)
                        'If dv IsNot Nothing Then
                        '    cs = hs.DeviceValue(dv.Ref(Nothing))
                        '    If cs <> p Then
                        '        hs.SetDeviceValueByRef(dv.Ref(Nothing), p, True)
                        '        If gLogComms Then
                        '            WriteLog("SECU16 Setting device " & unit.ToString & j.ToString & " to " & p.ToString)
                        '        End If
                        '    End If
                        'End If
                    Next
                    ' relay outputs
                    For j = 1 To 8
                        p = GetPointValue(unit, (j - 1) + 8)
                        ' these read inverted from inputs
                        p = p Xor 1
                        'If p = 0 Then
                        '    p = 1
                        'Else
                        '    p = 0
                        'End If
                        UpdateHSDeviceValue(unit, j, -1, False, False, p, CPUXA_UnitTypes.Secu16)
                        'dv = FindDevice(unit, j, -1, False, False)
                        'If dv IsNot Nothing Then
                        '    cs = hs.DeviceValue(dv.Ref(Nothing))
                        '    If cs <> p Then
                        '        hs.SetDeviceValueByRef(dv.Ref(Nothing), p, True)
                        '        If gLogComms Then
                        '            WriteLog("SECU16 Setting device " & unit.ToString & Trim(Str(j + 8)) & " to " & Str(p))
                        '        End If
                        '    End If
                        'End If
                    Next

                Case 12 ' secu16i
                    ' 16 inputs only
                    unit = um.unit
                    If h <> "" Then
                        ' digital IO are first 16 codes
                        For j = 1 To 16
                            p = GetPointValue(unit, j - 1)
                            If gInvert Then
                                p = p Xor 1
                                'If p = 1 Then
                                '    p = 0
                                'Else
                                '    p = 1
                                'End If
                            End If
                            UpdateHSDeviceValue(unit, j, -1, False, True, p, CPUXA_UnitTypes.Secu16i)
                            'dv = FindDevice(unit, j, -1, False, True)
                            'If dv IsNot Nothing Then
                            '    cs = hs.DeviceValue(dv.Ref(Nothing))
                            '    If cs <> p Then
                            '        hs.SetDeviceValueByRef(dv.Ref(Nothing), p, True)
                            '        If gLogComms Then
                            '            WriteLog("SECU16I Setting device " & unit.ToString & Trim(Str(j)) & " to " & Str(stat))
                            '        End If
                            '    End If
                            'End If
                        Next
                    End If
                Case 13 ' rly8
                    ' 8 outputs only
                    unit = um.unit
                    If h <> "" Then
                        ' relays are first 8 codes
                        For j = 1 To 8
                            p = GetPointValue(unit, (j - 1) + 8)
                            ' these read inverted from inputs
                            p = p Xor 1
                            'If p = 0 Then
                            '    p = 1
                            'Else
                            '    p = 0
                            'End If
                            UpdateHSDeviceValue(unit, j, -1, False, False, p, CPUXA_UnitTypes.Rly08XA)
                            'dv = FindDevice(unit, j, -1, False, False)
                            'If dv IsNot Nothing Then
                            '    cs = hs.DeviceValue(dv.Ref(Nothing))
                            '    If cs <> p Then
                            '        hs.SetDeviceValueByRef(dv.Ref(Nothing), p, True)
                            '        If gLogComms Then
                            '            WriteLog("RLY8 Setting device " & unit.ToString & Trim(Str(j + 8)) & " to " & Str(stat))
                            '        End If
                            '    End If
                            'End If
                        Next
                    End If

            End Select
        Next um
    End Sub


    Private Sub PollVars()
        Dim i As Short
        Dim p As Short
        Dim buf(3) As Byte

        Try
            ' exit if in the middle of sending a command

            If gSendBusy = 1 Then Exit Sub
            If gbusy = 1 Then Exit Sub
            If gQsize Then Exit Sub

            'hs.writelog("OCELOT", "Poll Vars")
start:
            ' get IR matches
            'PollIR

            ' get analog IO, init get first
            If Not gAinit Then
                ' init get first
            End If

            ' just exit if call times out or other error
            ' want to avoid bogus readings
            'FlushCommInput

            ' get vars
            ' no auto notification for these, do at poll interval
            'FlushCommInput

            If GetVarsFromHW() Then
                ' assign to devices
                p = 0

                For i = 1 To 128
                    If gQsize > 0 Then Exit For

                    UpdateHSDeviceValue(0, 0, i, False, False, Convert.ToInt32(gVars(p)), CPUXA_UnitTypes.Variable)

                    p = p + 1
                    If gQsize Then Exit For
                Next

                ' we now have a local copy of the HS Ocelot vars
                gHaveRealValues = True
            End If

        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error in PollVars(" & Erl.ToString & "): " & ex.Message)
        End Try
    End Sub

    ' speed up response and kick poll when data received
    ' will do nothing if we are waiting for chars
    Private Sub MSComm1_OnComm(ByVal Sender As Object, ByVal e As System.IO.Ports.SerialDataReceivedEventArgs) Handles MSComm1.DataReceived
        GotRcvData = True
    End Sub
    'Private Sub MSComm1_OnComm(ByVal event_type As rs232.comm_event_types) Handles MSComm1.OnComm
    '    Dim r As rs232.comm_event_types
    '    r = MSComm1.CommEvent
    '    If r = rs232.comm_event_types.comEvReceive Then
    '        'If gLogComms Then
    '        '    WriteLog "MSCOMM"
    '        'End If
    '        GotRcvData = True
    '    End If
    'End Sub

#End Region

End Class
