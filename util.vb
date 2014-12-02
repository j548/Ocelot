Option Strict Off
Option Explicit On

Imports System.Runtime.InteropServices
Imports System.IO
Imports System.Threading
Imports HomeSeerAPI.PlugExtraData
Imports HomeSeerAPI.VSVGPairs
Imports HomeSeerAPI

Module util

    Public Instance As String = ""
    Public callback As HomeSeerAPI.IAppCallbackAPI
    Public hs As HomeSeerAPI.IHSApplication

    Friend ConfigPage As ConfigWebPag = Nothing

    ' capabilites of device (this OCX) (bits)
    Public Const CA_X10 As Short = 1 ' supports X10
    Public Const CA_IR As Short = 2 ' supports infrared
    Public Const CA_IO As Short = 4 ' supports I/O
    Public Const CA_SEC As Short = 8 ' supports security


    ' X10 commands
    Public Const ALL_UNITS_OFF As Short = 0
    Public Const ALL_LIGHTS_ON As Short = 1
    Public Const UON As Short = 2
    Public Const UOFF As Short = 3
    Public Const UDIM As Short = 4
    Public Const UBRIGHT As Short = 5
    Public Const ALL_LIGHTS_OFF As Short = 6
    Public Const EXTENDED_CODE As Short = 7
    Public Const HAIL_REQUEST As Short = 8
    Public Const HAIL_ACK As Short = 9
    Public Const PRESET_DIM_1 As Short = 10
    Public Const PRESET_DIM_2 As Short = 11
    Public Const X_DATA_XFER As Short = 12
    Public Const STATUS_ON As Short = 13
    Public Const STATUS_OFF As Short = 14
    Public Const STATUS_REQUEST As Short = 15
    Public Const DIM_TO_OFF As Short = 16 ' not x10
    Public Const NO_X10 As Short = 17 ' not x10 (unknown state)
    Public Const X10_ANY As Short = 18 ' for trigger, any X10 command
    Public Const VALUE_SET As Short = 19 ' for hspi setio call

    ' for this module only
    Public Const SEND_IR As Short = 100
    Public Const SET_RELAY As Short = 101
    Public Const SEND_REMOTE_IR As Short = 102
    Public Const LEARN_IR As Short = 103

    ' interface status
    ' for InterfaceStatus function call
    Public Const ERR_NONE As Short = 0
    Public Const ERR_SEND As Short = 1
    Public Const ERR_INIT As Short = 2

    ' types of devices
    Public Const IRC_SERIAL As Short = 1
    Public Const IRC_LPT As Short = 2
    Public Const IRC_OTHER As Short = 3


    Public gDevices As String
    Public gReverse As Object
    Public gSecu16IR_present As Boolean ' true if secu16ir unit is installed
    Public gSecu16IR_Unit As Short

    'Public rbuffer(RBUF_SIZE) As Byte

    Public rcount As Integer 'Short
    'Dim rbuf_head_n As Integer
    'Dim rbuf_tail_n As Integer
    'Dim rbuf_count_n As Integer
    Public rbuf_head As Integer 'Short
    Public rbuf_tail As Integer 'Short
    Public rbuf_count As Integer 'Short
    Public gbusy As Integer 'Short
    Public Const MAX_IRKEYS As Integer = 36 'Short = 36 ' max keys per device
    Public Const MAX_DEVICES As Integer = 8 'Short = 8
    Public irtable As Object
    Public irdevtable As Object
    Public gComPort As String
    Public gNotify As Integer 'Short
    Public gBufNotify As Integer 'Short ' true = buffering notification bytes
    Friend gUnits(128) As Byte ' units installed
    Public gIO(256) As Byte ' io input for all units
    Public gNotifyBusy As Boolean
    Public gTimeOut As Integer 'Short
    Public gIRSize As Integer 'Short ' max number of IR locations
    Public Status As Integer 'Short ' 0=ok 1=cannot send

    Public gX10Enabled As Boolean
    Public gIREnabled As Boolean
    Public gIOEnabled As Boolean
    'Public gIRInitialized As Boolean ' TRUE=initialized
    'Public gIOInitialized As Boolean

    Public gmoduleOK As Boolean ' true = initialized

    Public gConfigPollVars As Boolean


    Friend gPollInterval As Integer
    Public gVars(128) As Integer ' ocelot variables
    'Public gVarsCopy(128) As Integer
    Public gHaveRealValues As Boolean
    Public gInvert As Boolean ' 1=invert inputs
    Public gmappings As String
    Public gLogErrors As Boolean
    Public gLogComms As Boolean ' log communications to ocelot.log
    Public InterfaceVersion As Integer 'Short

    Public gUnitMappings As Collection ' unit to housecode mappings

    ' io codes start at "[" and wrap to "#" at "a" and end at "@"
    Public Const MAX_IO_CODES As Short = 36
    Public Const MAX_DEVICE_CODES As Short = 64
    Public Const MAX_HOUSE_CODES As Short = 26 + MAX_IO_CODES
    Public Const IO_WRAP As Short = &H61 ' "a"   wrap to #
    Public Const FIRST_IO_CODE As Short = &H5B ' "["
    Public Const FIRST_HC As Short = 35 ' "#"

    Public Const IOTYPE_INPUT As Short = 0
    Public Const IOTYPE_OUTPUT As Short = 1
    Public Const IOTYPE_ANALOG_INPUT As Short = 2
    Public Const IOTYPE_VARIABLE As Short = 3
    Public Const IFACE_NAME As String = "Applied Digital Ocelot"
    Public Const IFACE_OTHER As String = "Ocelot"

    ' for device class misc flags, bits in long value
    Public Const MISC_PRESET_DIM As Short = 1 ' supports preset dim if set
    Public Const MISC_EXT_DIM As Short = 2 ' extended dim command
    Public Const MISC_SMART_LINC As Short = 4 ' smart linc switch
    Public Const MISC_NO_LOG As Short = 8 ' no logging to event log for this device
    Public Const MISC_STATUS_ONLY As Short = 16 ' device cannot be controlled
    Public Const MISC_HIDDEN As Short = &H20S ' device is hidden from views
    Public Const MISC_THERM As Short = &H40S ' device is a thermostat. Copied from dev attr

    Friend Class DataCache
        Public dvRef As Integer = 0
        Public dv As Scheduler.Classes.DeviceClass = Nothing
        Public DevValue As Integer = -1
        Public LastUpdate As Date = Date.MinValue
        Public OKToLog As Boolean = True
    End Class
    Friend colDataCache As New Collections.Concurrent.ConcurrentDictionary(Of String, DataCache)

    Friend Enum PEDID
        '_Leopard
        _Type
        '_Variable
        _Unit
        _Point
        '_Unit_Type
        _Input_Output
        '_Unit_Info
    End Enum
    Friend Function PEDString(ByVal PID As PEDID) As String
        Select Case PID
            'Case PEDID._Leopard
            '    Return "leopard"
            Case PEDID._Type
                Return "type"
                'Case PEDID._Variable
                '    Return "variable"
            Case PEDID._Unit
                Return "unit"
            Case PEDID._Point
                Return "point"
                'Case PEDID._Unit_Type
                '    Return "unit_type"
            Case PEDID._Input_Output
                Return "input_output"
                'Case PEDID._Unit_Info
                '    Return "unit_info"
        End Select
        Return "ERROR"
    End Function

    Public FrameGenerator As clsFrameGenerator = Nothing

    Friend Class clsFrameGenerator

        Public Function Debug(ByVal Frame() As Byte) As String
            If Frame Is Nothing Then Return ""
            If Frame.Length < 1 Then Return ""
            Try
                Dim s As String = ""
                For i As Integer = 0 To Frame.Length - 1
                    If s.Trim = "" Then
                        s = Frame(i).ToString("x2").ToUpper
                    Else
                        s &= "-" & Frame(i).ToString("x2").ToUpper
                    End If
                Next
                Return s
            Catch ex As Exception
                Return "ERROR"
            End Try
        End Function

        Public Function UNDOCUMENTED_GetMemoryData(ByVal gLoVar As Byte, ByVal gHiVar As Byte) As Byte()
            Dim snd(7) As Byte
            Dim t As Integer
            snd(0) = 42
            snd(1) = 0
            snd(2) = 0
            snd(3) = gLoVar
            snd(4) = gHiVar
            snd(5) = 0
            Err.Clear()
            t = CalcCRC(snd, 6)
            snd(7) = Convert.ToByte(t And &HFF)
            snd(6) = Convert.ToByte(Convert.ToInt16(t / 256))
            Return snd
        End Function

        Public Function UNDOCUMENTED_GetMemoryPointer() As Byte()
            Dim snd(7) As Byte
            snd(0) = 42
            snd(1) = 0
            snd(2) = 0
            snd(3) = 143
            snd(4) = 183
            snd(5) = 0
            snd(6) = 45
            snd(7) = 235
            Return snd
        End Function

        Public Function UNDOCUMENTED_GetIRSize() As Byte()
            Dim snd(7) As Byte
            snd(0) = 42
            snd(1) = 0
            snd(2) = 0
            snd(3) = 0
            snd(4) = 127
            snd(5) = 0
            snd(6) = 165
            snd(7) = 125
            Return snd
        End Function

        Public Function GetRealtimeIO() As Byte()
            Dim snd(7) As Byte
            snd(0) = 42
            snd(1) = 0
            snd(2) = 0
            snd(3) = 141
            snd(4) = 181
            snd(5) = 0
            snd(6) = 37
            snd(7) = 233
            Return snd
        End Function
        Public Function GetLatchedIO() As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 50
            snd(2) = 0
            snd(3) = 0
            snd(4) = 0
            snd(5) = 0
            snd(6) = 0
            snd(7) = 250
            Return snd
        End Function
        Public Function GetUnitType() As Byte()
            Dim snd(7) As Byte
            snd(0) = 42
            snd(1) = 0
            snd(2) = 0
            snd(3) = 1
            snd(4) = 176
            snd(5) = 0
            snd(6) = 148
            snd(7) = 39
            Return snd
        End Function
        Public Function GetUnitFirmwareVersion() As Byte()
            Dim snd(7) As Byte
            snd(0) = 42
            snd(1) = 0
            snd(2) = 0
            snd(3) = 130
            snd(4) = 176
            snd(5) = 0
            snd(6) = 246
            snd(7) = 45
            Return snd
        End Function
        Public Function InitiateGetUnitParameters(ByVal ParameterNumber As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 31
            snd(2) = ParameterNumber
            snd(3) = 0
            snd(4) = 0
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function WriteRelayOutput(ByVal Unit As Byte, ByVal RelayNumber As Byte, ByVal OnOff As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 51
            snd(2) = Unit
            snd(3) = RelayNumber
            snd(4) = OnOff
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function GetUnitParameters() As Byte()
            Dim snd(7) As Byte
            snd(0) = 42
            snd(1) = 0
            snd(2) = 0
            snd(3) = 139
            snd(4) = 180
            snd(5) = 0
            snd(6) = 164
            snd(7) = 120
            Return snd
        End Function
        Public Function InitiateCPUXARescan() As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 32
            snd(2) = 0
            snd(3) = 0
            snd(4) = 0
            snd(5) = 0
            snd(6) = 0
            snd(7) = 232
            Return snd
        End Function
        Public Function InitiateCPUXARestart() As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 52
            snd(2) = 0
            snd(3) = 0
            snd(4) = 0
            snd(5) = 0
            snd(6) = 0
            snd(7) = 252
            Return snd
        End Function
        Public Function WriteUnitParameterData(ByVal Unit As Byte, ByVal ParameterNumber As Byte, _
                                               ByVal DataLSB As Byte, ByVal DataMSB As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 33
            snd(2) = Unit
            snd(3) = DataLSB
            snd(4) = ParameterNumber
            snd(5) = DataMSB
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function WriteUnitParameterData(ByVal Unit As Byte, ByVal ParameterNumber As Byte, _
                                       ByVal Data As UInt16) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 33
            snd(2) = Unit
            snd(3) = Data And &HF
            snd(4) = ParameterNumber
            snd(5) = Convert.ToByte(Data >> 8)
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function WriteCPUXAParameterData(ByVal ParameterNumber As Byte, _
                                                ByVal DataLSB As Byte, ByVal DataMSB As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 40
            snd(2) = ParameterNumber
            snd(3) = DataMSB
            snd(4) = DataLSB
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function WriteCPUXAParameterData(ByVal ParameterNumber As Byte, _
                                                ByVal Data As UInt16) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 40
            snd(2) = ParameterNumber
            snd(3) = Data And &HF
            snd(4) = Convert.ToByte(Data >> 8)
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Enum X10HouseCodes As Byte
            A = 0
            B = 1
            C = 2
            D = 3
            E = 4
            F = 5
            G = 6
            H = 7
            I = 8
            J = 9
            K = 10
            L = 11
            M = 12
            N = 13
            O = 14
            P = 15
        End Enum
        Public Shared Function X10HouseCodeToKeyCode(ByVal HC As String) As Byte
            Try
                If HC Is Nothing Then Return &HF0 ' values are 0 to 15, so this is Invalid
                If String.IsNullOrEmpty(HC.Trim) Then Return &HF0
                If HC.Trim.Length > 1 Then HC = HC.Trim.Substring(0, 1)
                HC = HC.Trim.ToUpper
                ' A = 65
                Select Case Asc(HC)
                    Case 65 To 80
                        Return Convert.ToByte((Asc(HC) - 65) And &HF)
                    Case Else
                        Return &HF0
                End Select
            Catch ex As Exception
                Return &HF0
            End Try
        End Function
        Public Enum X10KeyCode As Byte
            U01 = 0
            U02 = 1
            U03 = 2
            U04 = 3
            U05 = 4
            U06 = 5
            U07 = 6
            U08 = 7
            U09 = 8
            U10 = 9
            U11 = 10
            U12 = 11
            U13 = 12
            U14 = 13
            U15 = 14
            U16 = 15
            _All_Units_Off = 16
            _All_Lights_On = 17
            _On = 18
            _Off = 19
            _Dim = 20
            _Bright = 21
            _All_Lights_Off = 22
            _Extend_Code = 23
            _Hail_Request = 24
            _Hail_Acknowledge = 25
            _Preset_Dim_0 = 26
            _Extend_Data = 27
            _Status_On = 28
            _Status_Off = 29
            _Status_Request = 30
            _Preset_Dim_1 = 31
        End Enum
        Public Function SendX10(ByVal HouseCode As Byte, ByVal KeyCode As Byte, _
                                ByVal Repeat As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 55
            snd(2) = HouseCode
            snd(3) = KeyCode
            snd(4) = Repeat
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function SendLevitonX10PresetDim(ByVal HouseCode As Byte, ByVal KeyCode As Byte, _
                                                ByVal _Dim As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 97
            snd(2) = HouseCode
            snd(3) = KeyCode
            snd(4) = _Dim
            snd(5) = 49
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function GetReceivedX10() As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 56
            snd(2) = 0
            snd(3) = 0
            snd(4) = 0
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function GetCPUXARealTimeClock() As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 62
            snd(2) = 0
            snd(3) = 0
            snd(4) = 0
            snd(5) = 0
            snd(6) = 0
            snd(7) = 6
            Return snd
        End Function
        Public Function SetCPUXARealTimeClock() As Byte()
            Dim snd(49) As Byte
            snd(0) = 241
            snd(1) = 61
            snd(2) = 0
            snd(3) = 0
            For i As Integer = 4 To 14
                Select Case i
                    Case 4 'Min
                        Try
                            snd(i) = Convert.ToByte(Now.Minute Mod 10)
                        Catch ex As Exception
                        End Try
                    Case 5 '10 Min
                        Try
                            snd(i) = Convert.ToByte((Now.Minute - (Now.Minute Mod 10)) / 10)
                        Catch ex As Exception
                        End Try
                    Case 6 'Hour
                        Try
                            snd(i) = Convert.ToByte(Now.Hour Mod 10)
                        Catch ex As Exception
                        End Try
                    Case 7 '10 Hour
                        Try
                            snd(i) = Convert.ToByte((Now.Hour - (Now.Hour Mod 10)) / 10)
                        Catch ex As Exception
                        End Try
                    Case 8 'Day
                        Try
                            snd(i) = Convert.ToByte(Now.Day Mod 10)
                        Catch ex As Exception
                        End Try
                    Case 9 '10 Day
                        Try
                            snd(i) = Convert.ToByte((Now.Day - (Now.Day Mod 10)) / 10)
                        Catch ex As Exception
                        End Try
                    Case 10 'Month
                        Try
                            snd(i) = Convert.ToByte(Now.Month Mod 10)
                        Catch ex As Exception
                        End Try
                    Case 11 '10 Month
                        Try
                            snd(i) = Convert.ToByte((Now.Month - (Now.Month Mod 10)) / 10)
                        Catch ex As Exception
                        End Try
                    Case 12 'Year
                        Try
                            snd(i) = Convert.ToByte((Now.Year - 2000) Mod 10)
                        Catch ex As Exception
                        End Try
                    Case 13 '10 Year
                        Try
                            snd(i) = Convert.ToByte(((Now.Year - 2000) - ((Now.Year - 2000) Mod 10)) / 10)
                        Catch ex As Exception
                        End Try
                    Case 14 ' Weekday
                        snd(i) = Convert.ToByte(Now.DayOfWeek)
                End Select
            Next
            For i As Integer = 15 To 48
                snd(i) = 0
            Next
            snd(49) = CheckSum50(snd)
            Return snd
        End Function
        Public Function SendResidentIR(ByVal IRNumberLSB As Byte, ByVal IRNumberMSB As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 90
            snd(2) = IRNumberLSB
            snd(3) = IRNumberMSB
            snd(4) = 0
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function SendResidentIR(ByVal IRNumber As UInt16) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 90
            snd(2) = IRNumber And &HF
            snd(3) = Convert.ToByte(IRNumber >> 8)
            snd(4) = 0
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function InitiateIRLearn(ByVal IRNumberLSB As Byte, ByVal IRNumberMSB As Byte, _
                                        ByVal Frequency As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 91
            snd(2) = IRNumberLSB
            snd(3) = IRNumberMSB
            snd(4) = Frequency
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function InitiateIRLearn(ByVal IRNumber As UInt16, _
                                        ByVal Frequency As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 90
            snd(2) = IRNumber And &HF
            snd(3) = Convert.ToByte(IRNumber >> 8)
            snd(4) = Frequency
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function WriteCPUXAVariableData(ByVal VariableNumber As Byte, _
                                               ByVal DataLSB As Byte, ByVal DataMSB As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 41
            snd(2) = VariableNumber
            snd(3) = DataLSB
            snd(4) = DataMSB
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function WriteCPUXAVariableData(ByVal VariableNumber As Byte, _
                                               ByVal Data As UInt16) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 41
            snd(2) = VariableNumber
            snd(3) = Data And &HF
            snd(4) = Convert.ToByte(Data >> 8)
            snd(5) = 0
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function

        Public Function SendRemoteIRCommand(ByVal Secu16IRUnit As Byte, ByVal Zone As Byte, _
                                            ByVal IRLocLSB As Byte, ByVal IRLocMSB As Byte) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 92
            snd(2) = Secu16IRUnit
            snd(3) = Zone
            snd(4) = IRLocLSB
            snd(5) = IRLocMSB
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function
        Public Function SendRemoteIRCommand(ByVal Secu16IRUnit As Byte, ByVal Zone As Byte, _
                                            ByVal IRLocation As UInt16) As Byte()
            Dim snd(7) As Byte
            snd(0) = 200
            snd(1) = 92
            snd(2) = Secu16IRUnit
            snd(3) = Zone
            snd(4) = IRLocation And &HF
            snd(5) = Convert.ToByte(IRLocation >> 8)
            snd(6) = 0
            snd(7) = CheckSum(snd)
            Return snd
        End Function

        Public Function WriteTexttoLeopardTouchScreen(ByVal Attributes As Byte, ByVal Count As Byte, _
                                                      ByVal XCoordLSB As Byte, ByVal XCoordMSB As Byte, _
                                                      ByVal YCoordLSB As Byte, ByVal YCoordMSB As Byte, _
                                                      ByVal ToWrite As String) As Byte()
            Dim snd(49) As Byte
            snd(0) = 241
            snd(1) = 109
            snd(2) = Attributes
            snd(3) = Count
            snd(4) = XCoordLSB
            snd(5) = XCoordMSB
            snd(6) = YCoordLSB
            snd(7) = YCoordMSB
            If ToWrite Is Nothing Then ToWrite = ""
            If ToWrite.Length > 41 Then
                ToWrite = ToWrite.Substring(0, 41)
            End If
            For i As Integer = 8 To 48
                snd(i) = Convert.ToByte(Asc(ToWrite.Substring(i - 8, 1)))
            Next
            snd(49) = CheckSum50(snd)
            Return snd
        End Function
        Public Function WriteTexttoLeopardTouchScreen(ByVal Attributes As Byte, _
                                                      ByVal XCoord As UInt16, _
                                                      ByVal YCoord As UInt16, _
                                                      ByVal ToWrite As String) As Byte()
            Dim snd(49) As Byte
            If ToWrite Is Nothing Then ToWrite = ""
            If ToWrite.Length > 41 Then
                ToWrite = ToWrite.Substring(0, 41)
            End If
            snd(0) = 241
            snd(1) = 109
            snd(2) = Attributes
            snd(3) = Convert.ToByte(ToWrite.Length)
            snd(4) = Convert.ToByte(XCoord And &HF)
            snd(5) = Convert.ToByte((XCoord >> 8) And &HF)
            snd(6) = Convert.ToByte(YCoord And &HF)
            snd(7) = Convert.ToByte((YCoord >> 8) And &HF)
            For i As Integer = 8 To 48
                snd(i) = Convert.ToByte(Asc(ToWrite.Substring(i - 8, 1)))
            Next
            snd(49) = CheckSum50(snd)
            Return snd
        End Function

    End Class

    Friend Function CheckSum(ByRef buf() As Byte) As Byte
        Dim i As Short
        Dim t As Short

        t = 0
        Try
            For i = 0 To 6
                t = t + buf(i)
            Next
            Dim r As Short = t And &HFF
            Return Convert.ToByte(r)
        Catch ex As Exception
            Return 0
        End Try

    End Function

    Friend Function CheckSum50(ByRef buf() As Byte) As Byte
        Dim i As Short
        Dim t As Short

        t = 0
        Try
            For i = 0 To 48
                t = t + buf(i)
            Next
            Dim r As Short = t And &HFF
            Return Convert.ToByte(r)
        Catch ex As Exception
            Return 0
        End Try
    End Function

    Friend Function LogPrefix(ByVal Unit As Integer, _
                              ByVal Point As Integer, _
                              ByVal Var As Integer, _
                              ByVal LeopardDevice As Boolean, _
                              ByVal InputDev As Boolean, _
                              ByVal UnitType As CPUXA_UnitTypes) As String
        Dim s As String = ""
        If Not [Enum].IsDefined(GetType(CPUXA_UnitTypes), UnitType) Then
            s = "(Unit " & Unit.ToString & ", Point " & Point.ToString
            If Var >= 0 Then
                s &= ", Var " & Var.ToString
            End If
            If LeopardDevice Then
                s &= ", Leopard"
            End If
            If InputDev Then
                s &= ", Input"
            Else
                s &= ", Output"
            End If
            s &= ")"
        Else
            Select Case UnitType
                Case CPUXA_UnitTypes.Leopard
                    s = "Leopard"
                Case CPUXA_UnitTypes.Variable
                    s = "Variable " & Var.ToString
                Case CPUXA_UnitTypes.Rly08XA
                    s = "RLY08XA Unit " & Unit.ToString & ", Relay " & Point.ToString
                Case CPUXA_UnitTypes.Secu16IR
                    s = "SECU16IR Unit " & Unit.ToString & ", Emitter " & Point.ToString
                Case CPUXA_UnitTypes.Secu16i
                    s = "SECU16I Unit " & Unit.ToString & ", Input " & Point.ToString
                Case CPUXA_UnitTypes.Secu16
                    If InputDev Then
                        s = "SECU16 Unit " & Unit.ToString & ", Input " & Point.ToString
                    Else
                        s = "SECU16 Unit " & Unit.ToString & ", Output " & Point.ToString
                    End If
            End Select
        End If
        Return s
    End Function

    Public Sub SaveSettingLocal(ByRef loc_Renamed As String, ByRef section As String, ByRef key As String, ByRef Value As Object)
        On Error Resume Next
        Dim v As String
        v = CStr(Value)
        hs.SaveINISetting("hspi_ocelot", key, v, "hspi_ocelot.ini")
    End Sub

    Public Function GetSettingLocal(ByRef loc As String, ByRef section As String, ByRef key As String, Optional ByRef default_value As String = "") As String
        Return hs.GetINISetting("hspi_ocelot", key, default_value, "hspi_ocelot.ini")
    End Function


    Sub DeleteAllDevices()

        Try
            Dim dv As Scheduler.Classes.DeviceClass
            Dim DE As Scheduler.Classes.clsDeviceEnumeration

            DE = hs.GetDeviceEnumerator
            If Not DE Is Nothing Then
                Do While Not DE.Finished
                    dv = DE.GetNext
                    If Not dv Is Nothing Then
                        If dv.Interface(Nothing) = IFACE_NAME Then
                            hs.DeleteDevice(dv.Ref(Nothing))
                        End If
                    End If
                Loop
            End If

            gmappings = ""
            SaveSettingLocal("hspi_ocelot", "Settings", "mappings", gmappings)
        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error deleting devices in DeleteAllDevices: " & ex.Message)
        End Try
        
        
    End Sub

    <Serializable()> _
    Public Enum enumIO
        _Input = 1
        _Output = 2
    End Enum

    Public Sub SetHSDeviceAddress(ByRef dv As Scheduler.Classes.DeviceClass, _
                                  ByVal Unit_Type As CPUXA_UnitTypes, _
                                  ByVal Unit As Short, _
                                  ByVal Point As Short, _
                                  ByVal IO As enumIO)
        Try
            If dv Is Nothing Then Exit Sub
            If Not [Enum].IsDefined(GetType(CPUXA_UnitTypes), Unit_Type) Then Exit Sub
            If Not [Enum].IsDefined(GetType(enumIO), IO) Then Exit Sub
            Dim Addr As String = ""
            Select Case Unit_Type
                Case CPUXA_UnitTypes.Leopard
                    Addr = "Leopard"
                Case CPUXA_UnitTypes.Variable
                    Addr = "VAR " & Point.ToString("000")
                Case CPUXA_UnitTypes.Secu16i
                    Addr = "IN_U" & Unit.ToString("000") & "_P" & Point.ToString("000")
                Case CPUXA_UnitTypes.Rly08XA
                    Addr = "OUT_U" & Unit.ToString("000") & "_P" & Point.ToString("000")
                Case CPUXA_UnitTypes.Secu16IR
                    Addr = "IR_U" & Unit.ToString("000") & "_P" & Point.ToString("000")
                Case CPUXA_UnitTypes.Secu16
                    If IO = enumIO._Input Then
                        Addr = "IN_U" & Unit.ToString("000") & "_P" & Point.ToString("000")
                    ElseIf IO = enumIO._Output Then
                        Addr = "OUT_U" & Unit.ToString("000") & "_P" & Point.ToString("000")
                    End If
            End Select
            If Not String.IsNullOrEmpty(Addr.Trim) Then
                dv.Address(hs) = Addr
                hs.SaveEventsDevices()
            End If
        Catch ex As Exception
            hs.WriteLogEx("Ocelot", "Exception setting HS device address: " & ex.Message, COLOR_RED)
        End Try
    End Sub
    Public Function UpdatePEDInfo(ByRef dv As Scheduler.Classes.DeviceClass, _
                             ByVal Unit_Type As CPUXA_UnitTypes, _
                             ByVal Unit As Short, _
                             ByVal Point As Short, _
                             ByVal IO As enumIO) As Boolean
        Dim PED As clsPlugExtraData = Nothing
        Try
            If dv Is Nothing Then Return False
            PED = dv.PlugExtraData_Get(hs)
            If PED Is Nothing Then PED = New clsPlugExtraData
            If Not [Enum].IsDefined(GetType(CPUXA_UnitTypes), Unit_Type) Then Return False
            If Not [Enum].IsDefined(GetType(enumIO), IO) Then Return False
            Try
                PED.RemoveNamed(PEDString(PEDID._Type))
                PED.AddNamed(PEDString(PEDID._Type), Unit_Type)
            Catch ex As Exception
            End Try
            Try
                PED.RemoveNamed(PEDString(PEDID._Unit))
                PED.AddNamed(PEDString(PEDID._Unit), Unit)
            Catch ex As Exception
            End Try
            Try
                PED.RemoveNamed(PEDString(PEDID._Point))
                PED.AddNamed(PEDString(PEDID._Point), Point)
            Catch ex As Exception
            End Try
            Try
                PED.RemoveNamed(PEDString(PEDID._Input_Output))
                PED.AddNamed(PEDString(PEDID._Input_Output), IO)
            Catch ex As Exception
            End Try

            dv.PlugExtraData_Set(hs) = PED
            hs.SaveEventsDevices()

            Return True
        Catch ex As Exception
            hs.WriteLogEx("Ocelot", "Exception setting HS device PED Info: " & ex.Message, COLOR_RED)
            Return False
        End Try
    End Function
    Public Function GetPEDInfo(ByRef dv As Scheduler.Classes.DeviceClass, _
                               ByRef Unit_Type As CPUXA_UnitTypes, _
                               ByRef Unit As Short, _
                               ByRef Point As Short, _
                               ByRef IO As enumIO) As Boolean
        Dim PED As clsPlugExtraData = Nothing
        Try
            Unit_Type = 0
            Unit = -1
            Point = -1
            IO = 0

            If dv Is Nothing Then Return False
            PED = dv.PlugExtraData_Get(hs)
            If PED Is Nothing Then Return False
            Try
                Unit_Type = PED.GetNamed(PEDString(PEDID._Type))
            Catch ex As Exception
                Return False
            End Try
            Try
                Unit = PED.GetNamed(PEDString(PEDID._Unit))
            Catch ex As Exception
                Return False
            End Try
            Try
                Point = PED.GetNamed(PEDString(PEDID._Point))
            Catch ex As Exception
                Return False
            End Try
            Select Case Unit_Type
                Case CPUXA_UnitTypes.Rly08XA, _
                     CPUXA_UnitTypes.Secu16, _
                     CPUXA_UnitTypes.Secu16i
                    Try
                        IO = PED.GetNamed(PEDString(PEDID._Input_Output))
                    Catch ex As Exception
                        Return False
                    End Try
            End Select

            If Not [Enum].IsDefined(GetType(CPUXA_UnitTypes), Unit_Type) Then Return False
            If Not [Enum].IsDefined(GetType(enumIO), IO) Then Return False

            Return True

        Catch ex As Exception
            hs.WriteLogEx("Ocelot", "Exception getting HS device PED Info: " & ex.Message, COLOR_RED)
        End Try
    End Function

    Public Sub RemoveOldPEDs(ByRef dv As Scheduler.Classes.DeviceClass)
        Dim PED As HomeSeerAPI.clsPlugExtraData = Nothing
        Try
            If dv Is Nothing Then Exit Sub
            PED = dv.PlugExtraData_Get(hs)
            Try
                PED.RemoveNamed("unit_type")
            Catch ex As Exception
            End Try
            Try
                PED.RemoveNamed("variable")
            Catch ex As Exception
            End Try
            Try
                PED.RemoveNamed("leopard")
            Catch ex As Exception
            End Try
            dv.PlugExtraData_Set(hs) = PED
            hs.SaveEventsDevices()
        Catch ex As Exception
        End Try
    End Sub
    Public Sub CheckConvertDevice(ByRef dv As Scheduler.Classes.DeviceClass, _
                                  ByRef PED As HomeSeerAPI.clsPlugExtraData)
        If dv Is Nothing Then Exit Sub
        If PED Is Nothing Then PED = New clsPlugExtraData

        Dim ptype As String = Nothing
        Dim IO As enumIO = 4
        Dim Unit_Type As CPUXA_UnitTypes
        Dim Changed As Boolean = False
        Dim obj As Object = Nothing
        Dim uif As String = ""

        Try
            hs.WriteLogEx("Ocelot", "Converting PED information on " & hs.DeviceName(dv.Ref(Nothing)), COLOR_PURPLE)

            'Diagnostic:	 type = unit_info
            'unit_info = 4!13!5
            ptype = PED.GetNamed(PEDString(PEDID._Type))

            If ptype IsNot Nothing Then
                If ptype = "unit_info" Then ' Apparently only RLY8XA did it this way, who knows why it is different.

                    uif = PED.GetNamed("unit_info")

                    Dim punit As Short = Convert.ToInt16(MidString(uif, 1, "!"))
                    Dim ppoint As Short = Convert.ToInt16(MidString(uif, 3, "!"))

                    If gUnits(punit - 1) = 0 Then Exit Sub
                    Unit_Type = gUnits(punit - 1)

                    If Not [Enum].IsDefined(GetType(CPUXA_UnitTypes), Unit_Type) Then Exit Sub

                    Dim pinput_output As String = ""
                    If Unit_Type = CPUXA_UnitTypes.Rly08XA Then
                        pinput_output = "output"
                    End If

                    If pinput_output = "input" Then
                        IO = enumIO._Input
                    ElseIf pinput_output = "output" Then
                        IO = enumIO._Output
                    Else
                        ' Must be one of these!
                        Exit Sub
                    End If

                    If UpdatePEDInfo(dv, Unit_Type, punit, ppoint, IO) Then
                        SetHSDeviceAddress(dv, Unit_Type, punit, ppoint, IO)
                        RemoveOldPEDs(dv)
                        Changed = True
                    End If
                ElseIf ptype = "leopard" Then
                    Unit_Type = CPUXA_UnitTypes.Leopard
                    If UpdatePEDInfo(dv, Unit_Type, Convert.ToInt16(0), Convert.ToInt16(0), enumIO._Input) Then
                        SetHSDeviceAddress(dv, Unit_Type, Convert.ToInt16(0), Convert.ToInt16(0), enumIO._Input)
                        RemoveOldPEDs(dv)
                        Changed = True
                    End If
                ElseIf ptype = "variable" Then
                    Unit_Type = CPUXA_UnitTypes.Variable
                    Dim ppoint As Short
                    Try
                        If PED.GetNamed("variable") Is Nothing Then Exit Sub
                        ppoint = Convert.ToInt16(PED.GetNamed("variable"))
                    Catch ex As Exception
                        Exit Sub
                    End Try
                    If UpdatePEDInfo(dv, Unit_Type, Convert.ToInt16(0), ppoint, enumIO._Output) Then
                        SetHSDeviceAddress(dv, Unit_Type, Convert.ToInt16(0), ppoint, enumIO._Output)
                        RemoveOldPEDs(dv)
                        Changed = True
                    End If
                End If
            End If

            If Not Changed Then
                ptype = Nothing
                ptype = PED.GetNamed("unit_type")
                Dim KeepGoing As Boolean = False
                If ptype IsNot Nothing Then
                    Select Case ptype
                        Case "6" ' CPUXA_UnitTypes.Secu16IR
                            Unit_Type = CPUXA_UnitTypes.Secu16IR
                            KeepGoing = True
                        Case "11" 'CPUXA_UnitTypes.Secu16
                            Unit_Type = CPUXA_UnitTypes.Secu16
                            KeepGoing = True
                        Case "12" 'CPUXA_UnitTypes.Secu16i
                            Unit_Type = CPUXA_UnitTypes.Secu16i
                            KeepGoing = True
                        Case "13" 'CPUXA_UnitTypes.Rly08XA
                            Unit_Type = CPUXA_UnitTypes.Rly08XA
                            KeepGoing = True
                    End Select
                End If
                If KeepGoing Then
                    Dim punit As Short = Convert.ToInt16(PED.GetNamed(PEDString(PEDID._Unit)))
                    If gUnits(punit - 1) = 0 Then Exit Sub
                    Unit_Type = gUnits(punit - 1)
                    If Not [Enum].IsDefined(GetType(CPUXA_UnitTypes), Unit_Type) Then Exit Sub
                    Dim ppoint As Short = Convert.ToInt16(PED.GetNamed(PEDString(PEDID._Point)))
                    Dim pinput_output As String = PED.GetNamed(PEDString(PEDID._Input_Output))
                    If pinput_output = "input" Then
                        IO = enumIO._Input
                    ElseIf pinput_output = "output" Then
                        IO = enumIO._Output
                    Else
                        ' Must be one of these!
                        Exit Sub
                    End If
                    If UpdatePEDInfo(dv, Unit_Type, punit, ppoint, IO) Then
                        SetHSDeviceAddress(dv, Unit_Type, punit, ppoint, IO)
                        RemoveOldPEDs(dv)
                        Changed = True
                    End If
                End If
            End If

            If Changed Then
                hs.SaveEventsDevices()
            End If

        Catch ex As Exception
        End Try

    End Sub
    Function Unit_Type_String(ByVal UT As CPUXA_UnitTypes) As String
        Select Case UT
            Case CPUXA_UnitTypes.Leopard
                Return "Leopard"
            Case CPUXA_UnitTypes.Variable
                Return "Variable"
            Case CPUXA_UnitTypes.Secu16
                Return "SECU16"
            Case CPUXA_UnitTypes.Secu16i
                Return "SECU16I"
            Case CPUXA_UnitTypes.Secu16IR
                Return "SECU16IR"
            Case CPUXA_UnitTypes.Rly08XA
                Return "RLY08XA"
            Case Else
                Return ""
        End Select
    End Function
    Function FindDevice(unit As Short, point As Short, variable As Short, _
                        leopard_device As Boolean, input As Boolean) As Scheduler.Classes.DeviceClass

        Dim dv As Scheduler.Classes.DeviceClass = Nothing
        Dim PED As HomeSeerAPI.clsPlugExtraData = Nothing
        Dim DE As Scheduler.Classes.clsDeviceEnumeration = Nothing
        'Dim ptype As String
        Dim obj As Object = Nothing

        Dim Find_Unit_Type As CPUXA_UnitTypes
        Dim Find_Unit As Short
        Dim Find_Point As Short
        Dim Find_IO As enumIO
        Dim sUnit As String = ""

        Try
            DE = hs.GetDeviceEnumerator

            If DE Is Nothing Then Return Nothing

            Do While Not DE.Finished
                dv = DE.GetNext
                If dv Is Nothing Then Continue Do
                If Not dv.Interface(Nothing) = IFACE_NAME Then Continue Do
                PED = dv.PlugExtraData_Get(Nothing)
                If PED Is Nothing Then Continue Do

                'Diagnostic:	 input_output, point, type, unit
                obj = Nothing
                obj = PED.GetNamed(PEDString(PEDID._Type))
                If obj IsNot Nothing AndAlso TypeOf obj Is String Then
                    CheckConvertDevice(dv, PED)
                    'PED = dv.PlugExtraData_Get(hs)
                End If

                If Not GetPEDInfo(dv, Find_Unit_Type, Find_Unit, Find_Point, Find_IO) Then Continue Do
                sUnit = Unit_Type_String(Find_Unit_Type)
                AddToCache(dv, Find_Unit, Find_Point, Find_Point, _
                           IIf(Find_Unit_Type = CPUXA_UnitTypes.Leopard, True, False), _
                           IIf(Find_IO = enumIO._Input, True, False), _
                           sUnit)

                If Find_Unit_Type = CPUXA_UnitTypes.Leopard Then
                    If leopard_device Then Return dv
                    Continue Do
                End If

                If Find_Unit_Type = CPUXA_UnitTypes.Variable Then
                    If variable >= 0 AndAlso variable = Find_Point Then Return dv
                    Continue Do
                End If

                If Find_Unit_Type = CPUXA_UnitTypes.Secu16IR Then
                    If Find_Unit = unit AndAlso Find_Point = point Then Return dv
                    Continue Do
                End If

                If Find_Unit = unit AndAlso Find_Point = point Then
                    If input Then
                        If Find_IO = enumIO._Input Then Return dv
                    Else
                        If Find_IO = enumIO._Output Then Return dv
                    End If
                    Continue Do
                End If

            Loop

        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error in FindDevice: " & ex.Message)
        End Try

        Return Nothing
    End Function


    Friend Sub AddToCache(ByRef dv As Scheduler.Classes.DeviceClass, _
                           ByVal Unit As Integer, _
                           ByVal Point As Integer, _
                           ByVal Var As Integer, _
                           ByVal LeopardDevice As Boolean, _
                           ByVal InputDev As Boolean, _
                           ByVal UnitType As String)
        If UnitType Is Nothing Then UnitType = ""

        Dim sKey As String = ""
        Dim Cache As DataCache = Nothing

        Try

            If dv Is Nothing Then Exit Sub

            sKey = "K" & Unit.ToString & "_" & Point.ToString & "_" & Var.ToString & "_" & LeopardDevice.ToString & "_" & InputDev.ToString

            If colDataCache Is Nothing Then
                colDataCache = New Collections.Concurrent.ConcurrentDictionary(Of String, DataCache)
            End If

            Try
                Cache = colDataCache.Item(sKey)
            Catch ex As Exception
                Cache = Nothing
            End Try
            If Cache IsNot Nothing Then Exit Sub

            Cache = New DataCache
            Cache.dv = dv
            Cache.dvRef = dv.Ref(Nothing)
            Cache.LastUpdate = Now
            Cache.DevValue = hs.DeviceValue(Cache.dvRef)
            Cache.OKToLog = Not dv.MISC_Check(hs, Enums.dvMISC.NO_LOG)
            Try
                If Not colDataCache.TryAdd(sKey, Cache) Then
                    WriteLog("Error, could not add a new Cache object to colDataCache.")
                End If
            Catch ex As Exception
            End Try
        Catch ex As Exception
            hs.WriteLogEx(IFACE_OTHER & " Error", "Exception adding HS device to Cache for unit " & Unit.ToString & ", I/O Point " & Point.ToString & " := " & ex.Message, COLOR_RED)
        End Try

    End Sub

    Friend Sub UpdateHSDeviceValue(ByVal Unit As Short, _
                                    ByVal Point As Short, _
                                    ByVal Var As Integer, _
                                    ByVal LeopardDevice As Boolean, _
                                    ByVal InputDev As Boolean, _
                                    ByVal NewValue As Integer, _
                                    ByVal UnitType As CPUXA_UnitTypes)
        'If UnitType Is Nothing Then UnitType = ""

        Dim dv As Scheduler.Classes.DeviceClass = Nothing
        Dim sKey As String = ""
        Dim Cache As DataCache = Nothing
        Dim Update As Boolean = False
        Dim Bad As Boolean = False
        Dim First As Boolean = False

        Try
            sKey = "K" & Unit.ToString & "_" & Point.ToString & "_" & Var.ToString & "_" & LeopardDevice.ToString & "_" & InputDev.ToString

            If colDataCache Is Nothing Then
                colDataCache = New Collections.Concurrent.ConcurrentDictionary(Of String, DataCache)
            End If

            Try
                Cache = colDataCache.Item(sKey)
            Catch ex As Exception
                Cache = Nothing
            End Try
            If Cache Is Nothing Then
                dv = FindDevice(Unit, Point, Var, LeopardDevice, InputDev)
                If dv Is Nothing Then
                    If gLogErrors Then
                        WriteLog("No HomeSeer device (A) for Key (Unit, Point, Var, Leopard, Input)=" & sKey)
                        Exit Sub
                    End If
                End If
                Cache = New DataCache
                Cache.dv = dv
                Cache.dvRef = dv.Ref(Nothing)
                Cache.LastUpdate = Now
                Cache.DevValue = hs.DeviceValue(Cache.dvRef)
                Cache.OKToLog = Not dv.MISC_Check(hs, Enums.dvMISC.NO_LOG)
                First = True
                Try
                    If Not colDataCache.TryAdd(sKey, Cache) Then
                        WriteLog("Error, could not add a new Cache object to colDataCache.")
                    End If
                Catch ex As Exception
                End Try
            Else
                dv = Cache.dv
                If dv Is Nothing Then Bad = True
                If Not Bad AndAlso Cache.dvRef <> dv.Ref(Nothing) Then Bad = True
                If Not Bad AndAlso hs.DeviceExistsRef(Cache.dvRef) = False Then Bad = True
                If Bad Then
                    Cache.LastUpdate = Date.MinValue
                    dv = FindDevice(Unit, Point, Var, LeopardDevice, InputDev)
                    If dv Is Nothing Then
                        Try
                            colDataCache.TryRemove(sKey, Cache)
                        Catch ex As Exception
                        End Try
                        If gLogErrors Then
                            WriteLog("No HomeSeer device (B) for Key (Unit, Point, Var, Leopard, Input)=" & sKey)
                            Exit Sub
                        End If
                    Else
                        Cache.dv = dv
                        Cache.dvRef = dv.Ref(Nothing)
                    End If
                End If
            End If
            If Cache Is Nothing Then Exit Sub
            If Cache.dv Is Nothing Then Exit Sub


            Try
                If Now.Subtract(Cache.LastUpdate).TotalSeconds > 299 Then
                    Cache.DevValue = hs.DeviceValue(Cache.dvRef)
                    Cache.OKToLog = Not Cache.dv.MISC_Check(hs, Enums.dvMISC.NO_LOG)
                    Cache.LastUpdate = Now
                    Update = True
                End If
            Catch ex As Exception
            End Try

            If NewValue <> Cache.DevValue Then
                Cache.DevValue = NewValue
                Cache.LastUpdate = Now
                Update = True
                If Cache.OKToLog Then
                    Try
                        Dim sNV As String = NewValue.ToString
                        If NewValue = 0 Then sNV = "Off"
                        If NewValue = 1 Then sNV = "On"
                        hs.WriteLogEx(IFACE_OTHER, LogPrefix(Unit, Point, Var, LeopardDevice, InputDev, UnitType) & _
                                      " [" & hs.DeviceName(Cache.dvRef) & "]" & _
                                      " is being set to " & sNV, COLOR_PURPLE2)
                    Catch ex As Exception
                    End Try
                End If
                hs.SetDeviceValueByRef(Cache.dvRef, NewValue, True)
                If gLogComms Then
                    WriteLog(UnitType & " Setting device " & LogPrefix(Unit, Point, Var, LeopardDevice, InputDev, UnitType) & " to " & NewValue.ToString)
                End If
            ElseIf First Then
                hs.SetDeviceValueByRef(Cache.dvRef, NewValue, False)
            End If

            Try
                If Update Then
                    If Not colDataCache.TryUpdate(sKey, Cache, Cache) Then
                        If gLogErrors Then
                            WriteLog("Update of cache for key failed: " & sKey)
                        End If
                    End If
                End If
            Catch ex As Exception
                If gLogErrors Then
                    WriteLog("Exception updating data cache entry for " & sKey)
                End If
            End Try

        Catch exa As Threading.ThreadAbortException
        Catch ex As Exception
            hs.WriteLogEx("Ocelot", "Exception updating HS device value for unit " & Unit.ToString & ", I/O Point " & Point.ToString & " := " & ex.Message, COLOR_RED)
        End Try
    End Sub

    Public Sub LoadAllSettings()
        Try
            gPollInterval = Convert.ToInt32(GetSettingLocal("hspi_ocelot", "Settings", "poll_interval", "500"))
        Catch ex As Exception
        End Try
        Try
            gConfigPollVars = Convert.ToBoolean(GetSettingLocal("hspi_ocelot", "Settings", "gConfigPollVars", "False"))
        Catch ex As Exception
        End Try
        Try
            gInvert = GetBooleanNumber("invert", "0")
        Catch ex As Exception
        End Try
        Try
            gLogErrors = GetBooleanNumber("log_errors", "1")
        Catch ex As Exception
        End Try
        Try
            gLogComms = GetBooleanNumber("log_comms", "0")
        Catch ex As Exception
        End Try
    End Sub
    Friend Function GetBooleanNumber(ByVal ConfigItem As String, ByVal DefValue As String) As Boolean
        Try
            Dim i As Short
            i = Convert.ToInt16(GetSettingLocal("hspi_ocelot", "Settings", ConfigItem, DefValue))
            Return IIf(i < 1, False, True)
        Catch ex As Exception
        End Try
        Try
            Return Convert.ToBoolean(GetSettingLocal("hspi_ocelot", "Settings", ConfigItem, DefValue))
        Catch ex As Exception
        End Try
    End Function

    Sub CreateUnitMappings()
        Dim i As Short
        Dim s As String
        Dim u() As String
        Dim t As String
        Dim um As umap
        Try

            gUnitMappings = New Collection
            s = GetSettingLocal("hspi_ocelot", "Settings", "mappings", "")
            If gLogComms Then
                WriteLog("Unit Mappings: " & s)
            End If
            gmappings = s
            If s <> "" Then
                u = Split(s, ",")
                For i = 0 To UBound(u)
                    t = u(i)
                    If t <> "" Then
                        um = New umap
                        um.unit = Convert.ToInt32(MidString(t, 1, "!"))
                        um.utype = Convert.ToInt32(MidString(t, 2, "!"))
                        gUnitMappings.Add(um)
                    End If
                Next
            End If
        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error in CreateUnitMappings: " & ex.Message)
        End Try
    End Sub

    Sub CreateLeopardDevice()
        Dim lref As Integer
        Dim dv As Scheduler.Classes.DeviceClass                             '*
        Dim DE As Scheduler.Classes.clsDeviceEnumeration                    '*
        Dim ped As clsPlugExtraData

        Dim Find_Unit_Type As CPUXA_UnitTypes
        Dim Find_Unit As Short
        Dim Find_Point As Short
        Dim Find_IO As enumIO

        Try

            DE = hs.GetDeviceEnumerator
            If DE Is Nothing Then Exit Sub

            Do While Not DE.Finished
                dv = DE.GetNext
                If dv Is Nothing Then Continue Do
                If Not dv.Interface(Nothing) = IFACE_NAME Then Continue Do
                ped = dv.PlugExtraData_Get(Nothing)
                If ped Is Nothing Then Continue Do
                ' SetHSDeviceAddress    UpdatePEDInfo   GetPEDInfo
                If Not GetPEDInfo(dv, Find_Unit_Type, Find_Unit, Find_Point, Find_IO) Then Continue Do

                If Find_Unit_Type = CPUXA_UnitTypes.Leopard Then
                    hs.WriteLogEx(IFACE_OTHER, "Leopard device is already created.", COLOR_ORANGE)
                    'System.Windows.Forms.Cursor.Current = System.Windows.Forms.Cursors.Default
                    Exit Sub
                End If
            Loop


            lref = hs.NewDeviceRef("Leopard Screen")
            dv = hs.GetDeviceByRef(lref)
            dv.Location(hs) = "Leopard"
            dv.Location2(hs) = "Ocelot"
            dv.Interface(hs) = IFACE_NAME
            dv.MISC_Set(hs, HomeSeerAPI.Enums.dvMISC.STATUS_ONLY)
            hs.SaveEventsDevices()
            UpdatePEDInfo(dv, CPUXA_UnitTypes.Leopard, 0, 0, enumIO._Output)
            dv.Device_Type_String(hs) = "Leopard Screen"
            Dim DTI As DeviceTypeInfo = Nothing
            DTI = dv.DeviceType_Get(hs)
            If DTI Is Nothing Then DTI = New DeviceTypeInfo
            DTI.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
            dv.DeviceType_Set(hs) = DTI
            SetHSDeviceAddress(dv, CPUXA_UnitTypes.Leopard, 0, 0, enumIO._Output)
            hs.SaveEventsDevices()


        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error in CreateLeopardDevice: " & ex.Message)
        End Try
    End Sub

    Sub CreateVarDevices()
        Dim i As Integer
        Dim dvref As Integer
        Dim dv As Scheduler.Classes.DeviceClass                             '*
        Dim DE As Scheduler.Classes.clsDeviceEnumeration                    '*
        Dim PED As clsPlugExtraData
        Dim Pair As VSPair

        Dim Find_Unit_Type As CPUXA_UnitTypes
        Dim Find_Unit As Short
        Dim Find_Point As Short
        Dim Find_IO As enumIO

        ' SetHSDeviceAddress    UpdatePEDInfo   GetPEDInfo
        Dim Found(128) As Boolean

        Try

            gHaveRealValues = False

            DE = hs.GetDeviceEnumerator
            If Not DE Is Nothing Then
                Do While Not DE.Finished
                    dv = DE.GetNext
                    If dv Is Nothing Then Continue Do
                    PED = dv.PlugExtraData_Get(Nothing)
                    If PED Is Nothing Then Continue Do
                    If Not GetPEDInfo(dv, Find_Unit_Type, Find_Unit, Find_Point, Find_IO) Then Continue Do
                    If Find_Unit_Type <> CPUXA_UnitTypes.Variable Then Continue Do
                    If Find_Point < 1 Then Continue Do
                    If Find_Point > 128 Then Continue Do
                    Found(Find_Point) = True
                Loop
            End If

            Dim Missing As Boolean = False
            For i = 1 To 128
                If Not Found(i) Then
                    Missing = True
                    Exit For
                End If
            Next

            If Not Missing Then
                hs.WriteLog(IFACE_OTHER, "Your variables appear to be created already. Please delete them before attempting to rebuild them.")
                Exit Sub
            End If
            ' create variables (ocelot supports 128)

            For i = 1 To 128
                If Found(i) Then Continue For
                dvref = hs.NewDeviceRef("Variable " & i.ToString)
                dv = hs.GetDeviceByRef(dvref)
                dv.Location(hs) = "Variables"
                dv.Location2(hs) = "Ocelot"
                dv.Interface(hs) = IFACE_NAME
                dv.Device_Type_String(hs) = "Interface Variable"
                Dim DTI As DeviceTypeInfo = Nothing
                DTI = dv.DeviceType_Get(hs)
                If DTI Is Nothing Then DTI = New DeviceTypeInfo
                DTI.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
                dv.DeviceType_Set(hs) = DTI
                hs.SaveEventsDevices()
                UpdatePEDInfo(dv, CPUXA_UnitTypes.Variable, 0, Convert.ToInt16(i), enumIO._Input)
                SetHSDeviceAddress(dv, CPUXA_UnitTypes.Variable, 0, Convert.ToInt16(i), enumIO._Input)

                ' add 2 buttons for ON and OFF
                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = 1
                Pair.Status = "On"
                Pair.Render = Enums.CAPIControlType.Button
                hs.DeviceVSP_AddPair(dvref, Pair)

                Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
                Pair.PairType = VSVGPairType.SingleValue
                Pair.Value = 0
                Pair.Status = "Off"
                Pair.Render = Enums.CAPIControlType.Button
                hs.DeviceVSP_AddPair(dvref, Pair)

            Next
            hs.SaveEventsDevices()

            CreateUnitMappings()
        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error in CreateVarDevices: " & ex.Message)
        End Try
    End Sub

    Friend Enum CPUXA_UnitTypes
        Secu16IR = 6
        Secu16 = 11
        Secu16i = 12
        Rly08XA = 13
        Leopard = 200
        Variable = 201
    End Enum
 
    Sub CreateIODevices(ByVal unit As Short)
        Dim j As Short
        Dim dvref As Integer
        Dim unit_type As CPUXA_UnitTypes 'Short
        Dim dv As Scheduler.Classes.DeviceClass = Nothing
        Dim DE As Scheduler.Classes.clsDeviceEnumeration = Nothing
        Dim PED As clsPlugExtraData = Nothing
        'Dim unit_info As String = ""
        Dim Pair As VSPair

        Dim Found_Unit As Short
        Dim Found_Unit_Type As CPUXA_UnitTypes
        Dim Found_Point As Short
        Dim Found_IO As enumIO

        Dim Found(16) As Boolean

        unit_type = gUnits(unit - 1)
        Select Case unit_type
            Case CPUXA_UnitTypes.Rly08XA, _
                 CPUXA_UnitTypes.Secu16, _
                 CPUXA_UnitTypes.Secu16i, _
                 CPUXA_UnitTypes.Secu16IR
            Case Else
                hs.WriteLogEx(IFACE_OTHER, "Could not create devices for unit " & unit.ToString & " because the Unit Type for that unit is not yet set.", COLOR_RED)
                Exit Sub
        End Select

        gHaveRealValues = False

        Try
            DE = hs.GetDeviceEnumerator
            If DE Is Nothing Then
                hs.WriteLog(IFACE_OTHER & " Error", "Could not get a device enumerator from HomeSeer - Unable to create devices.")
                Exit Sub
            End If

            Do While Not DE.Finished
                dv = DE.GetNext
                If dv Is Nothing Then Continue Do
                If dv.Interface(Nothing) <> IFACE_NAME Then Continue Do
                PED = dv.PlugExtraData_Get(Nothing)
                If PED Is Nothing Then Continue Do
                If Not GetPEDInfo(dv, Found_Unit_Type, Found_Unit, Found_Point, Found_IO) Then Continue Do
                If Found_Unit_Type = unit_type AndAlso Found_Unit = unit Then
                    If unit_type = CPUXA_UnitTypes.Secu16 AndAlso Found_IO = enumIO._Output Then
                        Found(Found_Point + 8) = True
                    Else
                        Found(Found_Point) = True
                    End If
                End If
            Loop

            Dim iEnd As Integer = 8
            Select Case unit_type
                Case CPUXA_UnitTypes.Secu16, _
                     CPUXA_UnitTypes.Secu16i, _
                     CPUXA_UnitTypes.Secu16IR
                    iEnd = 16
            End Select

            Dim Missing As Boolean = False
            For f As Integer = 1 To iEnd
                If Not Found(f) Then
                    Missing = True
                    Exit For
                End If
            Next

            If Not Missing Then
                hs.WriteLog(IFACE_OTHER, "Your devices appear to be created already. Please delete them before attempting to rebuild them.")
                Exit Sub
            End If


            Select Case unit_type
                Case CPUXA_UnitTypes.Secu16 ' secu16
                    gmappings = gmappings & unit.ToString & "!" & "11" & ","

                    ' inputs (8 inputs, and 8 outputs)
                    For j = 1 To 8
                        If Found(j) Then Continue For
                        dvref = hs.NewDeviceRef("Input " & j.ToString)
                        dv = hs.GetDeviceByRef(dvref)
                        dv.Location(hs) = "Secu16 Unit " & unit.ToString
                        dv.Location2(hs) = "Ocelot"
                        dv.Interface(hs) = IFACE_NAME
                        dv.MISC_Set(hs, HomeSeerAPI.Enums.dvMISC.STATUS_ONLY)
                        dv.Device_Type_String(hs) = "Secu16 Input"

                        Dim DTI As DeviceTypeInfo = Nothing
                        DTI = dv.DeviceType_Get(hs)
                        If DTI Is Nothing Then DTI = New DeviceTypeInfo
                        DTI.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
                        dv.DeviceType_Set(hs) = DTI

                        hs.SaveEventsDevices()

                        UpdatePEDInfo(dv, CPUXA_UnitTypes.Secu16, unit, j, enumIO._Input)
                        SetHSDeviceAddress(dv, CPUXA_UnitTypes.Secu16, unit, j, enumIO._Input)

                        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                        Pair.PairType = VSVGPairType.SingleValue
                        Pair.Value = 1
                        Pair.Status = "On"
                        Pair.Render = Enums.CAPIControlType.Button
                        hs.DeviceVSP_AddPair(dvref, Pair)

                        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                        Pair.PairType = VSVGPairType.SingleValue
                        Pair.Value = 0
                        Pair.Status = "Off"
                        Pair.Render = Enums.CAPIControlType.Button
                        hs.DeviceVSP_AddPair(dvref, Pair)
                    Next

                    ' outputs
                    For j = 1 To 8
                        If Found(j + 8) Then Continue For
                        dvref = hs.NewDeviceRef("Relay Output " & j.ToString)
                        dv = hs.GetDeviceByRef(dvref)
                        dv.Location(hs) = "Secu16 Unit " & unit.ToString
                        dv.Location2(hs) = "Ocelot"
                        dv.Interface(hs) = IFACE_NAME

                        dv.Device_Type_String(hs) = "Secu16 Output"
                        Dim DTI As DeviceTypeInfo = Nothing
                        DTI = dv.DeviceType_Get(hs)
                        If DTI Is Nothing Then DTI = New DeviceTypeInfo
                        DTI.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
                        dv.DeviceType_Set(hs) = DTI

                        hs.SaveEventsDevices()

                        UpdatePEDInfo(dv, CPUXA_UnitTypes.Secu16, unit, j, enumIO._Output)
                        SetHSDeviceAddress(dv, CPUXA_UnitTypes.Secu16, unit, j, enumIO._Output)

                        ' add 2 buttons for ON and OFF
                        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
                        Pair.PairType = VSVGPairType.SingleValue
                        Pair.Value = 1
                        Pair.Status = "On"
                        Pair.Render = Enums.CAPIControlType.Button
                        Pair.ControlUse = ePairControlUse._On
                        Pair.Render_Location.Row = 1
                        Pair.Render_Location.Column = 1
                        Pair.Render_Location.ColumnSpan = 1
                        hs.DeviceVSP_AddPair(dvref, Pair)

                        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
                        Pair.PairType = VSVGPairType.SingleValue
                        Pair.Value = 0
                        Pair.Status = "Off"
                        Pair.ControlUse = ePairControlUse._Off
                        Pair.Render_Location.Row = 1
                        Pair.Render_Location.Column = 2
                        Pair.Render_Location.ColumnSpan = 1
                        Pair.Render = Enums.CAPIControlType.Button
                        hs.DeviceVSP_AddPair(dvref, Pair)

                        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)

                        dv.Status_Support(hs) = True
                    Next

                Case CPUXA_UnitTypes.Secu16i '12 ' secu16i
                    gmappings = gmappings & unit.ToString & "!" & "12" & ","
                    ' inputs (16 inputs)
                    For j = 1 To 16
                        If Found(j) Then Continue For
                        dvref = hs.NewDeviceRef("Input " & j.ToString)
                        dv = hs.GetDeviceByRef(dvref)
                        dv.Location(hs) = "Secu16i Unit " & unit.ToString
                        dv.Location2(hs) = "Ocelot"
                        dv.Interface(hs) = IFACE_NAME

                        dv.Device_Type_String(hs) = "Secu16i Input"
                        Dim DTI As DeviceTypeInfo = Nothing
                        DTI = dv.DeviceType_Get(hs)
                        If DTI Is Nothing Then DTI = New DeviceTypeInfo
                        DTI.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
                        dv.DeviceType_Set(hs) = DTI

                        hs.SaveEventsDevices()

                        UpdatePEDInfo(dv, CPUXA_UnitTypes.Secu16i, unit, j, enumIO._Input)
                        SetHSDeviceAddress(dv, CPUXA_UnitTypes.Secu16i, unit, j, enumIO._Input)

                        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                        Pair.PairType = VSVGPairType.SingleValue
                        Pair.Value = 1
                        Pair.Status = "On"
                        Pair.Render = Enums.CAPIControlType.Button
                        hs.DeviceVSP_AddPair(dvref, Pair)

                        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Status)
                        Pair.PairType = VSVGPairType.SingleValue
                        Pair.Value = 0
                        Pair.Status = "Off"
                        Pair.Render = Enums.CAPIControlType.Button
                        hs.DeviceVSP_AddPair(dvref, Pair)

                        dv.Status_Support(hs) = True
                    Next

                Case CPUXA_UnitTypes.Rly08XA '13 ' rly08-xa
                    gmappings = gmappings & unit.ToString & "!" & "13" & ","
                    ' outputs
                    For j = 1 To 8
                        If Found(j) Then Continue For
                        dvref = hs.NewDeviceRef("Relay Output " & j.ToString)
                        dv = hs.GetDeviceByRef(dvref)
                        dv.Location(hs) = "Rly08 Unit " & unit.ToString
                        dv.Location2(hs) = "Ocelot"
                        dv.Interface(hs) = IFACE_NAME

                        dv.Device_Type_String(hs) = "RLY08 Output"
                        Dim DTI As DeviceTypeInfo = Nothing
                        DTI = dv.DeviceType_Get(hs)
                        If DTI Is Nothing Then DTI = New DeviceTypeInfo
                        DTI.Device_API = DeviceTypeInfo.eDeviceAPI.Plug_In
                        dv.DeviceType_Set(hs) = DTI

                        hs.SaveEventsDevices()

                        UpdatePEDInfo(dv, CPUXA_UnitTypes.Rly08XA, unit, j, enumIO._Output)
                        SetHSDeviceAddress(dv, CPUXA_UnitTypes.Rly08XA, unit, j, enumIO._Output)

                        ' add 2 buttons for ON and OFF
                        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
                        Pair.PairType = VSVGPairType.SingleValue
                        Pair.Value = 1
                        Pair.Status = "On"
                        Pair.Render = Enums.CAPIControlType.Button
                        Pair.ControlUse = ePairControlUse._On
                        Pair.Render_Location.Row = 1
                        Pair.Render_Location.Column = 1
                        Pair.Render_Location.ColumnSpan = 1
                        hs.DeviceVSP_AddPair(dvref, Pair)

                        Pair = New VSPair(HomeSeerAPI.ePairStatusControl.Both)
                        Pair.PairType = VSVGPairType.SingleValue
                        Pair.Value = 0
                        Pair.Status = "Off"
                        Pair.ControlUse = ePairControlUse._Off
                        Pair.Render_Location.Row = 1
                        Pair.Render_Location.Column = 2
                        Pair.Render_Location.ColumnSpan = 1
                        Pair.Render = Enums.CAPIControlType.Button
                        hs.DeviceVSP_AddPair(dvref, Pair)

                        dv.MISC_Set(hs, Enums.dvMISC.SHOW_VALUES)

                        dv.Status_Support(hs) = True
                    Next

            End Select

            hs.SaveEventsDevices()

            SaveSettingLocal("hspi_ocelot", "Settings", "mappings", gmappings)


            CreateUnitMappings()
        Catch ex As Exception
            hs.WriteLog(IFACE_OTHER, "Error in CreateIODevices: " & ex.Message)
        End Try
    End Sub


    Sub Wait(ByRef secs As Short)

        Thread.Sleep(secs * 1000)


    End Sub

    Sub waitms(ByRef ms As Short)
        If ms = 0 Then Exit Sub
        Thread.Sleep(ms)
    End Sub

    Function MidString(ByRef st As String, ByRef index As Integer, ByRef sep As String) As String
        ' return the string at the index
        ' strings are separated by the "sep" parameter
        Dim i As Integer
        Dim ind As Integer
        Dim start As Integer
        Dim s As String
        'On Error Resume Next


        start = 1
        ind = 1
        Do
            i = InStr(start, st, sep)
            If i <> 0 Then
                If ind = index Then
                    ' found the end of the requested string
                    MidString = Mid(st, start, i - start)
                    Exit Function
                Else
                    ind = ind + 1
                    start = i + Len(sep) ' next string starts after the seperator
                End If
            Else
                ' no more seperators, string not found
                ' see if last string is the correct one, it may not be followed by a seperator
                If ind = index Then
                    s = Mid(st, start)
                Else
                    s = ""
                End If
                MidString = s
                Exit Function
            End If
        Loop
    End Function

    Public Function char_to_int(ByRef msb As Short, ByRef lsb As Short) As Integer
        On Error Resume Next
        char_to_int = (Val(CStr(msb)) * 256) + Val(CStr(lsb))
    End Function

    Sub WriteLog(ByRef txt As String)
        If Not gLogComms Then Exit Sub
        Dim FileName As String = "\Ocelot.log"
        Dim FilePath As String = My.Application.Info.DirectoryPath & "\Debug Logs"
        Dim EMsg As String = ""
        Try
            If Not IO.Directory.Exists(FilePath) Then
                IO.Directory.CreateDirectory(FilePath)
            End If
        Catch ex As Exception
            EMsg = ex.Message
            GoTo error_Path
        End Try
        Try
            My.Computer.FileSystem.WriteAllText(FilePath & FileName, Now.ToString & vbTab & txt & vbCrLf, True)
        Catch ex As Exception
            EMsg = ex.Message
            GoTo error_File
        End Try
        Exit Sub

error_Path:
        hs.WriteLog("Error", "Ocelot plug-in: unable to write ocelog.log file to path " & FilePath & " " & EMsg)
        Exit Sub

error_File:
        hs.WriteLog("Error", "Ocelot plug-in: unable to write to ocelog.log file (maybe no disk space) " & EMsg)

    End Sub






    Function CalcCRC(ByVal buf() As Byte, ByVal length As Integer) As Integer

        Dim a As UShort
        Dim b As UShort
        Dim crcc As UShort
        Dim i As UShort
        Dim local As Integer
        Dim index As Integer
        Dim temp As Integer
        Dim temp1 As Integer


        crcc = &HFFFF 'seed value
        i = 0
        Do
            Try
                local = buf(index)
                temp = crcc Xor ((local << 8) And &HFFFF)
                a = temp And &HFFFF
                b = a
                temp = (((a And &HF000) >> 4) Xor b)
                a = temp And &HFFFF
                b = a
                temp1 = b
                temp = ((temp1 >> 8) And &HFF) + ((temp1 And &HFF) << 8)
                b = temp And &HFFFF
                a = a And &HFF00
                temp = (a >> 3) + ((a And 7) << 13)
                a = temp And &HFFFF
                b = b Xor a
                a = a >> 1
                temp = ((a >> 8) And &HFF) + ((a And &HFF) << 8)
                a = temp And &HFFFF
                temp = (a And &HF000) Xor b
                a = temp And &HFFFF
                crcc = a
                i += 1
                index += 1
            Catch ex As Exception
                Dim s As String = ex.Message
            End Try
        Loop While (i < length)
        Return crcc
    End Function

    Public Function GetVersionInfo(ByVal sFile As String) As String

        GetVersionInfo = "N/A"

        Dim myFileVersionInfo As FileVersionInfo = FileVersionInfo.GetVersionInfo(sFile)

        ' FileVersion is the file version from the file headers, not the numeric version info, so we do not want that.
        'GetVersionInfo = myFileVersionInfo.FileVersion
        GetVersionInfo = myFileVersionInfo.FileMajorPart.ToString & "." & _
          myFileVersionInfo.FileMinorPart.ToString & "." & _
          myFileVersionInfo.FileBuildPart.ToString & "." & _
          myFileVersionInfo.FilePrivatePart.ToString
    End Function

    Public Function GetResourceToFile(ByVal resname As String, ByVal filename As String) As String
        Dim thisExe As System.Reflection.Assembly
        Dim bteArray As Byte()
        Dim i As Integer
        Dim sReturn As String = ""
        Dim fs As FileStream

        sReturn = ""

        If File.Exists(filename) Then
            Try
                File.Delete(filename)
            Catch ex As Exception
                sReturn = "Deleting file " & filename & ", Error returned is " & ex.Message
                Return sReturn
            End Try
        End If

        Try
            fs = New FileStream(filename, FileMode.CreateNew)
        Catch ex As DirectoryNotFoundException
            sReturn = "Cannot create file, the containing folder does not exist: " & filename
            Return sReturn
        Catch ex As Exception
            sReturn = "Creating file stream for " & resname & ", Error returned is " & ex.Message
            Return sReturn
        End Try
        Dim w As New BinaryWriter(fs)
        thisExe = System.Reflection.Assembly.GetExecutingAssembly()
        Dim filest As System.IO.Stream
        Try
            filest = thisExe.GetManifestResourceStream(resname)
        Catch ex As Exception
            sReturn = "Creating resource file stream for " & resname & ", Error returned is " & ex.Message
            Return sReturn
        End Try

        ReDim bteArray(filest.Length)
        Try
            i = filest.Read(bteArray, 0, filest.Length)
        Catch ex As Exception
            sReturn = "Reading resource stream from " & resname & ", Error returned is " & ex.Message
            Return sReturn
        End Try
        Try
            fs.Write(bteArray, 0, bteArray.Length)
        Catch ex As Exception
            sReturn = "Writing file from stream resource " & resname & ", Error returned is " & ex.Message
            Return sReturn
        End Try

        w.Close()
        fs.Close()

        Return sReturn
    End Function

End Module
