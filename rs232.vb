Imports System.Threading

Public Class rs232
    Private rsport As System.IO.Ports.SerialPort
    Private _portnum As Integer
    Private _settings As String
    Private _RThreshold As Integer
    Private _InputLen As Integer
    Private _InputMode As input_modes_values
    Public Enum input_modes_values
        INPUT_BINARY
        INPUT_STRING
    End Enum
    Private ReadThread As Thread
    Enum init_values
        INIT_WAITING
        INIT_OK
        INIT_FAILED_OPEN
        INIT_FAILED_CREATE
    End Enum

    Private InitializedStatus As init_values
    Private LastError As String
    Private RBuffer As New Queue
    Private WriteQ As New Queue

    Public CommEvent As comm_event_types

    Public Event OnComm(ByVal event_type As comm_event_types)
    Public Enum comm_event_types
        comEvSend = 1
        comEvReceive = 2
        comEvCTS = 3
        comEvDSR = 4
        comEvCD = 5
        comEvRing = 6
        comEvEOF = 7
    End Enum

    Private Sub ReadThreadProc()
        Try
            rsport = New System.IO.Ports.SerialPort
        Catch ex As Exception
            InitializedStatus = init_values.INIT_FAILED_CREATE
            LastError = ex.Message
            Exit Sub
        End Try

        Try
            Dim parts() As String
            parts = Split(_settings, ",")
            rsport.BaudRate = CInt(Val(parts(0).Trim))
            Select Case Left(parts(1).Trim, 1).ToUpper
                Case "N" : rsport.Parity = IO.Ports.Parity.None
                Case "O" : rsport.Parity = IO.Ports.Parity.Odd
                Case "E" : rsport.Parity = IO.Ports.Parity.Even
                Case "M" : rsport.Parity = IO.Ports.Parity.Mark
                Case "S" : rsport.Parity = IO.Ports.Parity.Space
            End Select
            rsport.DataBits = CInt(Val(parts(2).Trim))
            Select Case parts(3).Trim.ToUpper
                Case "1"
                    rsport.StopBits = IO.Ports.StopBits.One
                Case "1.5"
                    rsport.StopBits = IO.Ports.StopBits.OnePointFive
                Case "2"
                    rsport.StopBits = IO.Ports.StopBits.Two
                Case "0", "N"
                    rsport.StopBits = IO.Ports.StopBits.None
            End Select
            rsport.PortName = "COM" & _portnum.ToString
            rsport.Open()
        Catch ex As Exception
            InitializedStatus = init_values.INIT_FAILED_OPEN
            LastError = ex.Message
            Exit Sub
        End Try

        InitializedStatus = init_values.INIT_OK

        Try
            Do
                Dim cnt As Integer
                Dim buf() As Byte
                cnt = rsport.BytesToRead()
                If cnt > 0 Then
                    ReDim buf(cnt - 1)
                    rsport.Read(buf, 0, cnt)
                    Dim i As Integer
                    For i = 0 To cnt - 1
                        RBuffer.Enqueue(buf(i))
                    Next
                    CommEvent = comm_event_types.comEvReceive
                    RaiseEvent OnComm(comm_event_types.comEvReceive)
                End If
                If WriteQ.Count > 0 Then
                    Dim wbuf() As Byte
                    wbuf = WriteQ.Dequeue
                    rsport.Write(wbuf, 0, wbuf.Length)
                End If
                Thread.Sleep(10)
            Loop
        Catch ex As ThreadAbortException
            Try
                rsport.Close()
            Catch ex1 As Exception
            End Try
        Catch ex As Exception
            Dim s As String = ex.Message
        End Try
    End Sub

    Sub New()
        
    End Sub

    Public Sub DiscardInBuffer()
        rsport.DiscardInBuffer()
    End Sub

    Public Property Input() As Byte()
        Get
            Dim buf() As Byte
            Dim i As Integer = RBuffer.Count
            ReDim buf(_InputLen)
            Dim j As Integer
            For j = 0 To _InputLen - 1
                Try
                    buf(j) = RBuffer.Dequeue
                Catch ex As Exception
                End Try
            Next
            Input = buf
        End Get
        Set(ByVal value As Byte())
        End Set
    End Property

    Public Property InBufferCount() As Integer
        Get
            InBufferCount = RBuffer.Count
        End Get
        Set(ByVal value As Integer)
            RBuffer.Clear()
        End Set
    End Property

    Public Property InBufferSize() As Integer
        Get
            Return rsport.ReadBufferSize
        End Get
        Set(ByVal value As Integer)
            rsport.ReadBufferSize = value
        End Set
    End Property

    Public Property Handshaking() As System.IO.Ports.Handshake
        Get
            Return rsport.Handshake
        End Get
        Set(ByVal value As System.IO.Ports.Handshake)
            rsport.Handshake = value
        End Set
    End Property

    Public WriteOnly Property Output() As Byte()
        Set(ByVal value() As Byte)
            WriteQ.Enqueue(value)
        End Set
    End Property

    Public Property InputMode() As input_modes_values
        Get
            InputMode = _InputMode
        End Get
        Set(ByVal value As input_modes_values)
            _InputMode = value
        End Set
    End Property

    Public Property InputLen() As Integer
        Get
            InputLen = _InputLen
        End Get
        Set(ByVal value As Integer)
            _InputLen = value
        End Set
    End Property

    Public Property RThreshold() As Integer
        Get
            RThreshold = _RThreshold
        End Get
        Set(ByVal value As Integer)
            _RThreshold = value
        End Set
    End Property

    Public Property Settings() As String
        Get
            Settings = _settings
        End Get
        Set(ByVal value As String)
            _settings = value
        End Set
    End Property

    Public Property CommPort() As Integer
        Get
            CommPort = _portnum
        End Get
        Set(ByVal value As Integer)
            _portnum = value
        End Set
    End Property

    Public Property PortOpen() As Boolean
        Get

        End Get
        Set(ByVal value As Boolean)
            If value Then
                Try
                    InitializedStatus = init_values.INIT_WAITING
                    ReadThread = New Thread(AddressOf ReadThreadProc)
                    ReadThread.Start()
                    Dim tm As Date = Now
                    Do
                        Thread.Sleep(10)
                    Loop While InitializedStatus = init_values.INIT_WAITING And DateDiff(DateInterval.Second, tm, Now) < 15
                    If InitializedStatus <> init_values.INIT_OK Then
                        Throw New Exception(LastError)
                    End If
                Catch ex As Exception
                    Throw New Exception("Unable to start com port read thread: " & ex.Message)
                End Try
            Else
                Try
                    ReadThread.Abort()
                Catch ex As Exception
                End Try
            End If
        End Set
    End Property
End Class
