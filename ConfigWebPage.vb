Imports System.Text
Imports System.Web
Imports Scheduler

<System.Reflection.ObfuscationAttribute(Exclude:=True, ApplyToMembers:=True)> _
Public Class ConfigWebPag
    Inherits clsPageBuilder

    Public Sub New(ByVal pagename As String)
        MyBase.New(pagename)
    End Sub

    Private Enum BoolResult
        _True = 1
        _False = 2
        _NotThere = 3
    End Enum
    Private Function GetBoolValue(ByVal sKey As String, _
                                  ByRef parts As Collections.Specialized.NameValueCollection) As BoolResult
        Try
            If parts(sKey).Trim.ToLower = "on" Then
                Return BoolResult._True
            ElseIf parts(sKey).Trim.ToLower = "checked" Then
                Return BoolResult._True
            ElseIf parts(sKey).Trim.ToLower = "off" Then
                Return BoolResult._False
            ElseIf parts(sKey).Trim.ToLower = "unchecked" Then
                Return BoolResult._False
            End If
        Catch ex As Exception
            Return BoolResult._NotThere
        End Try
    End Function


    Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String

        Dim parts As Collections.Specialized.NameValueCollection
        Dim sKey As String = ""
        Dim sValue As String = ""

        parts = HttpUtility.ParseQueryString(data)

        If parts("action") = "updatetime" Then
            'Me.divToUpdate.Add("Stats", UpdateStats)
            Return MyBase.postBackProc(page, data, user, userRights)
        End If

        Dim BR As BoolResult
        BR = GetBoolValue("InvertInputs", parts)
        If BR <> BoolResult._NotThere Then
            HSPI.ConfigInvertInput = IIf(BR = BoolResult._True, True, False)
        End If
        BR = GetBoolValue("LogErrors", parts)
        If BR <> BoolResult._NotThere Then
            HSPI.ConfigLogErrors = IIf(BR = BoolResult._True, True, False)
        End If
        BR = GetBoolValue("LogComms", parts)
        If BR <> BoolResult._NotThere Then
            HSPI.ConfigLogComms = IIf(BR = BoolResult._True, True, False)
        End If
        BR = GetBoolValue("PollVars", parts)
        If BR <> BoolResult._NotThere Then
            HSPI.ConfigPollVars = IIf(BR = BoolResult._True, True, False)
        End If

        Dim iVal As Integer
        Try
            iVal = Convert.ToInt32(parts("PollInterval"))
            If iVal > 250 Then
                HSPI.ConfigPollInterval = iVal
            End If
        Catch ex As Exception
            iVal = -1
        End Try

        'Dim dl As New clsJQuery.jqDropList("ModList", PageName, False)
        'dl.AddItem("Unit 0 Ocelot/CPU-XA/Leopard", "0", True)
        'Dim gUnits() As Byte = HSPI.ConfigUnits
        'Selected = IIf(SelectedUnit = (i + 1), True, False)
        'Butt = New clsJQuery.jqButton("CDevs", "Create Devices", PageName, True)

        'Butt = New clsJQuery.jqButton("DeleteAll", "Delete All Devices", PageName, True)
        'Butt = New clsJQuery.jqButton("CreateLeopard", "Create Leopard Devices", PageName, True)

        Dim CDevs As Boolean = False
        Dim DeleteAll As Boolean = False
        Dim CreateLeopard As Boolean = False
        Dim Unit As Integer = -1

        For i As Integer = 0 To parts.Count - 1
            sKey = parts.GetKey(i)
            sValue = parts(sKey).Trim
            If sKey Is Nothing Then Continue For
            If String.IsNullOrEmpty(sKey.Trim) Then Continue For
            If sKey.ToLower = "id" And sValue.Trim = "CDevs" Then CDevs = True
            If sKey.ToLower = "id" And sValue.Trim = "DeleteAll" Then DeleteAll = True
            If sKey.ToLower = "id" And sValue.Trim = "CreateLeopard" Then CreateLeopard = True
            If sKey = "ModList" Then
                Try
                    If Not Integer.TryParse(sValue, Unit) Then
                        Unit = -1
                    End If
                Catch ex As Exception
                    Unit = -1
                End Try
            End If
        Next
        If CDevs AndAlso Unit > -1 Then
            SelectedUnit = Unit
            If SelectedUnit = 0 Then
                CreateVarDevices()
            Else
                CreateIODevices(SelectedUnit)
            End If
            Me.pageCommands.Add("REFRESH", "TRUE")
        ElseIf CreateLeopard Then
            CreateLeopardDevice()
            Me.pageCommands.Add("REFRESH", "TRUE")
        ElseIf DeleteAll Then
            DeleteAllDevices()
            Me.pageCommands.Add("REFRESH", "TRUE")
        End If

        Return MyBase.postBackProc(page, data, user, userRights)

    End Function

    Private Function NavyWrap(ByVal Input As String) As String
        If Input Is Nothing Then Input = ""
        Return HTML_StartFont(COLOR_NAVY) & HTML_StartBold & Input & HTML_EndBold & HTML_EndFont
    End Function
 
    ' build and return the actual page
    ' a call to this must be added to Web_Server.vb in ServerRead_BigCaseStatement
    ' there are 2 options here:
    ' 1: Generate the full page and return it. Call BuildRealPageTemplate from Web_Server
    ' 2: Delay load the page. Use this if the page will take some time to build. Call BuildPageTemplate from Web_Server, then return the real page in the PostBack
    ' this example uses the delay load
    Private SelectedUnit As Short = 0
    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String
        Dim stb As New StringBuilder

        Try

            Me.reset()

            Dim TB As clsJQuery.jqTextBox = Nothing
            Dim CB As clsJQuery.jqCheckBox = Nothing
            Dim LB As clsJQuery.jqListBox = Nothing
            Dim Butt As clsJQuery.jqButton = Nothing

            ' add the normal title
            Me.AddHeader(hs.GetPageHeader(pageName, "AppDig Ocelot Plug-In Configuration", "", "", False, False))

            stb.Append(clsPageBuilder.DivStart(pageName, ""))

            ' this demonstrates a ajax timer that will call back to this page and update the time
            'Me.RefreshIntervalMilliSeconds = 1000          ' # of seconds between callbacks
            ' add a callback post string, this is what will be posted when the page timer expires
            'stb.Append(Me.AddAjaxHandlerPost("action=updatetime", pageName))

            ' specific page starts here
            stb.Append(clsPageBuilder.FormStart("ConfigID", "ConfigName", "post") & vbCrLf)

            stb.Append(HTML_StartTable(2, , , , , HTML_TableAlign.LEFT, , 600))

            stb.Append(HTML_StartRow())
            stb.Append(HTML_StartCell("", 2, HTML_Align.RIGHT, True))
            stb.Append("Polling Interval (ms):")
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True))
            TB = New clsJQuery.jqTextBox("PollInterval", "number", HSPI.ConfigPollInterval.ToString, pageName, 2, True)
            stb.Append(TB.Build)
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)

            stb.Append(HTML_StartRow() & vbCrLf)
            stb.Append(HTML_StartCell("", 2, HTML_Align.RIGHT, True))
            stb.Append("Invert Inputs:")
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True))
            CB = New clsJQuery.jqCheckBox("InvertInputs", "", pageName, True, False)
            CB.checked = HSPI.ConfigInvertInput
            stb.Append(CB.Build)
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)

            stb.Append(HTML_StartRow() & vbCrLf)
            stb.Append(HTML_StartCell("", 2, HTML_Align.RIGHT, False))
            stb.Append("Log Checksum Errors and Serial Timeouts:")
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True))
            CB = New clsJQuery.jqCheckBox("LogErrors", "", pageName, True, False)
            CB.checked = HSPI.ConfigLogErrors
            stb.Append(CB.Build)
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)

            stb.Append(HTML_StartRow() & vbCrLf)
            stb.Append(HTML_StartCell("", 2, HTML_Align.RIGHT, False))
            stb.Append("Log Communications to Ocelot.log:")
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True))
            CB = New clsJQuery.jqCheckBox("LogComms", "", pageName, True, False)
            CB.checked = HSPI.ConfigLogComms
            stb.Append(CB.Build)
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)

            stb.Append(HTML_StartRow() & vbCrLf)
            stb.Append(HTML_StartCell("", 2, HTML_Align.RIGHT, False))
            stb.Append("Poll Variables for Changes:")
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True))
            CB = New clsJQuery.jqCheckBox("PollVars", "", pageName, True, False)
            CB.checked = HSPI.ConfigPollVars
            stb.Append(CB.Build)
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)

            stb.Append(HTML_StartRow() & vbCrLf)
            stb.Append(HTML_StartCell("", 3, HTML_Align.CENTER, False))
            stb.Append("Select a module from the list and click the ")
            stb.Append(NavyWrap("Create Devices"))
            stb.Append(" button to have HomeSeer devices created")
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)

            stb.Append(HTML_StartRow() & vbCrLf)
            stb.Append(HTML_StartCell("", 2, HTML_Align.RIGHT, False))
            Dim dl As New clsJQuery.jqDropList("ModList", pageName, False)
            dl.AddItem("Unit 0 Ocelot/CPU-XA/Leopard", "0", True)
            Dim SysUnits() As Byte = HSPI.ConfigUnits
            Dim Unit As String = ""
            Dim Selected As Boolean = False
            If SysUnits IsNot Nothing AndAlso SysUnits.Length > 0 Then
                For i As Integer = 0 To SysUnits.Length - 1
                    Unit = (i + 1).ToString
                    Selected = IIf(SelectedUnit = (i + 1), True, False)
                    Select Case SysUnits(i)
                        Case 6
                            dl.AddItem("Unit " & Unit & " SECU16-IR", Unit, Selected)
                        Case 11
                            dl.AddItem("Unit " & Unit & " SECU16", Unit, Selected)
                        Case 12
                            dl.AddItem("Unit " & Unit & " SECU16I", Unit, Selected)
                        Case 13
                            dl.AddItem("Unit " & Unit & " RLY08-XA", Unit, Selected)
                    End Select
                Next
            End If
            stb.Append(dl.Build)
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_StartCell("", 1, HTML_Align.LEFT, True))
            Butt = New clsJQuery.jqButton("CDevs", "Create Devices", pageName, True)
            stb.Append(Butt.Build)
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)

            stb.Append(HTML_StartRow() & vbCrLf)
            stb.Append(HTML_StartCell("", 3, HTML_Align.CENTER, False))
            Butt = New clsJQuery.jqButton("DeleteAll", "Delete All Devices", pageName, True)
            stb.Append(Butt.Build)
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)

            stb.Append(HTML_StartRow() & vbCrLf)
            stb.Append(HTML_StartCell("", 3, HTML_Align.CENTER, False))
            Butt = New clsJQuery.jqButton("CreateLeopard", "Create Leopard Devices", pageName, True)
            stb.Append(Butt.Build)
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)

            stb.Append(HTML_EndTable & vbCrLf)



            stb.Append(clsPageBuilder.FormEnd & vbCrLf)

            stb.Append(HTML_NewLine & vbCrLf)
            stb.Append(HTML_StartTable(0, , 100))
            stb.Append(HTML_StartRow() & vbCrLf)
            stb.Append(HTML_StartCell("", 10))
            stb.Append("&nbsp;")
            stb.Append(HTML_EndCell & vbCrLf)
            stb.Append(HTML_EndRow & vbCrLf)
            stb.Append(HTML_EndTable & vbCrLf)
            stb.Append(HTML_NewLine & vbCrLf)


            stb.Append(clsPageBuilder.DivEnd)



            ' add the body html to the page
            Me.AddBody(stb.ToString)

            ' return the full page
            'Me.AddFooter(hs.GetPageFooter)
            Me.suppressDefaultFooter = False

            Return Me.BuildPage()
        Catch ex As Exception
            'WriteMon("Error", "Building page: " & ex.Message)
            Return "Error building Ocelot configuration page:=" & ex.Message
        End Try
    End Function




End Class

