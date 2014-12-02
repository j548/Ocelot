<%@ Page Language="VB" %>

<!DOCTYPE html PUBLIC "-//W3C//DTD XHTML 1.0 Transitional//EN" "http://www.w3.org/TR/xhtml1/DTD/xhtml1-transitional.dtd">

<script runat="server">
    
    'Dim i As Integer
    Dim hs As Object
    Dim pi As Object = Nothing
    Dim a As System.Reflection.Assembly
    Protected Sub Page_Load(ByVal sender As Object, ByVal e As System.EventArgs)
        Try
            hs = Context.Items("Content")
#If 0 Then
            Try
                a = System.Reflection.Assembly.LoadFrom(hs.GetAppPath & "\HSPI_OCELOT.DLL")
                a.CreateInstance("HSPI_OCELOT.HSPI", False, Reflection.BindingFlags.CreateInstance, Nothing, Nothing, Nothing, Nothing)
            Catch ex As Exception
                Response.Write("Error loading assembly: " & ex.Message)
            End Try
#End If
            pi = hs.Plugin("Applied Digital Ocelot")
            If pi Is Nothing Then
                Response.Write("Error getting a reference to the plug-in.  Is it loaded and enabled?")
                Exit Sub
            End If
            'Response.Write("Name: " & pi.name)
            If Not IsPostBack Then
                LoadItems()
            End If
            
        Catch ex As Exception
            Response.Write("Error: " & ex.Message)
        End Try
        
    End Sub
    
    Private Sub LoadItems()
        ListBox1.Items.Add("Unit 0 Ocelot/CPU-XA/Leopard")
        Dim gUnits() As Byte = pi.ConfigUnits
        For i As Integer = 0 To 127
            Select Case gUnits(i)
                Case 6
                    ListBox1.Items.Add("Unit " & Trim(Str(i + 1)) & " SECU16-IR")
                Case 11
                    ListBox1.Items.Add("Unit " & Trim(Str(i + 1)) & " SECU16")
                Case 12
                    ListBox1.Items.Add("Unit " & Trim(Str(i + 1)) & " SECU16I")
                Case 13
                    ListBox1.Items.Add("Unit " & Trim(Str(i + 1)) & " RLY08-XA")
            End Select
        Next
        'TxtBoxPort.Text = CStr(pi.ConfigComPort)
        TxtBoxPoll.Text = pi.ConfigPollInterval.ToString
        CheckInvert.Checked = pi.ConfigInvertInput
        CheckLogErrors.Checked = pi.ConfigLogErrors
        CheckLogComms.Checked = pi.ConfigLogComms
        CheckPollVars.Checked = pi.ConfigPollVars
    End Sub
    
    Private Function GetHeadContent() As String
        Try
            
            Return hs.GetPageHeader("Ocelot Configuration", "", "", False, False, True, False, False)
        Catch ex As Exception
        End Try
        Return ""
    End Function
    
    Private Function GetBodyContent() As String
        Try
            Return hs.GetPageHeader("Ocelot Configuration", "", "", False, False, False, True, False)
        Catch ex As Exception
        End Try
        Return ""
    End Function
    
    Protected Sub ButSave_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        'pi.ConfigComPort = CInt(TxtBoxPort.Text)
        pi.ConfigPollInterval = Convert.ToInt32(TxtBoxPoll.Text)
        pi.ConfigLogComms = CheckLogComms.Checked
        pi.ConfigLogErrors = CheckLogErrors.Checked
        pi.ConfigInvertInput = CheckInvert.Checked
        pi.ConfigPollVars = CheckPollVars.Checked
    End Sub

    Protected Sub ButCreateDevices_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        Dim i As Integer
        Dim s As String = ""
        Dim parts() As String
        
        For i = 0 To ListBox1.Items.Count - 1
            If ListBox1.Items(i).Selected Then
                s = ListBox1.Items(i).Text
                Exit For
            End If
        Next
		
        parts = Split(s, " ")
        If CInt(parts(1)) = 0 Then
            pi.ConfigCreateVarDevices()
        Else
            pi.ConfigCreateIODevices(CInt(parts(1)))
        End If
    End Sub

    Protected Sub ButCreateLeopardDevice_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pi.ConfigCreateLeopardDevice()
    End Sub

    Protected Sub ButDeleteAllDevices_Click(ByVal sender As Object, ByVal e As System.EventArgs)
        pi.ConfigDeleteAllDevices()
    End Sub
</script>

<html xmlns="http://www.w3.org/1999/xhtml" >
<head runat="server">
    <title>Ocelot Configuration</title>
    <%response.write(GetHeadContent())%>
</head>
<body>
    <% response.write(GetBodyContent()) %>
    <form id="form1" runat="server">
    <div>
        <table style="background-color: whitesmoke; border-right: black thin solid; border-top: black thin solid; border-left: black thin solid; border-bottom: black thin solid;" cellpadding="0" cellspacing="0">
            <tr>
                <td style="width: 271px; background-color: gainsboro;">
                    <strong>
                    Poll Interval</strong></td>
                <td style="width: 55px; background-color: gainsboro;">
                    <asp:TextBox ID="TxtBoxPoll" runat="server" Width="50px"></asp:TextBox></td>
            </tr>
            <tr>
                <td style="width: 271px; background-color: whitesmoke;">
                    <strong>
                    Invert Inputs</strong></td>
                <td style="width: 55px">
                    <asp:CheckBox ID="CheckInvert" runat="server" /></td>
            </tr>
            <tr>
                <td style="width: 271px; background-color: gainsboro;">
                    <strong>
                    Log checksum errors and serial timeouts</strong></td>
                <td style="width: 55px; background-color: gainsboro;">
                    <asp:CheckBox ID="CheckLogErrors" runat="server" /></td>
            </tr>
            <tr>
                <td style="width: 271px; background-color: whitesmoke;">
                    <strong>
                    Log communications to ocelot.log</strong></td>
                <td style="width: 55px">
                    <asp:CheckBox ID="CheckLogComms" runat="server" /></td>
            </tr>
            <tr>
                <td style="width: 271px; background-color: gainsboro; height: 21px;">
                    <strong>Poll variables for changes</strong></td>
                <td style="width: 55px; background-color: gainsboro; height: 21px;"><asp:CheckBox ID="CheckPollVars" runat="server" /></td>
            </tr>
            <tr>
                <td style="width: 271px; height: 21px; background-color: whitesmoke">
                </td>
                <td style="width: 55px; height: 21px; background-color: whitesmoke">
                </td>
            </tr>
            <tr>
                <td style="background-color: whitesmoke;" align="center" colspan="2">
                    <asp:ListBox ID="ListBox1" runat="server"></asp:ListBox><br />
                    <br />
                    <asp:Button ID="ButCreateDevices" runat="server" Text="Create Devices" OnClick="ButCreateDevices_Click" /><br />
                    <br />
                    Select a module from list then click the <em>Create Devices</em> button to have HomeSeer devices created<br />
                    &nbsp;</td>
            </tr>
            <tr>
                <td style="background-color: whitesmoke; text-align: center;" colspan="2">
                    <asp:Button ID="ButDeleteAllDevices" runat="server" Text="Delete all devices for all units" OnClick="ButDeleteAllDevices_Click" /><br />
                    &nbsp;</td>
            </tr>
            <tr>
                <td colspan="2" style="background-color: whitesmoke; text-align: center">
                    <asp:Button ID="ButCreateLeopardDevice" runat="server" Text="Create Leopard Device" OnClick="ButCreateLeopardDevice_Click" /><br />
                    <hr />
                </td>
            </tr>
            <tr>
                <td colspan="2" style="background-color: whitesmoke; text-align: center">
                    <asp:Button ID="ButSave" runat="server" OnClick="ButSave_Click" Text="Save" /></td>
            </tr>
        </table>
    
    
    </div>
    </form>
</body>
</html>
