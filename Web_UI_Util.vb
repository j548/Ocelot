Option Explicit On
Imports System.IO
Imports System.Text
Imports System.Threading
Imports System.Globalization
Imports VB = Microsoft.VisualBasic
Imports System.Web
Imports HomeSeerAPI

Friend Module Web_UI_Util


    ' HTTP constants
  
    Friend Const HTML_EndForm As String = "</form>" & vbCrLf
    Friend Const HTML_NewLine As String = "<br>" & vbCrLf
    Friend Const HTML_Line As String = "<hr>"
    Friend Const HTML_EndTable As String = "</table>" & vbCrLf
    Friend Const HTML_EndTableAlign As String = "</table></div>" & vbCrLf
    Friend Const HTML_EndRow As String = "</tr>" & vbCrLf
    Friend Const HTML_EndCell As String = "</td>"
    Friend Const HTML_StartBold As String = "<b>"
    Friend Const HTML_EndBold As String = "</b>"
    Friend Const HTML_EndFont As String = "</font>"


    Friend Const COLOR_WHITE As String = "#FFFFFF"
    Friend Const COLOR_KEWARE As String = "#0080C0"
    Friend Const COLOR_RED As String = "#FF0000"
    Friend Const COLOR_BLACK As String = "#000000"
    Friend Const COLOR_NAVY As String = "#000080"
    Friend Const COLOR_LT_BLUE As String = "#D9F2FF"
    Friend Const COLOR_LT_GRAY As String = "#E1E1E1"
    Friend Const COLOR_LT_PINK As String = "#FFB6C1"
    Friend Const COLOR_ORANGE As String = "#D58000"
    Friend Const COLOR_GREEN As String = "#008000"
    Friend Const COLOR_PURPLE As String = "#4B088A"
    Friend Const COLOR_PURPLE2 As String = "#9933FF"
    Friend Const COLOR_FUSCIA As String = "#FF33CC"
    Friend Const COLOR_BLUE2 As String = "#0066FF"
    Friend Const COLOR_GOLD As String = "#FFCC66"
    Friend Const COLOR_BLUE As String = "#0000FF"


    Friend Structure EventWebControlInfo
        Public Decoded As Boolean
        Public EventTriggerGroupID As Integer
        Public GroupID As Integer
        Public EvRef As Integer
        Public TriggerORActionID As Integer
        Public Name_or_ID As String
        Public Additional As String
    End Structure

    Friend gMultiConfigResult As String = ""


    <Serializable()> _
    Friend Enum HTML_Align
        RIGHT = 1
        LEFT = 2
        CENTER = 3
        JUSTIFY = 4
    End Enum

    <Serializable()> _
    Friend Enum HTML_VertAlign
        TOP = 1
        MIDDLE = 2
        BOTTOM = 3
        BASELINE = 4
    End Enum

    <Serializable()> _
    Friend Enum HTML_TableAlign
        RIGHT = 1
        LEFT = 2
        INHERIT = 3
    End Enum

#Region "    HTML Public Functions     "

    ' =====================================================================================================
    '                                  START OF HTML FUNCTIONS
    ' =====================================================================================================

    Friend Function AddNavTransLink(ByRef url As String, ByRef label As String, _
                                     Optional ByVal bSelected As Boolean = False, _
                                     Optional ByVal AltText As String = "") As String
        Dim st As String = ""
        On Error Resume Next

        If AltText Is Nothing Then
            AltText = " "
        End If
        '<input type="button" class="linkrowbutton" value="Webring" onClick="location.href='/misc/webring.asp'">
        If bSelected Then
            st = "<input type=""button"" class=""buttontrans ui-corner-all"" value=""" & label & """ alt=""" & AltText & """ onClick=""location.href='" & url & "'"" onmouseover=""this.className='buttontrans ui-corner-all';"" onmouseout=""this.className='buttontrans ui-corner-all';"">"
        Else
            st = "<input type=""button"" class=""buttontrans ui-corner-all"" value=""" & label & """ alt=""" & AltText & """ onClick=""location.href='" & url & "'"" onmouseover=""this.className='buttontrans ui-corner-all';"" onmouseout=""this.className='buttontrans ui-corner-all';"">"
        End If

        Return st
    End Function

    Friend Function HTML_Graphic(ByVal file As String, ByVal width As Integer, ByVal height As Integer) As String
        Dim st As String = ""
        If width = 0 Then
            st = "<img src=""" & file & """ />"
        Else
            st = "<img src=""" & file & """ height=""" & height.ToString & """" & "width=""" & width.ToString & """ />"
        End If
        Return st
    End Function

    Friend Function AddHidden(ByRef Name As String, ByRef Value As String) As Object
        Dim st As String = ""
        On Error Resume Next

        st = "<input type=""hidden"" name=""" & Name & """ value=""" & Value & """>"
        Return st
    End Function

    Friend Function AddHiddenWithID(ByRef Name As String, ByRef Value As String) As Object
        Dim st As String = ""
        On Error Resume Next

        st = "<input type=""hidden"" ID=""" & Name & """ name=""" & Name & """ value=""" & Value & """>"
        Return st
    End Function

    Friend Function HTML_StartFont(ByVal Color As String) As String
        Return HTML_Font(Color, "", "")
    End Function

    Friend Function HTML_Font(Optional ByVal color As String = "", Optional ByVal Style As String = "", Optional ByVal ClassName As String = "") As String
        Dim st As String = "<font "
        On Error Resume Next

        If color IsNot Nothing AndAlso Not String.IsNullOrEmpty(color.Trim) Then
            st &= "color='" & color.Trim & "' "
        End If
        If ClassName IsNot Nothing AndAlso Not String.IsNullOrEmpty(ClassName.Trim) Then
            st &= "class='" & ClassName.Trim & "' "
        End If
        If Style IsNot Nothing AndAlso Not String.IsNullOrEmpty(Style.Trim) Then
            st &= "style='" & Style.Trim & "' "
        End If
        st &= ">"
        Return st
    End Function

    Friend Function HTML_StartCell(ByRef Class_name As String, _
                                   ByRef colspan As Integer, _
                                   Optional ByVal align As HTML_Align = 0, _
                                   Optional ByVal nowrap As Boolean = False, _
                                   Optional ByVal RowHeight As Integer = 0, _
                                   Optional ByVal Style As String = "", _
                                   Optional ByVal VertAlign As HTML_VertAlign = 0, _
                                   Optional ByVal rowspan As Integer = 0) As String
        'Dim st As String = ""
        'Dim stalign As String = ""
        'Dim wrap As String = ""
        'Dim rheight As String = ""
        Dim s As String = "<td "
        Dim TotalStyle As String = ""

        On Error Resume Next

        'If RowHeight > 0 Then
        'rheight = " height=""" & RowHeight.ToString & """"
        'End If
        If Class_name IsNot Nothing AndAlso Not String.IsNullOrEmpty(Class_name) Then
            s &= "class='" & Class_name & "' "
        End If

        If Style IsNot Nothing AndAlso Not String.IsNullOrEmpty(Style.Trim) Then
            TotalStyle = Style.Trim
        End If

        If nowrap Then
            'wrap = " nowrap"
            If Not String.IsNullOrEmpty(TotalStyle.Trim) Then
                If TotalStyle.Trim.EndsWith(";") Then
                    ' Nothing to do.
                Else
                    TotalStyle &= ";"
                End If
            End If
            TotalStyle &= " white-space: nowrap;"
        End If

        If RowHeight > 0 Then
            If Not String.IsNullOrEmpty(TotalStyle.Trim) Then
                If TotalStyle.Trim.EndsWith(";") Then
                    ' Nothing to do.
                Else
                    TotalStyle &= ";"
                End If
            End If
            TotalStyle &= " height:" & RowHeight.ToString & "px;"
        End If

        If [Enum].IsDefined(GetType(HTML_Align), align) Then
            s &= "align='"
            Select Case align
                Case HTML_Align.LEFT
                    s &= "left"
                Case HTML_Align.CENTER
                    s &= "center"
                Case HTML_Align.RIGHT
                    s &= "right"
                Case HTML_Align.JUSTIFY
                    s &= "justify"
            End Select
            s &= "' "
        End If

        If [Enum].IsDefined(GetType(HTML_VertAlign), VertAlign) Then
            s &= "valign='"
            Select Case VertAlign
                Case HTML_VertAlign.TOP
                    s &= "top"
                Case HTML_VertAlign.MIDDLE
                    s &= "middle"
                Case HTML_VertAlign.BOTTOM
                    s &= "bottom"
                Case HTML_VertAlign.BASELINE
                    s &= "baseline"
            End Select
            s &= "' "
        End If

        If colspan > 0 Then
            s &= " colspan='" & colspan.ToString & "' "
        End If

        If rowspan > 0 Then
            s &= " rowspan='" & rowspan.ToString & "' "
        End If

        If Not String.IsNullOrEmpty(TotalStyle) Then
            s &= " style='" & TotalStyle & "' "
        End If

        s &= ">"

        Return s

    End Function
    

    Friend Function AddNBSP(ByVal amount As Short) As String
        Dim s As New StringBuilder
        Dim x As Short

        If amount < 1 Then
            Return ""
        End If
        For x = 1 To amount
            s.Append("&nbsp;")
        Next
        Return s.ToString

    End Function


    'Public Const HTML_StartRow As String = "<tr>"
    Friend Function HTML_StartRow(Optional ByVal ClassName As String = "", _
                                  Optional ByVal Style As String = "", _
                                  Optional ByVal Align As HTML_Align = 0, _
                                  Optional ByVal BackColor As String = "", _
                                  Optional ByVal VertAlign As HTML_VertAlign = 0) As String

        Dim s As String = "<tr "

        On Error Resume Next

        If ClassName IsNot Nothing AndAlso Not String.IsNullOrEmpty(ClassName.Trim) Then
            s &= "class='" & ClassName & "' "
        End If

        If Style IsNot Nothing AndAlso Not String.IsNullOrEmpty(Style) Then
            s &= "style='" & Style & "' "
        End If

        If [Enum].IsDefined(GetType(HTML_Align), Align) Then
            s &= "align='"
            Select Case Align
                Case HTML_Align.LEFT
                    s &= "left"
                Case HTML_Align.CENTER
                    s &= "center"
                Case HTML_Align.RIGHT
                    s &= "right"
                Case HTML_Align.JUSTIFY
                    s &= "justify"
            End Select
            s &= "' "
        End If

        If BackColor IsNot Nothing AndAlso Not String.IsNullOrEmpty(BackColor.Trim) Then
            s &= "bgcolor='" & BackColor & "' "
        End If

        If [Enum].IsDefined(GetType(HTML_VertAlign), VertAlign) Then
            s &= "valign='"
            Select Case VertAlign
                Case HTML_VertAlign.TOP
                    s &= "top"
                Case HTML_VertAlign.MIDDLE
                    s &= "middle"
                Case HTML_VertAlign.BOTTOM
                    s &= "bottom"
                Case HTML_VertAlign.BASELINE
                    s &= "baseline"
            End Select
            s &= "' "
        End If

        s &= ">"

        Return s

    End Function

    Friend Function HTML_StartTable(ByVal border As Integer, _
                                    Optional ByRef CellSpacing As Short = -1, _
                                    Optional ByRef TableWidthPercent As Short = -1, _
                                    Optional ByVal Style As String = "", _
                                    Optional ByVal ClassName As String = "", _
                                    Optional ByVal Align As HTML_TableAlign = 0, _
                                    Optional ByVal CellPadding As Short = -1, _
                                    Optional ByVal TableWidthPixels As Integer = -1) As String

        Dim s As String = "<table "
        Dim TotalStyle As String = ""

        On Error Resume Next

        If ClassName IsNot Nothing AndAlso Not String.IsNullOrEmpty(ClassName) Then
            s &= "class='" & ClassName & "' "
        End If

        If Style IsNot Nothing AndAlso Not String.IsNullOrEmpty(Style.Trim) Then
            TotalStyle = Style.Trim
        End If


        If [Enum].IsDefined(GetType(HTML_TableAlign), Align) Then
            If Not String.IsNullOrEmpty(TotalStyle.Trim) Then
                If TotalStyle.Trim.EndsWith(";") Then
                    ' Nothing to do.
                Else
                    TotalStyle &= ";"
                End If
            End If
            TotalStyle &= " float:"
            Select Case Align
                Case HTML_TableAlign.LEFT
                    TotalStyle &= "left"
                Case HTML_TableAlign.RIGHT
                    TotalStyle &= "right"
                Case HTML_TableAlign.INHERIT
                    TotalStyle &= "inherit"
            End Select
            TotalStyle &= ";"
        End If

        If CellPadding >= 0 Then
            s &= "cellpadding='" & CellPadding.ToString & "' "
        End If

        If CellSpacing >= 0 Then
            s &= "cellspacing='" & CellSpacing.ToString & "' "
        End If

        If TableWidthPixels >= 0 Then
            If TableWidthPercent >= 0 Then
                s &= "width='" & TableWidthPercent.ToString & "%' "
            Else
                s &= "width='" & TableWidthPixels.ToString & "' "
            End If
        Else
            If TableWidthPercent >= 0 Then
                s &= "width='" & TableWidthPercent.ToString & "%' "
            End If
        End If

        If Not String.IsNullOrEmpty(TotalStyle.Trim) Then
            s &= "style='" & TotalStyle & "'"
        End If

        s &= ">"

        Return s

    End Function

    Friend Function AddNavLinkPlugin(ByRef ref As String, ByRef label As String, _
       Optional ByVal bSelected As Boolean = False, _
       Optional ByVal AltText As String = "") As Object
        Dim st As String = ""
        On Error Resume Next

        If AltText Is Nothing Then
            AltText = " "
        End If
        '<input type="button" class="linkrowbutton" value="Webring" onClick="location.href='/misc/webring.asp'">
        If bSelected Then
            st = "<input type=""button"" class=""linkrowbuttonselectedplugin"" value=""" & label & """ alt=""" & AltText & """ onClick=""location.href='" & ref & "'"" onmouseover=""this.className='linkrowbuttonplugin';"" onmouseout=""this.className='linkrowbuttonselectedplugin';"">"
        Else
            st = "<input type=""button"" class=""linkrowbuttonplugin"" value=""" & label & """ alt=""" & AltText & """ onClick=""location.href='" & ref & "'"" onmouseover=""this.className='linkrowbuttonselectedplugin';"" onmouseout=""this.className='linkrowbuttonplugin';"">"
        End If

        AddNavLinkPlugin = st
    End Function

    Friend Function AddFuncLink(ByRef ref As String, ByRef label As String, Optional ByVal bSelected As Boolean = False) As Object
        Dim st As String = ""
        On Error Resume Next
        '<input type="button" class="linkrowbutton" value="Webring" onClick="location.href='/misc/webring.asp'">
        If bSelected Then
            st = "<input type=""button"" class=""functionrowbuttonselected"" value=""" & label & """ onClick=""location.href='" & ref & "'""  onmouseover=""this.className='functionrowbutton';"" onmouseout=""this.className='functionrowbuttonselected';"">"
        Else
            st = "<input type=""button"" class=""functionrowbutton"" value=""" & label & """ onClick=""location.href='" & ref & "'"" onmouseover=""this.className='functionrowbuttonselected';"" onmouseout=""this.className='functionrowbutton';"">"
        End If

        AddFuncLink = st
    End Function

    Friend Function FormCheckBox(ByRef label As String, ByRef name As String, ByRef Value As String, _
      ByRef checked As Boolean, Optional ByRef onChange As Boolean = False) As String
        Dim st As String = ""
        Dim chk As String = ""
        On Error Resume Next

        If checked Then
            chk = " checked "
        Else
            chk = ""
        End If
        If onChange Then
            st = "<input class=""formcheckbox"" type=""checkbox""" & chk & " name=""" & name & """ value=""" & Value & """ onClick=""submit();"" > " & label & vbCrLf
        Else
            st = "<input class=""formcheckbox"" type=""checkbox""" & chk & " name=""" & name & """ value=""" & Value & """ > " & label & vbCrLf
        End If
        Return st
    End Function

    Friend Function FormCheckBoxEx(ByRef label As String, ByRef name As String, ByRef Value As String, _
     ByRef checked As Boolean, Optional ByRef PrevNNL As Boolean = False, _
     Optional ByRef NNL As Boolean = False, Optional ByRef ColSpan As Integer = 0) As Object
        Dim st As New StringBuilder
        Dim chk As String = ""
        On Error Resume Next

        If ColSpan < 1 Then ColSpan = 1

        If PrevNNL Then
            st.Append(HTML_StartCell("", ColSpan, HTML_Align.LEFT, True))
        Else
            'st.Append(HTML_StartTable(0) & HTML_StartCell("", 1, ALIGN_LEFT, True))
            st.Append(HTML_EndRow & HTML_StartRow() & HTML_StartCell("", ColSpan, HTML_Align.LEFT, True))
        End If

        If checked Then
            chk = " checked "
        Else
            chk = ""
        End If

        st.Append("<input class=""FormCheckBoxEx"" type=""checkbox""" & chk & " name=""" & name & """ value=""" & Value & """ > " & label & vbCrLf)

        If NNL Then
            st.Append(HTML_EndCell)
        Else
            'st.Append(HTML_EndCell & HTML_EndTable)
            st.Append(HTML_EndCell & HTML_EndRow)
        End If

        FormCheckBoxEx = st.ToString
    End Function

    Friend Function FormRadio(ByRef label As String, ByRef name As String, ByRef Value As String, _
      ByRef checked As Boolean, Optional ByRef OnChange As Boolean = False) As Object
        Dim st As String = ""
        Dim chk As String = ""
        On Error Resume Next

        If checked Then
            chk = " checked "
        Else
            chk = ""
        End If
        If OnChange Then
            st = "<input class=""formradio"" type=""radio""" & chk & " name=""" & name & """ value=""" & Value & """ onClick=""submit();"" > " & label & vbCrLf
        Else
            st = "<input class=""formradio"" type=""radio""" & chk & " name=""" & name & """ value=""" & Value & """> " & label & vbCrLf
        End If

        FormRadio = st
    End Function

    Friend Function FormRadioEx(ByRef label As String, ByRef name As String, _
                                ByRef Value As String, ByRef checked As Boolean, _
                                Optional ByRef PrevNNL As Boolean = False, _
                                Optional ByRef NNL As Boolean = False, _
                                Optional ByRef ColSpan As Integer = 0) As Object
        Dim st As String = ""
        Dim chk As String = ""
        On Error Resume Next

        If ColSpan < 1 Then ColSpan = 1

        If PrevNNL Then
            st = HTML_StartCell("", ColSpan, HTML_Align.LEFT, True)
        Else
            st = HTML_StartTable(0) & HTML_StartCell("", ColSpan, HTML_Align.LEFT, True)
        End If

        If checked Then
            chk = " checked "
        Else
            chk = ""
        End If
        st = st & "<input class=""FormRadioEx"" type=""radio""" & chk & " name=""" & name & """ value=""" & Value & """> " & label & vbCrLf

        If NNL Then
            st = st & HTML_EndCell
        Else
            st = st & HTML_EndCell & HTML_EndTable
        End If

        FormRadioEx = st
    End Function

    Friend Function HTML_WrapSpan(ByVal id As String, ByVal wraptext As String, Optional ByVal bdisplay As Boolean = False, _
       Optional ByVal WithTable As Boolean = False) As String
        Dim st As New StringBuilder
        On Error Resume Next

        If WithTable Then
            st.Append("<td>" & vbCrLf)
        End If
        If bdisplay Then
            st.Append("<span  id=""" & id & """>" & vbCrLf)
        Else
            st.Append("<span style=""display:none;"" id=""" & id & """>" & vbCrLf)
        End If
        st.Append(wraptext)
        st.Append("</span>" & vbCrLf)
        If WithTable Then
            st.Append("</td>")
        End If

        Return st.ToString

    End Function

    Friend Function FormConfButton(ByRef Name As String, ByRef id As String, ByRef Source As String, _
     ByRef alt_title As String, Optional ByVal sel As Boolean = False, _
     Optional ByVal bHeight As Short = 33, Optional ByVal bWidth As Short = 100) As String
        Dim bup As New StringBuilder

        bup.Append("<input type=""image"" name=""")
        bup.Append(Name)
        bup.Append(""" border=""0"" id=""")
        bup.Append(id)
        bup.Append(""" src=""")
        If sel Then
            bup.Append(Source & "s.gif")
        Else
            bup.Append(Source & "n.gif")
        End If
        bup.Append(""" height=""" & bHeight.ToString & """ width=""" & bWidth.ToString & """ alt=""")
        bup.Append(alt_title)
        bup.Append(""" fp-style=""fp-btn: Glass Tab 1; fp-font-color-hover: #0000FF; ")
        bup.Append("fp-font-color-press: #FF0000; fp-transparent: 1"" fp-title=""")
        bup.Append(alt_title)
        bup.Append("" & vbCrLf)
        bup.Append("onmouseover = ""swapImg(1,0,/*id*/'" & id & "',/*url*/'" & Source & "h.gif')""" & vbCrLf)
        If sel Then
            bup.Append("onmouseout = ""swapImg(0,0,/*id*/'" & id & "',/*url*/'" & Source & "s.gif')""" & vbCrLf)
        Else
            bup.Append("onmouseout = ""swapImg(0,0,/*id*/'" & id & "',/*url*/'" & Source & "n.gif')""" & vbCrLf)
        End If
        If sel Then
            bup.Append("onmousedown = ""swapImg(1,0,/*id*/'" & id & "',/*url*/'" & Source & "n.gif')""" & vbCrLf)
        Else
            bup.Append("onmousedown = ""swapImg(1,0,/*id*/'" & id & "',/*url*/'" & Source & "s.gif')""" & vbCrLf)
        End If
        bup.Append("onmouseup=""swapImg(0,0,/*id*/'" & id & "',/*url*/'" & Source & "h.gif')"">" & vbCrLf)

        FormConfButton = bup.ToString

    End Function

#End Region

End Module
