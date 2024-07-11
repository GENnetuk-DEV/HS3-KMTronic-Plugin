Imports System.Text
Imports System.Web
Imports Scheduler
Imports HomeSeerAPI

Public Class web_config
    Inherits clsPageBuilder
    Dim TimerEnabled As Boolean

    Public Sub New(ByVal pagename As String)
        MyBase.New(pagename)
    End Sub

    Public Overrides Function postBackProc(page As String, data As String, user As String, userRights As Integer) As String
        Dim parts As Collections.Specialized.NameValueCollection = HttpUtility.ParseQueryString(data)


        Debug(2, "Web", "Processing Part [" + parts("id").ToString + "]")

        'Gets the control that caused the postback and handles accordingly
        Select Case parts("id")

            Case "oTextbox1"
                'This gets the value that was entered in the specified textbox
                Dim message As String = parts("Textbox1")

                '... posts it to the page
                PostMessage("Cmessage", message)

                '... and rebuilds the viewed textbox to contain the message
                BuildTextBox("Textbox1", True, message)

            Case "oScale"
                GlobalVariables.Scale = parts("Scale")
                PostMessage("Dmessage", "Scale Changed to " + GlobalVariables.Scale)
                Debug(5, "Web", "Scale Changed to " + GlobalVariables.Scale)
                AddWarning(0, "Scale Changed to " + GlobalVariables.Scale)

            Case "oDebugLevel"
                Dim DebugLevel As String = parts("DebugLevel")
                If IsNumeric(DebugLevel) Then
                    GlobalVariables.DebugLevel = CInt(DebugLevel)
                End If

                PostMessage("Dmessage", "Debug Level Now " + DebugLevel)
                Debug(5, "Web", "Changed DebugLevel to " + DebugLevel)
                AddWarning(0, "DebugLevel Changed to " + DebugLevel)

            Case "RefreshButtonStatus"
                'This button navigates to the sample status page.
                'Me.pageCommands.Add("newpage", plugin.Name & "Status")
                Debug(3, "Web", "Processed Status Refresh Button")
                Me.pageCommands.Add("Refresh", "true")
            Case "RefreshButtonConfig"
                'This button navigates to the sample status page.
                'Me.pageCommands.Add("newpage", plugin.Name & "Status")
                Debug(3, "Web", "Processed Config Refresh Button")
                Me.pageCommands.Add("Refresh", "true")
            Case "RefreshButtonDebug"
                'This button navigates to the sample status page.
                'Me.pageCommands.Add("newpage", plugin.Name & "Status")
                Debug(3, "Web", "Processed Debug Refresh Button")
                Me.divToUpdate.Add("debugq", DebugQTable())
                Me.divToUpdate.Add("statsdiv", Statistics)
            Case "ReloadButton"
                Debug(3, "Web", "Reload Initiated")
                'PostMessage("Cmessage", "Reload not yet implemented")
                AddWarning(0, "Reload Initiated")
                GlobalVariables.EnableQueue = False
                GlobalVariables.Reloading = True
                Threading.Thread.Sleep(5000)    '5 seconds
                plugin.ReadDevices()
                'GlobalVariables.EnableQueue = True
                Debug(3, "Web", "Reload Completed")
                AddWarning(0, "Reload Completed")
                PostMessage("Smessage", "Reload Completed")

            'Configs
            '
            Case "oLogFile"
                plugin.Settings.RawLogFile = parts("LogFile")
                Debug(4, "WEB", "Changed Logfile to [" + parts("LogFile") + "]")
            Case "SaveButton"
                Debug(3, "Web", "SAVE Initiated")
                GlobalVariables.EnableQueue = False
                Threading.Thread.Sleep(5000)    '5 seconds
                plugin.Settings.Save()
                Debug(3, "Web", "SAVE Completed")
                PostMessage("Cmessage", "Configuration Saved. Disable and Enable plug-in to take effect.")

            Case "timer"
                'This stops the timer and clears the message
                If TimerEnabled Then 'this handles the initial timer post that occurs immediately upon enabling the timer.
                    TimerEnabled = False
                Else
                    Me.pageCommands.Add("stoptimer", "")
                    Me.divToUpdate.Add("message", "&nbsp;")
                End If
        End Select

        Return MyBase.postBackProc(page, data, user, userRights)
    End Function

    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String
        Dim stb As New StringBuilder
        Dim instancetext As String = ""
        Try
            Me.reset()
            currentPage = Me

            ' handle any queries like mode=something
            Dim parts As Collections.Specialized.NameValueCollection = Nothing
            If queryString <> String.Empty Then parts = HttpUtility.ParseQueryString(queryString)
            If instance <> "" Then instancetext = " - " & instance

            'For some reason, I can't get the sample to add the title. So let's add it here.
            stb.Append("<title>" & plugin.Name & " " & pageName.Replace(plugin.Name, "") & "</title>")

            'Add menus and headers
            stb.Append(hs.GetPageHeader(pageName, plugin.Name & " " & pageName.Replace(plugin.Name, "") & instancetext, "", "", False, False))

            'Adds the div for the plugin page
            stb.Append(clsPageBuilder.DivStart("pluginpage", ""))

            ' a message area for error messages from jquery ajax postback (optional, only needed if using AJAX calls to get data)
            stb.Append(clsPageBuilder.DivStart("errormessage", "class='errormessage'"))
            stb.Append(clsPageBuilder.DivEnd)

            'Configures the timer that all pages apparently has
            Me.RefreshIntervalMilliSeconds = 3000
            stb.Append(Me.AddAjaxHandlerPost("id=timer", pageName)) 'This is so we can control it in postback

            ' specific page starts here
            stb.Append(BuildTabs())

            'Ends the div end tag for the plugin page
            stb.Append(clsPageBuilder.DivEnd)

            ' add the body html to the page
            Me.AddBody(stb.ToString)

            ' return the full page
            Return Me.BuildPage()
        Catch ex As Exception
            'WriteMon("Error", "Building page: " & ex.Message)
            Return "ERROR in GetPagePlugin: " & ex.Message
        End Try
    End Function

    Private Sub PostMessage(ByVal DIVName As String, ByVal message As String)
        'Updates the div
        Me.divToUpdate.Add(DIVName, message)
        'Starts the pages built-in timer
        Me.pageCommands.Add("starttimer", "")

        '... and updates the local variable so we can easily check if the timer is running
        TimerEnabled = True
    End Sub

    Public Function BuildTabs() As String
        Dim stb As New StringBuilder
        Dim tabs As clsJQuery.jqTabs = New clsJQuery.jqTabs("oTabs", Me.PageName)
        Dim tab As New clsJQuery.Tab

        tabs.postOnTabClick = True
        tab.tabTitle = "Status"
        tab.tabDIVID = "oTabStatus"
        tab.tabContent = "<div id='TabConfig_div'>" & BuildContent("Status") & "</div>"
        tabs.tabs.Add(tab)

        tab = New clsJQuery.Tab
        tab.tabTitle = "Config"
        tab.tabDIVID = "oTabConfig"
        tab.tabContent = "<div id='TabRelease_div'>" & BuildContent("Config") & "</div>"
        tabs.tabs.Add(tab)

        tab = New clsJQuery.Tab
        tab.tabTitle = "Release"
        tab.tabDIVID = "oTabRelease"
        tab.tabContent = "<div id='TabRelease_div'>" & BuildContent("Release") & "</div>"
        tabs.tabs.Add(tab)

        tab = New clsJQuery.Tab
        tab.tabTitle = "Debug"
        tab.tabDIVID = "oTabDebug"
        tab.tabContent = "<div id='TabDebug_div'>" & BuildContent("Debug") & "</div>"
        tabs.tabs.Add(tab)

        tab = New clsJQuery.Tab
        tab.tabTitle = "Help"
        tab.tabDIVID = "oTabHelp"
        tab.tabContent = "<div id='TabHelp_div'>" & BuildContent("Help") & "</div>"
        tabs.tabs.Add(tab)

        Return tabs.Build
    End Function
    Function BuildContent(PageName As String) As String
        Dim FG As String = ""
        Dim DType As String = ""
        Dim DState As String = ""
        Dim stb As New StringBuilder
        Dim Tablewidth As Int16 = 950
        Dim Levelname As String = ""

        Select Case PageName
            Case "Status"
                stb.Append("<h2>" + GlobalVariables.ModuleName + " Plug-in</h2>")
                stb.Append("<table cellpadding='0' cellspacing='0' width='" + Tablewidth.ToString + "'>")
                stb.Append("<tr class='tablerowodd'><td>Version</td><td>" + Version + " " + Release + "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Queue Processing</td><td>" + BoolToYesNo(GlobalVariables.EnableQueue) + "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Units / Devices</td><td>" + GlobalVariables.KMTUnits.Count.ToString + " / " + GlobalVariables.KMTDevices.Count.ToString + "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Logging to File</td><td>" + BoolToYesNo(GlobalVariables.LogFile) + "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Temperature Scale</td><td>" + GlobalVariables.Scale + "</td></tr>")
                stb.Append("<tr><td colspan='2' align='center' style='color:#FF0000; font-size:12pt;'><strong><div id='Smessage'>&nbsp;</div></strong></td></tr>")
                stb.Append("</table>")
                stb.Append("<h2>Units</h2>")
                stb.Append("<table style='border: 1px solid grey;' cellpadding='4' cellspacing='4' width='" + Tablewidth.ToString + "'>")
                stb.Append("<tr style='color:#FFFFFF; background-color:#2E64FE'><td><strong>IP</strong></td><td><strong>Map</strong></td><td><strong>Last Access</strong></td><td><strong>Active</strong></td></tr>")
                For Each TasDevice In GlobalVariables.KMTUnits
                    stb.Append("<tr><td>" + TasDevice.DeviceIP.ToString + "</td><td>" + TasDevice.Map + "</td><td>" + TasDevice.LastSeen.ToString + "</td><td>" + BoolToYesNo(TasDevice.Active.ToString) + "</td></tr>")
                Next
                stb.Append("</table>")
                stb.Append("<h2>Devices</h2>")
                stb.Append("<table style='border: 1px solid grey;' cellpadding='4' cellspacing='4' width='" + Tablewidth.ToString + "'>")
                stb.Append("<tr style='color:#FFFFFF; background-color:#2E64FE'><td><strong>ID</strong></td><td><strong>HS Device</strong></td><td><strong>Type</strong></td><td><strong>State</strong></td><td><strong>Device Name</strong></td><td><strong>Target</strong></td><td><Strong>Last Seen</strong></td><td><strong>Active</strong></td></tr>")
                For Each TasDevice In GlobalVariables.KMTDevices
                    If TasDevice.Active Then FG = "#000000" Else FG = "#6E6E6E"
                    Select Case TasDevice.Type
                        Case 1 : DType = "Output"
                            If hs.IsON(TasDevice.DeviceID) = True Then DState = "ON" Else DState = "OFF"
                        Case 2 : DType = "Sensor"
                            DState = hs.DeviceValueEx(TasDevice.DeviceID).ToString
                        Case Else
                            DType = "Other"
                    End Select
                    stb.Append("<tr style='color:" + FG + "'><td><a href='/deviceutility?ref=" + TasDevice.DeviceID.ToString + "&edit=1' target='_self'>" + TasDevice.DeviceID.ToString + "</a></td><td>" + TasDevice.HSDeviceName + "</td><td>" + DType + "</td><td>" + DState + "</td><td>" + TasDevice.DeviceName + "</td><td>" + TasDevice.DeviceTarget + "</td><td>" + TasDevice.LastSeen.ToString + "</td><td>" + BoolToYesNo(TasDevice.Active.ToString) + "</td></tr>")
                Next
                stb.Append("</table>")
                If GlobalVariables.Warnings.Count > 0 Then
                    stb.Append("<h2>Messages</h2>")
                    stb.Append("<table style='border: 1px solid grey;' cellpadding='4' cellspacing='4' width='" + Tablewidth.ToString + "'>")
                    stb.Append("<tr style='color:#FFFFFF; background-color:#2E64FE'><td><strong>Time</strong></td><td>Age (mins)</td><td><strong>Level</strong></td><td><strong>Warning</strong></td></tr>")
                    For Each W In GlobalVariables.Warnings '0 = Info, 1 = Notice, 2 = Warning, 3 = Error, 4 = Critical
                        Select Case W.Level
                            Case 0 : Levelname = "Info"
                            Case 1 : Levelname = "Notice"
                            Case 2 : Levelname = "Warning"
                            Case 3 : Levelname = "Error"
                            Case 4 : Levelname = "Critical"
                            Case Else
                                Levelname = "Unknown"
                        End Select
                        stb.Append("<tr style='color:#000000'><td>" + W.TimeStamp.ToString + "</td><td>" + DateDiff(DateInterval.Minute, W.TimeStamp, DateTime.Now).ToString + "</td><td>" + Levelname + "</td><td>" + W.Warning.ToString + "</td></tr>")
                    Next
                    stb.Append("</table>")
                Else
                    stb.Append("<p>No Messages Outstanding</p>")
                End If
                stb.Append("<p>" + BuildButton("RefreshButtonStatus", False) + "</p>")
                stb.Append("<p>" + BuildButton("ReloadButton", False) + "</p>")
            Case "Config"
                stb.Append("<h2>" + GlobalVariables.ModuleName + " Plug-in</h2>")
                stb.Append("<table cellpadding='0' cellspacing='0' width='" + Tablewidth.ToString + "'>")
                stb.Append("<tr class='tablerowodd'><td>Version</td><td>" + Version + " " + Release + "</td></tr>")

                stb.Append("<tr class='tablerowodd'><td>Queue Processing</td><td>" + BoolToYesNo(GlobalVariables.EnableQueue) + "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Devices</td><td>" + GlobalVariables.KMTDevices.Count.ToString + "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Units</td><td>" + GlobalVariables.KMTUnits.Count.ToString + "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Logging to File</td><td>" + BoolToYesNo(GlobalVariables.LogFile) + "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Scale</td><td>" + BuildDropList("Scale") + "</td></tr>")
                stb.Append("<tr><td colspan='2' align='center' style='color:#FF0000; font-size:12pt;'><strong><div id='Cmessage'>&nbsp;</div></strong></td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Log File</td><td>" + BuildTextBox("LogFile", False, plugin.Settings.RawLogFile) + "</td></tr>")
                stb.Append("<tr><td colspan='3'>&nbsp;</td></tr>")
                stb.Append("<tr><td colspan='2'>" + BuildButton("SaveButton", False) + "</td></tr>")
                stb.Append("<tr><td colspan='2'>" + BuildButton("RefreshButtonConfig", False) + "</td></tr>")
                stb.Append("</table>")

            Case "Release"
                stb.Append("<table cellpadding='0' cellspacing='0' width='" + Tablewidth.ToString + "'>")
                stb.Append("<tr class='tablerowodd'><td><h2>Latest Version & Release Information</td></tr>")
                stb.Append("<tr><td>Latest Version <b>" + GlobalVariables.LatestVersion + "</b> Currently Running Version <b>" + Version + "</b></td></tr>")
                stb.Append("<tr><td><h2>Release Notes</h2>&nbsp;" + GlobalVariables.LatestVersion + GlobalVariables.ReleaseNotes + "</td></tr>")
                stb.Append("</table>")
            Case "Debug"
                stb.Append("<table border='0' cellpadding='0' cellspacing='0' width='" + Tablewidth.ToString + "'>")
                stb.Append("<tr><td colspan='3' align='center' style='color:#FF0000; font-size:12pt;'><strong><div id='Dmessage'>&nbsp;</div></strong></td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Debug Level</td><td>" + GlobalVariables.DebugLevel.ToString + "</td><td rowspan='6' style='vertical-align: top;'><div id='statsdiv'>" + Statistics + "</div></td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Log File</td><td>" + GlobalVariables.RawLogFileName.ToString + " (" + GlobalVariables.LogFile.ToString + ")</td</tr>")
                stb.Append("<tr class='tablerowodd'><td>Queue</td><td>" + GlobalVariables.eQUEUE.Count().ToString + "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Buffer</td><td>" + DebugQ.Count().ToString + "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Debug Logging</td><td>" & BuildDropList("DebugLevel") & "</td></tr>")
                stb.Append("<tr class='tablerowodd'><td>Commands</td><td>" & BuildButton("RefreshButtonDebug", False) & "</td></tr>")
                stb.Append("</table>")
                stb.Append(DebugQTable())
            Case "Help"
                stb.Append("<table cellpadding='0' cellspacing='0' width='" + Tablewidth.ToString + "'>")
                stb.Append("<tr class='tablerowodd'><td><h2>KMTronic Ltd</h2>")
                stb.Append("<p>KMTronic is a European Designer and Manufacturer or Electronic devices based at Dobri Chintulov 28A, 5100 Gorna Oriachovitca, VT, Bulgaria</p>")
                stb.Append("<p>They have a range of binary output devices over USB, RS232, RS485, LPT, UART, Modbus, RF and LAN.</p>")
                stb.Append("</td></tr><tr><td><h2>LAN Controllers</h2>")
                stb.Append("<p>This plug-in for Homeseer HS3 works exclusively with the 8 channel WEB LAN Product (Product Code W8CR) which is also available in a DIN mounting case. We selected this unit because its fairly easy to setup, and as commands are over TCP/IP on port 80 there's a high degree of reliability in communicaitons.</p>")
                stb.Append("</td></tr><tr><td><h2>LAN Controllers</h2>")
                stb.Append("<p>This plug-in for Homeseer HS3 works exclusively with the 4 channel LAN DS18B20 WEB Temperature Monitor (Product Code SS_LAN_DS18B20_MONITOR). This LAN based sensor supports up to 4 DS18B20 Sensors. The High accuracy and easy availability of these sensors makes this an idea device for reliable temperature monitoring on the network</p>")
                stb.Append("</td></tr><tr><td><h2>Links</h2>")
                stb.Append("<ul><li>General Information is available <a href='https://www.gen.net.uk/kmtronic-homeseer' target='_blank'>Here</a></li>")
                stb.Append("<li>The BugTracker for Reporting Issues is available <a href='https://mantis.gen.net.uk/project_page.php?project_id=25' target='_blank'>Here</a></li></ul>")
                stb.Append("<tr class='tablerowodd'><td>When reporting bugs and issues please include a full description of the issue, a logfile and any screenshots etc that will help us to track down the issue and resolve it.</td></tr>")
                stb.Append("<tr class='tablerowodd'><td></td></tr>")
                stb.Append("</table>")
            Case Else
                stb.Append("<p>There is no content for this tab at the moment</p>")
        End Select
        stb.Append("<p>&copy; 2018 - Developed by Richard Taylor (109). For Support use the BugTracker in the Help Tab.</p>")

        Return stb.ToString
    End Function

    Public Function DebugQTable()
        Dim stb As New StringBuilder
        Dim DebugArray As Array = DebugQ.ToArray
        stb.Append("<div id='debugq'><table cellstacing='0' cellpadding='0' style='table-layout:fixed; padding: 0px; border-spacing: 0px; border-collapse: separate; border: 0px; width:950px'>")
        For i As Integer = UBound(DebugArray) To 0 Step -1
            stb.Append(DebugArray(i))
        Next
        stb.Append("</table></div>")
        Return stb.ToString
    End Function
    Public Function BuildTextBox(ByVal Name As String, Optional ByVal Rebuilding As Boolean = False, Optional ByVal Text As String = "") As String
        Dim textBox As New clsJQuery.jqTextBox(Name, "", Text, Me.PageName, 20, False)
        Dim ret As String = ""
        textBox.id = "o" & Name


        If Rebuilding Then
            ret = textBox.Build
            Me.divToUpdate.Add(Name & "_div", ret)
        Else
            ret = "<div style='float: left;'  id='" & Name & "_div'>" & textBox.Build & "</div>"
        End If

        Return ret
    End Function

    Function BuildButton(ByVal Name As String, Optional ByVal Rebuilding As Boolean = False) As String
        Dim button As New clsJQuery.jqButton(Name, "", Me.PageName, False)
        Dim buttonText As String = "Submit"
        Dim ret As String = String.Empty

        'Handles the text for different buttons, based on the button name
        Select Case Name
            Case "RefreshButtonStatus"
                buttonText = "Refresh"
                button.submitForm = False
            Case "RefreshButtonConfig"
                buttonText = "Refresh"
                button.submitForm = False
            Case "RefreshButtonDebug"
                buttonText = "Refresh"
                button.submitForm = False
            Case "ReloadButton"
                buttonText = "Reload"
                button.submitForm = False
            Case "SaveButton"
                buttonText = "Save"
                button.submitForm = False
        End Select

        'button.id = "o" & Name
        button.id = Name
        button.label = buttonText

        ret = button.Build

        If Rebuilding Then
            Me.divToUpdate.Add(Name & "_div", ret)
        Else
            ret = "<div style='float: left;' id='" & Name & "_div'>" & ret & "</div>"
        End If
        Return ret
    End Function

    Public Function BuildDropList(ByVal Name As String, Optional ByVal Rebuilding As Boolean = False) As String
        Dim ddl As New clsJQuery.jqDropList(Name, Me.PageName, False)
        ddl.id = "o" & Name
        ddl.autoPostBack = True

        Select Case Name
            Case "Scale"
                Dim eCelsius = False
                Dim eFahrenheit = False
                Dim eRankine = False
                Dim eDelisle = False
                Dim eNewton = False
                Dim eRéaumur = False
                Dim eRømer = False
                Dim ePlanck = False
                Select Case GlobalVariables.Scale
                    Case "Celsius" : eCelsius = True
                    Case "Fahrenheit" : eFahrenheit = True
                    Case "Rankine" : eRankine = True
                    Case "Delisle" : eDelisle = True
                    Case "Newton" : eNewton = True
                    Case "Réaumur" : eRéaumur = True
                    Case "Rømer" : eRømer = True
                    Case "Planck" : ePlanck = True
                End Select
                ddl.AddItem("Celsius", "Celsius", eCelsius)
                ddl.AddItem("Fahrenheit", "Fahrenheit", eFahrenheit)
                ddl.AddItem("Rankine", "Rankine", eRankine)
                ddl.AddItem("Delisle", "Delisle", eDelisle)
                ddl.AddItem("Newton", "Newton", eNewton)
                ddl.AddItem("Réaumur", "Réaumur", eRéaumur)
                ddl.AddItem("Rømer", "Rømer", eRømer)
                ddl.AddItem("Planck", "Planck", ePlanck)
            Case "DebugLevel"
                ddl.AddItem("Normal", "5", True)
                ddl.AddItem("Increased", "4", False)
                ddl.AddItem("Diagnostic 1", "3", False)
                ddl.AddItem("Diagnostic 2", "2", False)
                ddl.AddItem("Low Level", "1", False)
                ddl.AddItem("Everything", "0", False)
        End Select

        Dim ret As String
        If Rebuilding Then
            ret = ddl.Build
            Me.divToUpdate.Add(Name & "_div", ret)
        Else
            ret = "<div style='float: left;'  id='" & Name & "_div'>" & ddl.Build & "</div>"
        End If
        Return ret
    End Function

    Public Function BuildCheckbox(ByVal Name As String, Optional ByVal Rebuilding As Boolean = False, Optional ByVal label As String = "") As String
        Dim checkbox As New clsJQuery.jqCheckBox(Name, label, Me.PageName, True, False)
        'checkbox.id = "o" & Name
        checkbox.id = Name

        Select Case Name

            Case "CheckboxDebugLogging"
                If GlobalVariables.DebugLevel = 0 Then
                    checkbox.checked = True
                Else
                    checkbox.checked = False
                End If
        End Select

        Dim ret As String = String.Empty
        If Rebuilding Then
            ret = checkbox.Build
            Me.divToUpdate.Add(Name & "_div", ret)
        Else
            ret = "<div style='float: left;'  id='" & Name & "_div'>" & checkbox.Build & "</div>"
        End If
        Return ret
    End Function


End Class

