Imports System.Text
Imports System.IO
Imports System.Threading
Imports System.Web
Imports System
Imports HomeSeerAPI
Imports Scheduler
Imports System.Collections.Specialized

Public Class CurrentPlugin
    Implements IPlugInAPI
    Public Settings As Settings 'My simple class for reading and writing settings
    Dim WithEvents updateTimer As Threading.Timer
    Private lastRandomNumber As Integer 'The last random value

    Dim configPageName As String = Me.Name & "Config"
    Dim supportPageName As String = Me.Name & "Status"
    Dim configPage As New web_config(configPageName)
    Dim SupportPage As New web_support(supportPageName)
    Dim webPage As Object

    Dim actions As New hsCollection
    Dim triggers As New hsCollection
    Dim trigger As New Trigger
    Dim action As New Action

    Private Shared myTimer1 As System.Threading.Timer
    Private Shared MyTimer2 As System.Threading.Timer

    Const Pagename = "Events" 'The controls we build here (in Plugin.vb, typically for controlling actions, conditions and triggers) are all located on the Events page


#Region "Init"

    Public Function InitIO(ByVal port As String) As String
        Debug(5, "Plugin", "Starting InitIO initializiation.")

        'Loading settings before we do anything else
        Me.Settings = New Settings

        'Registering two pages
        RegisterWebPage(link:=configPageName, linktext:="Config", page_title:="Configuration")

        Debug(5, "Plugin", "Starting Version " + Version + " " + Release)
        Debug(5, "Plugin", "Developed by Richard Taylor (109)")
        Debug(5, "Plugin", "Hardware by KMTronic")
        Debug(5, "Plugin", "-----------------------------------------------------------------")
        GlobalVariables.DebugLevel = Settings.IniDebugLevel
        If Settings.RawLogFile <> "" Then
            Try
                GlobalVariables.RawLog = New System.IO.StreamWriter(Settings.RawLogFile, False)
                GlobalVariables.RawLogFileName = Settings.RawLogFile
                Debug(5, "Plugin", "Opened Log " + Settings.RawLogFile.ToString)
                GlobalVariables.LogFile = True
            Catch ex As Exception
                Debug(5, "Plugin", "Failed to Open Logfile " + Settings.RawLogFile.ToString + " Error " + ex.Message.ToString)
            End Try

        End If

        ReadDevices()

        callback.RegisterEventCB(Enums.HSEvent.VALUE_CHANGE, IFACE_NAME, "")
        Debug(5, "Plugin", "Debug Level " + GlobalVariables.DebugLevel.ToString)
        Debug(5, "Plugin", "Launching Threads")

        Dim myCallBack As New System.Threading.TimerCallback(AddressOf QueueProcessing.ProcessQueue)
        myTimer1 = New System.Threading.Timer(myCallBack, Nothing, 100, 2000)
        AddWarning(0, "Enabled Event Queue Processing Threads")

        Dim MyHouse As New System.Threading.TimerCallback(AddressOf Housekeeping.Maintenance)
        MyTimer2 = New System.Threading.Timer(MyHouse, Nothing, 100, 30000)
        AddWarning(0, "Enabled Maintenance Threads")

        Debug(5, "Plugin", "Ready")

        Return ""
    End Function

    Sub ReadDevices()
        Dim DType As UInt16 = 0
        Dim IPByte As System.Net.IPAddress = System.Net.IPAddress.Parse("0.0.0.0")
        Dim Map As String = "00000000"
        GlobalVariables.KMTDevices.Clear()
        GlobalVariables.KMTUnits.Clear()

        Dim devs = (From d In Devices()).ToList
        Dim DeviceType As String = ""

        For Each dev In (From d In devs Where d.Device_Type_String(hs).Contains("KMT:"))
            DeviceType = Trim(dev.Device_Type_String(hs).ToString)
            Debug(1, "Plugin", "Device Found " + dev.Ref(hs).ToString + " = " + DeviceType)
            DeviceType = Replace(DeviceType, "KMT:", "")    'KMT:192.168.16.10/output1 or KMT:192.168.16.10/sensor1
            Dim DeviceItems = Split(DeviceType, "/")
            If UBound(DeviceItems) = 1 Then
                If System.Net.IPAddress.TryParse(DeviceItems(0), IPByte) Then
                    If InStr(DeviceItems(1).ToUpper, "OUTPUT") > 0 Then
                        DeviceItems(1) = Replace(DeviceItems(1).ToUpper, "OUTPUT", "")
                        If IsNumeric(DeviceItems(1)) Then
                            If CInt(DeviceItems(1)) > 0 And CInt(DeviceItems(1)) < 9 Then
                                GlobalVariables.KMTDevices.Add(New DeviceInfo() With {.DeviceID = dev.Ref(hs), .DeviceString = DeviceType, .DeviceName = DeviceItems(0), .HSDeviceName = dev.Name(hs).ToString, .DeviceTarget = DeviceItems(1), .Active = False, .Type = 1})
                                Debug(5, "Plugin", "Output Device Added " + dev.Ref(hs).ToString + " : " + DeviceType + " [" + DeviceItems(0).ToString + "][" + DeviceItems(1).ToString + "]")
                            Else
                                Debug(5, "Plugin", "Device Output Number Out-of-range. Should be 1 to 8.")
                            End If
                        Else
                            Debug(5, "Plugin", "Output Device Output Number Malformed. Should be 1 to 8.")
                        End If
                    ElseIf InStr(DeviceItems(1).ToUpper, "SENSOR") > 0 Then
                        DeviceItems(1) = Replace(DeviceItems(1).ToUpper, "SENSOR", "")
                        If IsNumeric(DeviceItems(1)) Then
                            If CInt(DeviceItems(1)) > 0 And CInt(DeviceItems(1)) < 5 Then
                                GlobalVariables.KMTDevices.Add(New DeviceInfo() With {.DeviceID = dev.Ref(hs), .DeviceString = DeviceType, .DeviceName = DeviceItems(0), .HSDeviceName = dev.Name(hs).ToString, .DeviceTarget = DeviceItems(1), .Active = False, .Type = 2})
                                Debug(5, "Plugin", "Sensor Device Added " + dev.Ref(hs).ToString + " : " + DeviceType + " [" + DeviceItems(0).ToString + "][" + DeviceItems(1).ToString + "]")
                            Else
                                Debug(5, "Plugin", "Sensor Output Number Out-of-range. Should be 1 to 4.")
                            End If
                        Else
                            Debug(5, "Plugin", "Sensor Output Number Malformed. Should be numeric")
                        End If
                    Else
                        Debug(5, "Plugin", "Device Malformed. Should be output1 to 8 or sensor1 to 4.")
                    End If
                Else
                    Debug(4, "Plugin", "Device IP Address is malformed. Must be IPv4 address.")
                End If
            Else
                Debug(4, "Plugin", "Malformed Device Definition [" + DeviceType + "]")
            End If
        Next
        ' Now extrapolate units from devices
        Dim TheDeviceIP As String
        Dim TheDeviceTarget As Int16
        For Each Devices As DeviceInfo In GlobalVariables.KMTDevices

            If GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = Devices.DeviceName) IsNot Nothing Then
                TheDeviceIP = GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = Devices.DeviceName).DeviceIP
                Debug(1, "Plugin", "Processing Device " + Devices.DeviceName.ToString + " DeviceIP Find =[" + TheDeviceIP.ToString + "]")
            Else
                TheDeviceIP = Nothing
                Debug(1, "Plugin", "Processing Device " + Devices.DeviceName.ToString + " Not Found in KMTUnits")
                Map = "00000000"
            End If

            TheDeviceTarget = CInt(Devices.DeviceTarget)
            Select Case Devices.Type
                Case 1
                    If TheDeviceIP IsNot Nothing Then
                        Map = GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = Devices.DeviceName).Map
                        Mid(Map, CInt(TheDeviceTarget), 1) = "1"
                        GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = Devices.DeviceName).Map = Map
                        Debug(1, "Plugin", "Editing Unit " + Devices.DeviceName.ToString + " " + TheDeviceIP.ToString + " Map " + Map.ToString + " bit " + TheDeviceTarget.ToString)
                    Else
                        Debug(1, "Plugin", "Adding New Unit " + Devices.DeviceName.ToString + " Map " + Map.ToString)
                        Mid(Map, CInt(TheDeviceTarget), 1) = "1"
                        GlobalVariables.KMTUnits.Add(New UnitInfo() With {.DeviceIP = Devices.DeviceName, .Map = Map, .Bits = "00000000", .Active = 0, .LastSeen = DateAdd(DateInterval.Minute, -5, Now), .Type = 1})
                        AddWarning(0, "Discovered Unit at " + Devices.DeviceName)
                    End If
                Case 2
                    If TheDeviceIP IsNot Nothing Then
                        'GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = Devices.DeviceName).Map = Map
                        Debug(1, "Plugin", "Editing Unit " + Devices.DeviceName.ToString + " " + TheDeviceIP.ToString + " Map " + Map.ToString + " bit " + TheDeviceTarget.ToString)
                    Else
                        Debug(1, "Plugin", "Adding New Unit " + Devices.DeviceName.ToString)
                        GlobalVariables.KMTUnits.Add(New UnitInfo() With {.DeviceIP = Devices.DeviceName, .Map = "Sensor", .Bits = "00000000", .Active = 0, .LastSeen = DateAdd(DateInterval.Minute, -5, Now), .Type = 2})
                        AddWarning(0, "Discovered Unit at " + Devices.DeviceName)
                    End If
            End Select
        Next
        GlobalVariables.Reloading = False

    End Sub


    Sub HSEvent(ByVal EventType As Enums.HSEvent, ByVal parms() As Object) Implements HomeSeerAPI.IPlugInAPI.HSEvent

    End Sub

    Public Sub ShutdownIO()
        Try
            ''**********************
            ''For debugging only, this will delete all devices accociated by the plugin at shutdown, so new devices will be created on startup:
            ' DeleteDevices()
            ''**********************

            'Setting a flag that states that we are shutting down, this can be used to abort ongoing commands
            GlobalVariables.IsShuttingDown = True
            GlobalVariables.EnableQueue = False

            'Write any changes in the settings clas to the ini file
            'Me.Settings.Save()

            'Stopping the timer if it exists and runs
            If updateTimer IsNot Nothing Then
                updateTimer.Change(Timeout.Infinite, Timeout.Infinite)
                updateTimer.Dispose()
            End If

            callback.UnRegisterGenericEventCB(Enums.HSEvent.VALUE_CHANGE, IFACE_NAME, "")
            myTimer1.Change(Threading.Timeout.Infinite, Threading.Timeout.Infinite)
            MyTimer2.Change(Threading.Timeout.Infinite, Threading.Timeout.Infinite)
            'MQTT.CloseMQTT()


        Catch ex As Exception
            Debug(9, System.Reflection.MethodBase.GetCurrentMethod().Name, "Error ending " & Me.Name & " Plug-In")
        End Try
        Debug(5, "Plugin", "ShutdownIO complete.")
        If GlobalVariables.LogFile Then
            GlobalVariables.RawLog.Close()
        End If
    End Sub


#End Region

#Region "Action/Trigger/DeviceConfig Processes"

#Region "Device Config Interface"

    Public Function ConfigDevice(ref As Integer, user As String, userRights As Integer, newDevice As Boolean) As String
        Dim device As Scheduler.Classes.DeviceClass = Nothing
        Dim stb As New StringBuilder

        device = hs.GetDeviceByRef(ref)

        Dim PED As clsPlugExtraData = device.PlugExtraData_Get(hs)
        Dim PEDname As String = Me.Name

        'We'll use the device type string to determine how we should handle the device in the plugin
        Select Case device.Device_Type_String(hs).Replace(Me.Name, "").Trim
            Case ""
                Dim sample As SampleClass = PEDGet(PED, PEDname)
                Dim houseCodeDropDownList As New clsJQuery.jqDropList("HouseCode", "", False)
                Dim unitCodeDropDownList As New clsJQuery.jqDropList("DeviceCode", "", False)
                Dim saveButton As New clsJQuery.jqButton("Save", "Done", "DeviceUtility", True)
                Dim houseCode As String = ""
                Dim deviceCode As String = ""


                If sample Is Nothing Then
                    Console.WriteLine("ConfigDevice, sample is nothing")
                    ' Set the defaults
                    sample = New SampleClass
                    InitHSDevice(device, device.Name(hs))
                    sample.houseCode = "A"
                    sample.deviceCode = "1"
                    PEDAdd(PED, PEDname, sample)
                    device.PlugExtraData_Set(hs) = PED
                End If

                houseCode = sample.houseCode
                deviceCode = sample.deviceCode

                For Each l In "ABCDEFGHIJKLMNOP"
                    houseCodeDropDownList.AddItem(l, l, l = houseCode)
                Next
                For i = 1 To 16
                    unitCodeDropDownList.AddItem(i.ToString, i.ToString, i.ToString = deviceCode)
                Next

                Try
                    stb.Append("<form id='frmSample' name='SampleTab' method='Post'>")
                    stb.Append(" <table border='0' cellpadding='0' cellspacing='0' width='610'>")
                    stb.Append("  <tr><td colspan='4' align='Center' style='font-size:10pt; height:30px;' nowrap>Select a houseCode and Unitcode that matches one of the devices HomeSeer will be communicating with.</td></tr>")
                    stb.Append("  <tr>")
                    stb.Append("   <td nowrap class='tablecolumn' align='center' width='70'>House<br>Code</td>")
                    stb.Append("   <td nowrap class='tablecolumn' align='center' width='70'>Unit<br>Code</td>")
                    stb.Append("   <td nowrap class='tablecolumn' align='center' width='200'>&nbsp;</td>")
                    stb.Append("  </tr>")
                    stb.Append("  <tr>")
                    stb.Append("   <td class='tablerowodd' align='center'>" & houseCodeDropDownList.Build & "</td>")
                    stb.Append("   <td class='tablerowodd' align='center'>" & unitCodeDropDownList.Build & "</td>")
                    stb.Append("   <td class='tablerowodd' align='left'>" & saveButton.Build & "</td>")
                    stb.Append("  </tr>")
                    stb.Append(" </table>")
                    stb.Append("</form>")
                    Return stb.ToString
                Catch ex As Exception
                    Return "ConfigDevice ERROR: " & ex.Message 'Original is too old school: "Return Err.Description"
                End Try


            Case "Basic"
                stb.Append("<form id='frmSample' name='SampleTab' method='Post'>")
                stb.Append("Nothing special to configure for the basic device. :-)")
                stb.Append("</form>")
                Return stb.ToString

            Case "Advanced"
                Dim savedString As String = PEDGet(PED, PEDname)
                If savedString = String.Empty Then 'The pluginextradata is not configured for this device
                    savedString = "The text in this textbox is saved with the actual device"
                End If

                Dim savedTextbox As New clsJQuery.jqTextBox("savedTextbox", "", savedString, "", 100, False)
                Dim saveButton As New clsJQuery.jqButton("Save", "Done", "DeviceUtility", True)

                stb.Append("<form id='frmSample' name='SampleTab' method='Post'>")
                stb.Append(" <table border='0' cellpadding='0' cellspacing='0' width='610'>")
                stb.Append("  <tr><td colspan='4' align='Center' style='font-size:10pt; height:30px;' nowrap>Text to be saved with the device.</td></tr>")
                stb.Append("  <tr>")
                stb.Append("   <td nowrap class='tablecolumn' align='center' width='70'>Text:</td>")
                stb.Append("   <td nowrap class='tablecolumn' align='center' width='200'>&nbsp;</td>")
                stb.Append("  </tr>")
                stb.Append("  <tr>")
                stb.Append("   <td class='tablerowodd' align='center'>" & savedTextbox.Build & "</td>")
                stb.Append("   <td class='tablerowodd' align='left'>" & saveButton.Build & "</td>")
                stb.Append("  </tr>")
                stb.Append(" </table>")
                stb.Append("</form>")

                Return stb.ToString
        End Select

        Return String.Empty
    End Function

    Public Function ConfigDevicePost(ref As Integer, data As String, user As String, userRights As Integer) As Enums.ConfigDevicePostReturn
        Dim ReturnValue As Integer = Enums.ConfigDevicePostReturn.DoneAndCancel

        Try
            Dim device As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(ref)
            Dim PED As clsPlugExtraData = device.PlugExtraData_Get(hs)
            Dim PEDname As String = Me.Name
            Dim parts As Collections.Specialized.NameValueCollection = HttpUtility.ParseQueryString(data)

            'We'll use the device type string to determine how we should handle the device in the plugin
            Select Case device.Device_Type_String(hs).Replace(Me.Name, "").Trim
                Case ""

                    Dim sample As SampleClass = PEDGet(PED, PEDname)
                    If sample Is Nothing Then
                        InitHSDevice(device)
                    End If

                    sample.houseCode = parts("HouseCode")
                    sample.deviceCode = parts("DeviceCode")

                    PED = device.PlugExtraData_Get(hs)
                    PEDAdd(PED, PEDname, sample)
                    device.PlugExtraData_Set(hs) = PED
                    hs.SaveEventsDevices()

                Case "Basic"
                    'Nothing to store as this device doesn't have any extra data to save

                Case "Advanced"
                    'We'll get the string to save from the postback values
                    Dim savedString As String = parts("savedTextbox")

                    'We'll save this to the pluginextradata storage
                    PED = device.PlugExtraData_Get(hs)
                    PEDAdd(PED, PEDname, savedString) 'Adds the saveString to the plugin if it doesn't exist, and removes and adds it if it does.
                    device.PlugExtraData_Set(hs) = PED

                    'And then finally save the device
                    hs.SaveEventsDevices()

            End Select

            Return ReturnValue
        Catch ex As Exception
            Debug(9, System.Reflection.MethodBase.GetCurrentMethod().Name, "ConfigDevicePost: " & ex.Message)
        End Try
        Return ReturnValue
    End Function

    Public Sub SetIOMulti(ByVal colSend As List(Of HomeSeerAPI.CAPI.CAPIControl))
        'Multiple CAPIcontrols might be sent at the same time, so we need to check each one
        For Each CC In colSend
            Console.WriteLine("SetIOMulti triggered, checking CAPI '" & CC.Label & "' on device " & CC.Ref)

            hs.SetDeviceValueByRef(CC.Ref, CC.ControlValue, False)

            'Get the device sending the CAPIcontrol
            Dim device As Scheduler.Classes.DeviceClass = hs.GetDeviceByRef(CC.Ref)

            'We can get the PlugExtraData, if anything is stored in the device itself. What is stored is based on the device type.
            Select Case device.Device_Type_String(hs).Replace(Me.Name, "").Trim
                Case ""

                    Dim PED As clsPlugExtraData = device.PlugExtraData_Get(hs)
                    Dim sample As SampleClass = PEDGet(PED, "Sample")
                    If sample IsNot Nothing Then
                        Dim houseCode As String = sample.houseCode
                        Dim Devicecode As String = sample.deviceCode
                        SendCommand(houseCode, Devicecode) 'The HSPI_SAMPE control, in utils.vb as an example (but it doesn't do anything)
                    End If


                Case "Basic"
                    'There's nothing stored in the basic device

                Case "Advanced"
                    'Here we could choose to do something with the text string stored in the device

                Case Else
                    'Nothing to do at the moment
            End Select

        Next
    End Sub

#End Region

#Region "Trigger Properties"

    Public ReadOnly Property HasTriggers() As Boolean
        Get
            Return (TriggerCount() > 0)
        End Get
    End Property

    Public ReadOnly Property HasConditions(TriggerNumber As Integer) As Boolean
        Get
            Return True
        End Get
    End Property

    Public Function TriggerCount() As Integer
        Return triggers.Count
    End Function

    Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer
        Get
            Dim trigger As Trigger
            If IsValidTrigger(TriggerNumber) Then
                trigger = triggers(TriggerNumber)
                If Not (trigger Is Nothing) Then
                    Return trigger.Count
                Else
                    Return 0
                End If
            Else
                Return 0
            End If
        End Get
    End Property

    Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String
        Get
            If Not IsValidTrigger(TriggerNumber) Then
                Return ""
            Else
                Return Me.Name & ": " & triggers.Keys(TriggerNumber - 1)
            End If
        End Get
    End Property

    Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String
        Get
            Dim trigger As Trigger
            If IsValidSubTrigger(TriggerNumber, SubTriggerNumber) Then
                trigger = triggers(TriggerNumber)
                Return Me.Name & ": " & trigger.Keys(SubTriggerNumber - 1)
            Else
                Return ""
            End If
        End Get
    End Property

    Friend Function IsValidTrigger(ByVal TrigIn As Integer) As Boolean
        If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
            Return True
        End If
        Return False
    End Function

    Public Function IsValidSubTrigger(ByVal TrigIn As Integer, ByVal SubTrigIn As Integer) As Boolean
        Dim trigger As Trigger = Nothing
        If TrigIn > 0 AndAlso TrigIn <= triggers.Count Then
            trigger = triggers(TrigIn)
            If Not (trigger Is Nothing) Then
                If SubTrigIn > 0 AndAlso SubTrigIn <= trigger.Count Then Return True
            End If
        End If
        Return False
    End Function

#End Region

#Region "Trigger Interface"

    Private Enum TriggerTypes
        WithoutSubtriggers = 1
        WithSubtriggers = 2
    End Enum

    Private Enum SubTriggerTypes
        LowerThan = 1
        EqualTo = 2
        HigherThan = 3
    End Enum


    Function TriggerTrue(TrigInfo As IPlugInAPI.strTrigActInfo) As Boolean
        'Let's specify the key name of the value we are looking for
        Dim key As String = "SomeValue"

        'Get the value from the trigger
        Dim triggervalue As Integer = GetTriggerValue(key, TrigInfo)

        Console.WriteLine("Conditional value found for " & key & ": " & triggervalue & vbTab & "Last random: " & lastRandomNumber)

        'Let's return if this condition is True or False
        Return (triggervalue >= lastRandomNumber)
    End Function

    Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean
        Get
            Dim itemsConfigured As Integer = 0
            Dim itemsToConfigure As Integer = 1
            Dim UID As String = TrigInfo.UID.ToString

            If Not (TrigInfo.DataIn Is Nothing) Then
                DeSerializeObject(TrigInfo.DataIn, trigger)
                For Each key As String In trigger.Keys
                    Select Case True
                        Case key.Contains("SomeValue_" & UID) AndAlso trigger(key) <> ""
                            itemsConfigured += 1
                    End Select
                Next
                If itemsConfigured = itemsToConfigure Then Return True
            End If

            Return False
        End Get
    End Property

    Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String
        Dim UID As String = TrigInfo.UID.ToString
        Dim stb As New StringBuilder
        Dim someValue As Integer = -1 'We'll set the default value. This value will indicate that this trigger isn't properly configured
        Dim dd As New clsJQuery.jqDropList("SomeValue_" & UID & sUnique, Pagename, True)

        dd.autoPostBack = True
        dd.AddItem("--Please Select--", -1, False) 'A selected option with the default value (-1) means that the trigger isn't configured

        If Not (TrigInfo.DataIn Is Nothing) Then
            DeSerializeObject(TrigInfo.DataIn, trigger)
        Else 'new event, so clean out the trigger object
            trigger = New Trigger
        End If

        For Each key As String In trigger.Keys
            Select Case True

                'We'll fetch the selected value if this trigger has been configured before.
                Case key.Contains("SomeValue_" & UID)
                    someValue = trigger(key)
            End Select
        Next

        'We'll add all the different selectable values (numbers from 0 to 100 with 10 in increments)
        'and we'll select the option that was selected before if it's an old value (see ("i = someValue") which will be true or false)
        For i As Integer = 0 To 100 Step 10
            dd.AddItem(p_name:=i, p_value:=i, p_selected:=(i = someValue))
        Next

        'Finally we'll add this to the stringbuilder, and return the value
        stb.Append("Select value:")
        stb.Append(dd.Build)

        Return stb.ToString
    End Function

    Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn
        Dim Ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn
        Dim UID As String = TrigInfo.UID.ToString

        Ret.sResult = ""
        ' HST: We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
        '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
        '   we can still do that.
        Ret.DataOut = TrigInfo.DataIn
        Ret.TrigActInfo = TrigInfo

        If PostData Is Nothing Then Return Ret
        If PostData.Count < 1 Then Return Ret

        If Not (TrigInfo.DataIn Is Nothing) Then
            DeSerializeObject(TrigInfo.DataIn, trigger)
        End If

        Dim parts As Collections.Specialized.NameValueCollection
        parts = PostData
        Try
            For Each key As String In parts.Keys
                If key Is Nothing Then Continue For
                If String.IsNullOrEmpty(key.Trim) Then Continue For
                Select Case True
                    Case key.Contains("SomeValue_" & UID)
                        trigger.Add(CObj(parts(key)), key)
                End Select
            Next
            If Not SerializeObject(trigger, Ret.DataOut) Then
                Ret.sResult = Me.Name & " Error, Serialization failed. Signal Trigger not added."
                Return Ret
            End If
        Catch ex As Exception
            Ret.sResult = "ERROR, Exception in Trigger UI of " & Me.Name & ": " & ex.Message
            Return Ret
        End Try

        ' All OK
        Ret.sResult = ""
        Return Ret
    End Function

    Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String
        Dim stb As New StringBuilder
        Dim key As String
        Dim someValue As String = ""
        Dim UID As String = TrigInfo.UID.ToString

        If Not (TrigInfo.DataIn Is Nothing) Then
            DeSerializeObject(TrigInfo.DataIn, trigger)
        End If

        For Each key In trigger.Keys
            Select Case True
                Case key.Contains("SomeValue_" & UID)
                    someValue = trigger(key)
            End Select
        Next

        'We need different texts based on which trigger was used.
        Select Case TrigInfo.TANumber

            Case TriggerTypes.WithoutSubtriggers '= 1. The trigger without subtriggers only has one option:
                stb.Append(" the random value generator picks a number lower than " & someValue)

            Case TriggerTypes.WithSubtriggers '= 2. This trigger has subtriggers which also reflects on how the trigger is presented

                'let's start with the regular text for the trigger
                stb.Append(" the random value generator picks a number ")

                '... add the comparer (all subtriggers for the current trigger)
                Select Case TrigInfo.SubTANumber
                    Case SubTriggerTypes.LowerThan '= 1
                        stb.Append("lower than ")

                    Case SubTriggerTypes.EqualTo '= 2
                        stb.Append("equal to ")

                    Case SubTriggerTypes.HigherThan '3
                        stb.Append("higher than ")
                End Select

                '... and end with the selected value
                stb.Append(someValue)
        End Select


        hs.SaveEventsDevices()
        Return stb.ToString
    End Function

#End Region

#Region "Action Properties"

    Function ActionCount() As Integer
        Return actions.Count
    End Function

    ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String
        Get
            If ActionNumber > 0 AndAlso ActionNumber <= actions.Count Then
                Return Me.Name & ": " & actions.Keys(ActionNumber - 1)
            Else
                Return ""
            End If
        End Get
    End Property

    Private ReadOnly Property IPlugInAPI_Name As String Implements IPlugInAPI.Name
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public ReadOnly Property HSCOMPort As Boolean Implements IPlugInAPI.HSCOMPort
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Property ActionAdvancedMode As Boolean Implements IPlugInAPI.ActionAdvancedMode
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

    Private ReadOnly Property IPlugInAPI_ActionName(ActionNumber As Integer) As String Implements IPlugInAPI.ActionName
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IPlugInAPI_HasConditions(TriggerNumber As Integer) As Boolean Implements IPlugInAPI.HasConditions
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IPlugInAPI_HasTriggers As Boolean Implements IPlugInAPI.HasTriggers
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IPlugInAPI_TriggerCount As Integer Implements IPlugInAPI.TriggerCount
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IPlugInAPI_TriggerName(TriggerNumber As Integer) As String Implements IPlugInAPI.TriggerName
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IPlugInAPI_SubTriggerCount(TriggerNumber As Integer) As Integer Implements IPlugInAPI.SubTriggerCount
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IPlugInAPI_SubTriggerName(TriggerNumber As Integer, SubTriggerNumber As Integer) As String Implements IPlugInAPI.SubTriggerName
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Private ReadOnly Property IPlugInAPI_TriggerConfigured(TrigInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements IPlugInAPI.TriggerConfigured
        Get
            Throw New NotImplementedException()
        End Get
    End Property

    Public Property Condition(TrigInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements IPlugInAPI.Condition
        Get
            Throw New NotImplementedException()
        End Get
        Set(value As Boolean)
            Throw New NotImplementedException()
        End Set
    End Property

#End Region

#Region "Action Interface"

    Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean
        Dim houseCode As String = ""
        Dim deviceCode As String = ""
        Dim UID As String = ActInfo.UID.ToString

        Try
            If Not (ActInfo.DataIn Is Nothing) Then
                DeSerializeObject(ActInfo.DataIn, action)
            Else
                Return False
            End If

            For Each key As String In action.Keys
                Select Case True
                    Case key.Contains("HouseCodes_" & UID)
                        houseCode = action(key)
                    Case key.Contains("DeviceCodes_" & UID)
                        deviceCode = action(key)
                End Select
            Next

            Console.WriteLine("HandleAction, Command received with data: " & houseCode & ", " & deviceCode)
            SendCommand(houseCode, deviceCode) 'This could also return a value True/False if it was successful or not

        Catch ex As Exception
            Debug(9, System.Reflection.MethodBase.GetCurrentMethod().Name, "Error executing action: " & ex.Message)
        End Try
        Return True
    End Function

    Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean
        Dim Configured As Boolean = False
        Dim itemsConfigured As Integer = 0
        Dim itemsToConfigure As Integer = 2
        Dim UID As String = ActInfo.UID.ToString

        If Not (ActInfo.DataIn Is Nothing) Then
            DeSerializeObject(ActInfo.DataIn, action)
            For Each key In action.Keys
                Select Case True
                    Case key.Contains("HouseCodes_" & UID) AndAlso action(key) <> ""
                        itemsConfigured += 1
                    Case key.Contains("DeviceCodes_" & UID) AndAlso action(key) <> ""
                        itemsConfigured += 1
                End Select
            Next
            If itemsConfigured = itemsToConfigure Then Configured = True
        End If
        Return Configured
    End Function

    Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String
        Dim UID As String
        UID = ActInfo.UID.ToString
        Dim stb As New StringBuilder
        Dim Housecode As String = ""
        Dim DeviceCode As String = ""
        Dim dd As New clsJQuery.jqDropList("HouseCodes_" & UID & sUnique, Pagename, True)
        Dim dd1 As New clsJQuery.jqDropList("DeviceCodes_" & UID & sUnique, Pagename, True)
        Dim key As String


        dd.autoPostBack = True
        dd.AddItem("--Please Select--", "", False)
        dd1.autoPostBack = True
        dd1.AddItem("--Please Select--", "", False)

        If Not (ActInfo.DataIn Is Nothing) Then
            DeSerializeObject(ActInfo.DataIn, action)
        Else 'new event, so clean out the action object
            action = New Action
        End If

        For Each key In action.Keys
            Select Case True
                Case key.Contains("HouseCodes_" & UID)
                    Housecode = action(key)
                Case key.Contains("DeviceCodes_" & UID)
                    DeviceCode = action(key)
            End Select
        Next

        For Each C In "ABCDEFGHIJKLMNOP"
            dd.AddItem(C, C, (C = Housecode))
        Next

        stb.Append("Select House Code:")
        stb.Append(dd.Build)

        dd1.AddItem("All", "All", ("All" = DeviceCode))
        For i = 1 To 16
            dd1.AddItem(i.ToString, i.ToString, (i.ToString = DeviceCode))
        Next

        stb.Append("Select Unit Code:")
        stb.Append(dd1.Build)

        Return stb.ToString
    End Function

    Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, ByVal ActInfo As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn
        Dim ret As New HomeSeerAPI.IPlugInAPI.strMultiReturn
        Dim UID As String = ActInfo.UID.ToString

        ret.sResult = ""
        'HS: We cannot be passed info ByRef from HomeSeer, so turn right around and return this same value so that if we want, 
        '   we can exit here by returning 'Ret', all ready to go.  If in this procedure we need to change DataOut or TrigInfo,
        '   we can still do that.
        ret.DataOut = ActInfo.DataIn
        ret.TrigActInfo = ActInfo

        If PostData Is Nothing Then Return ret
        If PostData.Count < 1 Then Return ret

        If Not (ActInfo.DataIn Is Nothing) Then
            DeSerializeObject(ActInfo.DataIn, action)
        End If

        Dim parts As Collections.Specialized.NameValueCollection = PostData

        Try
            For Each key As String In parts.Keys
                If key Is Nothing Then Continue For
                If String.IsNullOrEmpty(key.Trim) Then Continue For
                Select Case True
                    Case key.Contains("HouseCodes_" & UID), key.Contains("DeviceCodes_" & UID)
                        action.Add(CObj(parts(key)), key)
                End Select
            Next
            If Not SerializeObject(action, ret.DataOut) Then
                ret.sResult = Me.Name & " Error, Serialization failed. Signal Action not added."
                Return ret
            End If
        Catch ex As Exception
            ret.sResult = "ERROR, Exception in Action UI of " & Me.Name & ": " & ex.Message
            Return ret
        End Try

        ' All OK
        ret.sResult = ""
        Return ret
    End Function

    Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String
        Dim stb As New StringBuilder
        Dim houseCode As String = ""
        Dim deviceCode As String = ""
        Dim UID As String = ActInfo.UID.ToString

        If Not (ActInfo.DataIn Is Nothing) Then
            DeSerializeObject(ActInfo.DataIn, action)
        End If

        For Each key As String In action.Keys
            Select Case True
                Case key.Contains("HouseCodes_" & UID)
                    houseCode = action(key)
                Case key.Contains("DeviceCodes_" & UID)
                    deviceCode = action(key)
            End Select
        Next

        stb.Append(" the system will do 'something' to a device with ")
        stb.Append("HouseCode " & houseCode & " ")
        If deviceCode = "ALL" Then
            stb.Append("for all Unitcodes")
        Else
            stb.Append("and Unitcode " & deviceCode)
        End If

        Return stb.ToString
    End Function

#End Region

#End Region

#Region "HomeSeer-Required Functions"


    Public Function Name() As String
        Return IFACE_NAME
    End Function

    Public Function AccessLevel() As Integer
        AccessLevel = 2
    End Function

#End Region

#Region "Web Page Processing"
    Private Function SelectPage(ByVal pageName As String) As Object
        Select Case pageName
            Case configPage.PageName
                Return configPage
                'Case statusPage.PageName
                'Return statusPage
            Case Else
                Return configPage
        End Select
        Return Nothing
    End Function

    Public Function PostBackProc(page As String, data As String, user As String, userRights As Integer) As String
        webPage = SelectPage(page)
        Return webPage.postBackProc(page, data, user, userRights)
    End Function

    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String
        ' build and return the actual page
        Debug(4, "Web", "Build and Serve Page " & pageName)
        webPage = SelectPage(pageName)
        Return webPage.GetPagePlugin(pageName, user, userRights, queryString)
    End Function

#End Region

#Region "Timers, trigging triggers"

    Private Function GetTriggerValue(ByVal key_to_find As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Object
        Dim _trigger As New Trigger

        'Loads the trigger from the serialized object (if it exists, and it should)
        If Not (TrigInfo.DataIn Is Nothing) Then
            DeSerializeObject(TrigInfo.DataIn, _trigger)
        End If

        'A trigger has "keys" with different stored values, let's go through them all.
        'In my sample we only have one key, which is "SomeValue"
        For Each key In _trigger.Keys
            Select Case True
                Case key.Contains(key_to_find & "_" & TrigInfo.UID)
                    'We found the correct key, so let's just return the value:
                    Return _trigger(key)
            End Select
        Next

        'Apparently we didn't find any matching keys in the trigger, so that's all we have to return
        Return Nothing
    End Function

#End Region

#Region "Device creation and management"

    Public Function Capabilities() As Integer Implements IPlugInAPI.Capabilities
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_AccessLevel() As Integer Implements IPlugInAPI.AccessLevel
        Throw New NotImplementedException()
    End Function

    Public Function SupportsMultipleInstances() As Boolean Implements IPlugInAPI.SupportsMultipleInstances
        Throw New NotImplementedException()
    End Function

    Public Function SupportsMultipleInstancesSingleEXE() As Boolean Implements IPlugInAPI.SupportsMultipleInstancesSingleEXE
        Throw New NotImplementedException()
    End Function

    Public Function SupportsAddDevice() As Boolean Implements IPlugInAPI.SupportsAddDevice
        Throw New NotImplementedException()
    End Function

    Public Function InstanceFriendlyName() As String Implements IPlugInAPI.InstanceFriendlyName
        Throw New NotImplementedException()
    End Function

    Public Function InterfaceStatus() As IPlugInAPI.strInterfaceStatus Implements IPlugInAPI.InterfaceStatus
        Throw New NotImplementedException()
    End Function

    Public Function GenPage(link As String) As String Implements IPlugInAPI.GenPage
        Throw New NotImplementedException()
    End Function

    Public Function PagePut(data As String) As String Implements IPlugInAPI.PagePut
        Throw New NotImplementedException()
    End Function

    Private Sub IPlugInAPI_ShutdownIO() Implements IPlugInAPI.ShutdownIO
        Throw New NotImplementedException()
    End Sub

    Public Function RaisesGenericCallbacks() As Boolean Implements IPlugInAPI.RaisesGenericCallbacks
        Throw New NotImplementedException()
    End Function

    Private Sub IPlugInAPI_SetIOMulti(colSend As List(Of CAPIControl)) Implements IPlugInAPI.SetIOMulti
        Throw New NotImplementedException()
    End Sub

    Private Function IPlugInAPI_InitIO(port As String) As String Implements IPlugInAPI.InitIO
        Throw New NotImplementedException()
    End Function

    Public Function PollDevice(dvref As Integer) As IPlugInAPI.PollResultInfo Implements IPlugInAPI.PollDevice
        Throw New NotImplementedException()
    End Function

    Public Function SupportsConfigDevice() As Boolean Implements IPlugInAPI.SupportsConfigDevice
        Throw New NotImplementedException()
    End Function

    Public Function SupportsConfigDeviceAll() As Boolean Implements IPlugInAPI.SupportsConfigDeviceAll
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_ConfigDevicePost(ref As Integer, data As String, user As String, userRights As Integer) As Enums.ConfigDevicePostReturn Implements IPlugInAPI.ConfigDevicePost
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_ConfigDevice(ref As Integer, user As String, userRights As Integer, newDevice As Boolean) As String Implements IPlugInAPI.ConfigDevice
        Throw New NotImplementedException()
    End Function

    Public Function Search(SearchString As String, RegEx As Boolean) As SearchReturn() Implements IPlugInAPI.Search
        Throw New NotImplementedException()
    End Function

    Public Function PluginFunction(procName As String, parms() As Object) As Object Implements IPlugInAPI.PluginFunction
        Throw New NotImplementedException()
    End Function

    Public Function PluginPropertyGet(procName As String, parms() As Object) As Object Implements IPlugInAPI.PluginPropertyGet
        Throw New NotImplementedException()
    End Function

    Public Sub PluginPropertySet(procName As String, value As Object) Implements IPlugInAPI.PluginPropertySet
        Throw New NotImplementedException()
    End Sub

    Public Sub SpeakIn(device As Integer, txt As String, w As Boolean, host As String) Implements IPlugInAPI.SpeakIn
        Throw New NotImplementedException()
    End Sub

    Private Function IPlugInAPI_ActionCount() As Integer Implements IPlugInAPI.ActionCount
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_ActionConfigured(ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements IPlugInAPI.ActionConfigured
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_ActionBuildUI(sUnique As String, ActInfo As IPlugInAPI.strTrigActInfo) As String Implements IPlugInAPI.ActionBuildUI
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_ActionProcessPostUI(PostData As NameValueCollection, TrigInfoIN As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn Implements IPlugInAPI.ActionProcessPostUI
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_ActionFormatUI(ActInfo As IPlugInAPI.strTrigActInfo) As String Implements IPlugInAPI.ActionFormatUI
        Throw New NotImplementedException()
    End Function

    Public Function ActionReferencesDevice(ActInfo As IPlugInAPI.strTrigActInfo, dvRef As Integer) As Boolean Implements IPlugInAPI.ActionReferencesDevice
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_HandleAction(ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements IPlugInAPI.HandleAction
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_TriggerBuildUI(sUnique As String, TrigInfo As IPlugInAPI.strTrigActInfo) As String Implements IPlugInAPI.TriggerBuildUI
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_TriggerProcessPostUI(PostData As NameValueCollection, TrigInfoIN As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn Implements IPlugInAPI.TriggerProcessPostUI
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_TriggerFormatUI(TrigInfo As IPlugInAPI.strTrigActInfo) As String Implements IPlugInAPI.TriggerFormatUI
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_TriggerTrue(TrigInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements IPlugInAPI.TriggerTrue
        Throw New NotImplementedException()
    End Function

    Public Function TriggerReferencesDevice(TrigInfo As IPlugInAPI.strTrigActInfo, dvRef As Integer) As Boolean Implements IPlugInAPI.TriggerReferencesDevice
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_GetPagePlugin(page As String, user As String, userRights As Integer, queryString As String) As String Implements IPlugInAPI.GetPagePlugin
        Throw New NotImplementedException()
    End Function

    Private Function IPlugInAPI_PostBackProc(page As String, data As String, user As String, userRights As Integer) As String Implements IPlugInAPI.PostBackProc
        Throw New NotImplementedException()
    End Function
#End Region




End Class
