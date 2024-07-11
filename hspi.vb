Imports System
Imports Scheduler
Imports HomeSeerAPI
Imports HSCF.Communication.Scs.Communication.EndPoints.Tcp
Imports HSCF.Communication.ScsServices.Client
Imports HSCF.Communication.ScsServices.Service
Imports System.Reflection

Public Class HSPI
    Implements IPlugInAPI        ' this API is required for ALL plugins
    'Implements IThermostatAPI   ' add this API if this plugin supports thermostats

    ''' <summary>
    ''' This will allow scripts to call any function in your plugin by name.
    ''' Used to allow scripting of your plugin functions.
    ''' </summary>
    ''' <param name="proc">The name of the process/sub/function to call</param>
    ''' <param name="parms">Parameters passed as an object array</param>
    ''' <returns></returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/custom_functions.htm</remarks>
    Public Function PluginFunction(ByVal proc As String, ByVal parms() As Object) As Object Implements IPlugInAPI.PluginFunction
        Try
            Dim [type] As Type = Me.GetType
            Dim methodInfo As MethodInfo = [type].GetMethod(proc)
            If methodInfo Is Nothing Then
                Debug(9, "PlugInFunction", "Method " & proc & " does not exist in this plugin.")
                Return Nothing
            End If
            Return (methodInfo.Invoke(Me, parms))
        Catch ex As Exception
            Debug(9, "PlugInFunction", "Error in PluginProc: " & ex.Message)
        End Try

        Return Nothing
    End Function


    ''' <summary>
    ''' This will allow scripts to call any property in your plugin by name.
    ''' Used to allow scripting of your plugin properties.
    ''' </summary>
    ''' <param name="proc">The name of the Property to get</param>
    ''' <param name="parms">Parameters passed as an object array</param>
    ''' <returns></returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/custom_functions.htm</remarks>
    Public Function PluginPropertyGet(ByVal proc As String, parms() As Object) As Object Implements IPlugInAPI.PluginPropertyGet
        Try
            Dim [type] As Type = Me.GetType
            Dim propertyInfo As PropertyInfo = [type].GetProperty(proc)
            If propertyInfo Is Nothing Then
                Debug(9, System.Reflection.MethodBase.GetCurrentMethod().Name, "Property " & proc & " does not exist in this plugin ")
                Return Nothing
            End If
            Return propertyInfo.GetValue(Me, parms)
        Catch ex As Exception
            Debug(9, System.Reflection.MethodBase.GetCurrentMethod().Name, "Error in PluginPropertyGet: " & ex.Message)
        End Try

        Return Nothing
    End Function

    ''' <summary>
    ''' This will allow scripts to set any property in your plugin by name.
    ''' Used to allow scripting of your plugin properties.
    ''' </summary>
    ''' <param name="proc">The name of the property to set</param>
    ''' <param name="value">The new value for the property</param>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/custom_functions.htm</remarks>
    Public Sub PluginPropertySet(ByVal proc As String, value As Object) Implements IPlugInAPI.PluginPropertySet
        Try
            Dim [type] As Type = Me.GetType
            Dim propertyInfo As PropertyInfo = [type].GetProperty(proc)
            If propertyInfo Is Nothing Then
                Debug(9, System.Reflection.MethodBase.GetCurrentMethod().Name, "Property " & proc & " does not exist in this plugin.")
            End If
            propertyInfo.SetValue(Me, value, Nothing)
        Catch ex As Exception
            Debug(9, System.Reflection.MethodBase.GetCurrentMethod().Name, "Error in PluginPropertySet: " & ex.Message)
        End Try
    End Sub

    ''' <summary>
    ''' Returns the name of your plug-in. This is used to identify your plug-in to HomeSeer and your users. Keep the name to 16 characters or less.
    ''' Do not access any hardware in this function as HomeSeer will call this function.
    ''' Do NOT use special characters in your plug-in name with the exception of "-", ".", and " " (space).
    ''' </summary>
    ''' <returns>Plugin name</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/name.htm</remarks>
    Public ReadOnly Property Name As String Implements HomeSeerAPI.IPlugInAPI.Name
        Get
            Return plugin.Name
        End Get
    End Property

    ''' <summary>
    ''' Determines if this plugin uses a COM port as configured on "Plug-ins -> Managepage" page, /interfaces
    ''' </summary>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/hscomport.htm</remarks>
    Public ReadOnly Property HSCOMPort As Boolean Implements HomeSeerAPI.IPlugInAPI.HSCOMPort
        Get
            Return False
        End Get
    End Property

    ''' <summary>
    ''' Return the API's that this plug-in supports. This is a bit field. All plug-ins must have bit 3 set for I/O. This value is 4.
    ''' </summary>
    ''' <returns>HomeSeerAPI.Enums.eCapabilities</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/capabilities.htm</remarks>
    Public Function Capabilities() As Integer Implements HomeSeerAPI.IPlugInAPI.Capabilities
        Return HomeSeerAPI.Enums.eCapabilities.CA_IO
    End Function

    ''' <summary>
    ''' Return the access level of this plug-in. Access level is the licensing mode.
    ''' </summary>
    ''' <returns>
    ''' 1 = Plug-in is not licensed and may be enabled and run without purchasing a license. Use this value for free plug-ins.
    ''' 2 = Plug-in is licensed and a user must purchase a license in order to use this plug-in. When the plug-in is first enabled, it will will run as a trial for 30 days.</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/accesslevel.htm</remarks>
    Public Function AccessLevel() As Integer Implements HomeSeerAPI.IPlugInAPI.AccessLevel
        Return plugin.AccessLevel
    End Function

    ''' <summary>
    ''' HomeSeer may call this function at any time to get the status of the plug-in. Normally it is displayed on the Interfaces page. The return is an object that represents the status. The object is of type HomeSeerAPI.IPlugInAPI.strInterfaceStatus
    ''' </summary>
    ''' <returns>The interface status as "strInterfaceStatus"</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/interfacestatus1.htm</remarks>
    Public Function InterfaceStatus() As HomeSeerAPI.IPlugInAPI.strInterfaceStatus Implements HomeSeerAPI.IPlugInAPI.InterfaceStatus
        Dim es As New IPlugInAPI.strInterfaceStatus
        es.intStatus = IPlugInAPI.enumInterfaceStatus.OK
        Return es
    End Function

    ''' <summary>
    ''' Return TRUE if the plug-in supports multiple instances. The plug-in may be launched multiple times and will be passed a unique instance name as a command line parameter to the Main function. The plug-in then needs to associate all local status with this particular instance.
    ''' The instance is passed to the main function in the plugin and should be saved for future reference.
    ''' </summary>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/supportsmultipleinstances.htm</remarks>
    Public Function SupportsMultipleInstances() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstances
        Return False
    End Function

    ''' <summary>
    ''' (Documentation not found)
    ''' </summary>
    ''' <returns>True/False</returns>
    ''' <remarks></remarks>
    Public Function SupportsMultipleInstancesSingleEXE() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsMultipleInstancesSingleEXE
        Return False
    End Function

    ''' <summary>
    ''' Returns the instance name of this instance of the plug-in. Only valid if SupportsMultipleInstances returns TRUE. The instance is set when the plug-in is started, it is passed as a command line parameter
    ''' </summary>
    ''' <returns>The instance name</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/instancefriendlyname.htm</remarks>
    Public Function InstanceFriendlyName() As String Implements HomeSeerAPI.IPlugInAPI.InstanceFriendlyName
        Return ""
    End Function

    ''' <summary>
    ''' If your plugin is set to start when HomeSeer starts, or is enabled from the interfaces page, then this function will be called to initialize your plugin. If you returned TRUE from HSComPort then the port number as configured in HomeSeer will be passed to this function. Here you should initialize your plugin fully. The hs object is available to you to call the HomeSeer scripting API as well as the callback object so you can call into the HomeSeer plugin API.  HomeSeer's startup routine waits for this function to return a result, so it is important that you try to exit this procedure quickly.  If your hardware has a long initialization process, you can check the configuration in InitIO and if everything is set up correctly, start a separate thread to initialize the hardware and exit InitIO.  If you encounter an error, you can always use InterfaceStatus to indicate this.
    ''' </summary>
    ''' <param name="port">Optional COM port passed by HomeSeer, used if HSComPort is set to TRUE</param>
    ''' <returns>(Empty string)</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/initialization_-_when_used.htm</remarks>
    Public Function InitIO(ByVal port As String) As String Implements HomeSeerAPI.IPlugInAPI.InitIO
        Return plugin.InitIO(port)
    End Function

    ''' <summary>
    ''' HSTouch uses the Generic Event callback in some music plug-ins so that it can be notified of when a song changes, rather than having to repeatedly query the music plug-in for the current song status.
    ''' If this property is present (and returns True), especially in a Music plug-in, then HSTouch (and other plug-ins) will know that your HSEvent procedure can handle generic callbacks.  
    ''' </summary>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/raisesgenericcallbacks.htm</remarks>
    Public Function RaisesGenericCallbacks() As Boolean Implements HomeSeerAPI.IPlugInAPI.RaisesGenericCallbacks
        Return True
    End Function

    ''' <summary>
    ''' SetIOMulti is called by HomeSeer when a device that your plug-in owns is controlled.
    ''' Your plug-in owns a device when it's INTERFACE property is set to the name of your plug
    ''' </summary>
    ''' <param name="colSend">
    ''' A collection of CAPIControl objects, one object for each device that needs to be controlled.
    ''' Look at the ControlValue property to get the value that device needs to be set to.</param>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/setio.htm</remarks>
    Public Sub SetIOMulti(colSend As System.Collections.Generic.List(Of HomeSeerAPI.CAPI.CAPIControl)) Implements HomeSeerAPI.IPlugInAPI.SetIOMulti
        plugin.SetIOMulti(colSend)
    End Sub

    ''' <summary>
    ''' When HomeSeer shuts down or a plug-in is disabled from the interfaces page this function is then called. You should terminate any threads that you started, close any COM ports or TCP connections and release memory.
    ''' After you return from this function the plugin EXE will terminate and must be allowed to terminate cleanly.
    ''' </summary>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/initialization_-_when_used.htm</remarks>
    Public Sub ShutdownIO() Implements HomeSeerAPI.IPlugInAPI.ShutdownIO
        plugin.ShutdownIO()
    End Sub

    Public Sub HSEvent(ByVal EventType As Enums.HSEvent, ByVal parms() As Object) Implements HomeSeerAPI.IPlugInAPI.HSEvent
        Dim dAddress As String = ""
        Dim dValue As Double = 0
        Dim dRef As Integer = 0
        Dim dName As String = ""
        Dim dLocation As String = ""
        Dim dLocation2 As String = ""
        Dim dString As String = ""
        Dim dIsOn As Boolean = False
        Dim mqttString As String = ""
        Dim Payload As String = ""
        Dim d As Scheduler.Classes.DeviceClass

        'Debug("HSPI", "HSEVENT:" + EventType.ToString)
        GlobalVariables.StatsHSEvents = GlobalVariables.StatsHSEvents + 1

        Select Case EventType
            Case Enums.HSEvent.STRING_CHANGE
                dAddress = parms(1)
                dString = parms(2)
                dRef = parms(3)
                'Debug(2, "HSEvent", "Device Ref " + dRef.ToString + "(" + dAddress.ToString + ") Has changed its String to [" + dString.ToString + "]")
                'If GlobalVariables.MQTTDevices.ContainsKey(dRef) Then
                'Debug(2, "HSPI", "Our target device changed its String to " + dString.ToString)
                'End If
            Case Enums.HSEvent.VALUE_CHANGE
                dAddress = parms(1)
                dValue = parms(2)
                dRef = parms(4)
                dName = hs.DeviceName(dRef)

                Dim OurDevice As DeviceInfo = GlobalVariables.KMTDevices.Find(Function(p) p.DeviceID = dRef)
                If Not OurDevice Is Nothing Then
                    Debug(2, "HSPI", "Ignored the previous " + GlobalVariables.IgnoredEvents.ToString + " events")
                    dString = OurDevice.DeviceString
                    GlobalVariables.IgnoredEvents = 0
                    If OurDevice.Type = 1 Then
                        If OurDevice.Active Then
                            dIsOn = hs.IsON(dRef)
                            If dIsOn Then Payload = "1" Else Payload = "0"
                            Debug(3, "HSPI", "Device " + dRef.ToString + " = " + dValue.ToString + " and isOn=" + dIsOn.ToString + " Queued")
                            GlobalVariables.eQUEUE.Enqueue("SETUN|" + dRef.ToString + "|" + Payload)
                        Else
                            Debug(5, "HSPI", "Device " + dRef.ToString + " is OFFLINE - ignoring command")
                        End If
                    Else
                        Debug(3, "HSPI", "Device " + dRef.ToString + " Is a Sensor - ignoring change")
                    End If

                Else
                    Debug(0, "HSEvent", "Device Changed, but not ours " + dRef.ToString + "[" + dName + "] (" + dAddress.ToString + ")")
                    GlobalVariables.IgnoredEvents = GlobalVariables.IgnoredEvents + 1
                End If
        End Select
    End Sub


    ''' <summary>
    ''' If a device is owned by your plug-in (interface property set to the name of the plug-in) and the device's status_support property is set to True, then this procedure will be called in your plug-in when the device's status is being polled, such as when the user clicks "Poll Devices" on the device status page.
    ''' </summary>
    ''' <param name="dvref">The device reference</param>
    ''' <returns>new value of the device (but apparently not anymore) </returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/polldevice.htm</remarks>
    Public Function PollDevice(ByVal dvref As Integer) As IPlugInAPI.PollResultInfo Implements HomeSeerAPI.IPlugInAPI.PollDevice
    End Function

    ''' <summary>
    ''' This function is available for the ease of converting older HS2 plugins, however, it is desirable to use the new clsPageBuilder class for all new development.
    ''' </summary>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/genpage.htm</remarks>
    Public Function GenPage(ByVal link As String) As String Implements HomeSeerAPI.IPlugInAPI.GenPage
        Return Nothing
    End Function

    ''' <summary>
    ''' This function is available for the ease of converting older HS2 plugins, however, it is desirable to use the new clsPageBuilder class for all new development.
    ''' </summary>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/pageput.htm</remarks>
    Public Function PagePut(ByVal data As String) As String Implements HomeSeerAPI.IPlugInAPI.PagePut
        Return Nothing
    End Function

    ''' <summary>
    ''' http://homeseer.com/support/homeseer/HS3/SDK/getpageplugin.htm
    ''' </summary>
    ''' <param name="pageName">The name of the page as passed to the hs.RegisterLink function</param>
    ''' <param name="user">The name of the user</param>
    ''' <param name="userRights">The associated user rights</param>
    ''' <param name="queryString">Extra URI queries (?)</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function GetPagePlugin(ByVal pageName As String, ByVal user As String, ByVal userRights As Integer, ByVal queryString As String) As String Implements HomeSeerAPI.IPlugInAPI.GetPagePlugin
        Return plugin.GetPagePlugin(pageName, user, userRights, queryString)
    End Function

    ''' <summary>
    ''' Web pages that use the clsPageBuilder class and registered with hs.RegisterLink and hs.RegisterConfigLink will then be called through this function.
    ''' A complete page needs to be created and returned.
    ''' </summary>
    ''' <param name="pageName">The name of the page as passed to the hs.RegisterLink function</param>
    ''' <param name="data">the post data</param>
    ''' <param name="user">The name of the user</param>
    ''' <param name="userRights">The associated user rights</param>
    ''' <returns></returns>
    ''' <remarks></remarks>
    Public Function PostBackProc(ByVal pageName As String, ByVal data As String, ByVal user As String, ByVal userRights As Integer) As String Implements HomeSeerAPI.IPlugInAPI.PostBackProc
        Return plugin.PostBackProc(pageName, data, user, userRights)
    End Function

    ''' <summary>
    ''' The HomeSeer events page has an option to set the editing mode to "Advanced Mode". This is typically used to enable options that may only be of interest to advanced users or programmers. The Set in this function is called when advanced mode is enabled. Your plug-in can also enable this mode if an advanced selection was saved and needs to be displayed.
    ''' </summary>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/actionadvancedmode.htm</remarks>
    Public Property ActionAdvancedMode As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionAdvancedMode
        Set(ByVal value As Boolean)
            _actionAdvancedMode = value
        End Set
        Get
            Return _actionAdvancedMode
        End Get
    End Property
    Private _actionAdvancedMode As Boolean

    ''' <summary>
    ''' This function is called from the HomeSeer event page when an event is in edit mode.
    ''' Your plug-in needs to return HTML controls so the user can make action selections.
    ''' Normally this is one of the HomeSeer jquery controls such as a clsJquery.jqueryCheckbox.
    ''' </summary>
    ''' <param name="sUnique">A unique string that can be used with your HTML controls to identify the control. All controls need to have a unique ID.</param>
    ''' <param name="ActInfo">Object that contains information about the action like current selections</param>
    ''' <returns> HTML controls that need to be displayed so the user can select the action parameters.</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/actionbuildui.htm</remarks>
    Public Function ActionBuildUI(ByVal sUnique As String, ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionBuildUI
        Return plugin.ActionBuildUI(sUnique, ActInfo)
    End Function

    ''' <summary>
    ''' Return TRUE if the given action is configured properly. There may be times when a user can select invalid selections for the action and in this case you would return FALSE so HomeSeer will not allow the action to be saved.
    ''' </summary>
    ''' <param name="ActInfo">Object that contains information about the action like current selections.</param>
    ''' <returns>Return TRUE if the given action is configured properly.</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/actionconfigured.htm</remarks>
    Public Function ActionConfigured(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionConfigured
        Return plugin.ActionConfigured(ActInfo)
    End Function

    ''' <summary>
    ''' Return True if the given device is referenced in the given action.
    ''' </summary>
    ''' <param name="ActInfo">The activity information</param>
    ''' <param name="dvRef">The device reference to check</param>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/actionreferencesdevice.htm</remarks>
    Public Function ActionReferencesDevice(ByVal ActInfo As IPlugInAPI.strTrigActInfo, ByVal dvRef As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.ActionReferencesDevice
        Return False

        'The exmample from the documentation doesn't work, so there must be a different way to figure that out.
    End Function

    ''' <summary>
    ''' "Body of text here"... Okay, my guess:
    ''' This formats the chosen action when the action is "minimized" based on the user selected options
    ''' </summary>
    ''' <param name="ActInfo">Information from the current activity as "strTrigActInfo" (which is funny, as it isn't a string at all)</param>
    ''' <returns>Simple string. Possibly HTML-formated.</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/actionformatui.htm</remarks>
    Public Function ActionFormatUI(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.ActionFormatUI
        Return plugin.ActionFormatUI(ActInfo)
    End Function

    ''' <summary>
    ''' Return the name of the action given an action number. The name of the action will be displayed in the HomeSeer events actions list.
    ''' </summary>
    ''' <param name="ActionNumber">The number of the action. Each action is numbered, starting at 1. (BUT WHY 1?!)</param>
    ''' <returns>The action name from the 1-based index</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/actionname.htm</remarks>
    Public ReadOnly Property ActionName(ByVal ActionNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.ActionName
        Get
            Return plugin.ActionName(ActionNumber)
        End Get
    End Property

    ''' <summary>
    ''' When a user edits your event actions in the HomeSeer events, this function is called to process the selections.
    ''' </summary>
    ''' <param name="PostData">A collection of name value pairs that include the user's selections.</param>
    ''' <param name="TrigInfoIN">Object that contains information about the action as "strTrigActInfo" (which is funny, as it isn't a string at all)</param>
    ''' <returns>Object the holds the parsed information for the action. HomeSeer will save this information for you in the database.</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/actionprocesspostui.htm</remarks>
    Public Function ActionProcessPostUI(ByVal PostData As Collections.Specialized.NameValueCollection, ByVal TrigInfoIN As IPlugInAPI.strTrigActInfo) As IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.ActionProcessPostUI
        Return plugin.ActionProcessPostUI(PostData, TrigInfoIN)
    End Function

    ''' <summary>
    ''' Return the number of actions the plugin supports.
    ''' </summary>
    ''' <returns>The plugin count</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/actioncount.htm</remarks>
    Public Function ActionCount() As Integer Implements HomeSeerAPI.IPlugInAPI.ActionCount
        Return plugin.ActionCount
    End Function

    ''' <summary>
    ''' HomeSeer will set this to TRUE if the trigger is being used as a CONDITION.  Check this value in BuildUI and other procedures to change how the trigger is rendered if it is being used as a condition or a trigger.
    ''' </summary>
    ''' <param name="TrigInfo">The trigger information</param>
    ''' <value>True/False</value>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/condition.htm</remarks>
    Public Property Condition(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.Condition
        Set(ByVal value As Boolean)
            _condition = value
        End Set
        Get
            Return _condition
        End Get
    End Property
    Dim _condition As Boolean

    ''' <summary>
    ''' When an event is triggered, this function is called to carry out the selected action.
    ''' </summary>
    ''' <param name="ActInfo">Use the ActInfo parameter to determine what action needs to be executed then execute this action.</param>
    ''' <returns>Return TRUE if the action was executed successfully, else FALSE if there was an error.</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/handleaction.htm</remarks>
    Public Function HandleAction(ByVal ActInfo As IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.HandleAction
        Return plugin.HandleAction(ActInfo)
    End Function

    ''' <summary>
    ''' Return True if the given trigger can also be used as a condition, for the given trigger number.
    ''' </summary>
    ''' <param name="TriggerNumber">The trigger number (surely 1 based)</param>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/hasconditions.htm</remarks>
    Public ReadOnly Property HasConditions(ByVal TriggerNumber As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.HasConditions
        Get
            Return plugin.HasConditions(TriggerNumber)
        End Get
    End Property

    ''' <summary>
    ''' Triggers notify HomeSeer of trigger states using TriggerFire , but Triggers can also be conditions, and that is where this is used.
    ''' If this function is called, TrigInfo will contain the trigger information pertaining to a trigger used as a condition.
    ''' </summary>
    ''' <param name="TrigInfo">The trigger information</param>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/triggertrue.htm</remarks>
    Public Function TriggerTrue(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerTrue
        Return plugin.TriggerTrue(TrigInfo)
    End Function

    ''' <summary>
    ''' Return True if your plugin contains any triggers, else return false.
    ''' </summary>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/hastriggers.htm</remarks>
    Public ReadOnly Property HasTriggers() As Boolean Implements HomeSeerAPI.IPlugInAPI.HasTriggers
        Get
            Return plugin.HasTriggers
        End Get
    End Property

    ''' <summary>
    ''' Return the number of sub triggers your plugin supports.
    ''' </summary>
    ''' <param name="TriggerNumber">The trigger number</param>
    ''' <returns>Integer</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/subtriggercount.htm</remarks>
    Public ReadOnly Property SubTriggerCount(ByVal TriggerNumber As Integer) As Integer Implements HomeSeerAPI.IPlugInAPI.SubTriggerCount
        Get
            Return plugin.SubTriggerCount(TriggerNumber)
        End Get
    End Property

    ''' <summary>
    ''' Return the text name of the sub trigger given its trigger number and sub trigger number.
    ''' </summary>
    ''' <param name="TriggerNumber">Integer</param>
    ''' <param name="SubTriggerNumber">Integer</param>
    ''' <returns>SubTriggerName String</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/subtriggername.htm</remarks>
    Public ReadOnly Property SubTriggerName(ByVal TriggerNumber As Integer, ByVal SubTriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.SubTriggerName
        Get
            Return plugin.SubTriggerName(TriggerNumber, SubTriggerNumber)
        End Get
    End Property

    ''' <summary>
    ''' Return HTML controls for a given trigger.
    ''' </summary>
    ''' <param name="sUnique">An unique string</param>
    ''' <param name="TrigInfo">The trigger information</param>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/triggerbuildui.htm</remarks>
    Public Function TriggerBuildUI(ByVal sUnique As String, ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerBuildUI
        Return plugin.TriggerBuildUI(sUnique, TrigInfo)
    End Function

    ''' <summary>
    ''' Given a strTrigActInfo object detect if this this trigger is configured properly, if so, return True, else False.
    ''' </summary>
    ''' <param name="TrigInfo">The trigger information</param>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/triggerconfigured.htm</remarks>
    Public ReadOnly Property TriggerConfigured(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerConfigured
        Get
            Return plugin.TriggerConfigured(TrigInfo)
        End Get
    End Property

    ''' <summary>
    ''' Return True if the given device is referenced by the given trigger.
    ''' </summary>
    ''' <param name="TrigInfo">The trigger information</param>
    ''' <param name="dvRef">The device reference</param>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/triggerreferencesdevice.htm</remarks>
    Public Function TriggerReferencesDevice(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo, ByVal dvRef As Integer) As Boolean Implements HomeSeerAPI.IPlugInAPI.TriggerReferencesDevice
        Return False
    End Function

    ''' <summary>
    ''' After the trigger has been configured, this function is called in your plugin to display the configured trigger. Return text that describes the given trigger.
    ''' </summary>
    ''' <param name="TrigInfo">Information of the trigger</param>
    ''' <returns>Text describing the trigger</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/triggerformatui.htm</remarks>
    Public Function TriggerFormatUI(ByVal TrigInfo As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As String Implements HomeSeerAPI.IPlugInAPI.TriggerFormatUI
        Return plugin.TriggerFormatUI(TrigInfo)
    End Function

    ''' <summary>
    ''' Return the name of the given trigger based on the trigger number passed.
    ''' </summary>
    ''' <param name="TriggerNumber">Integer</param>
    ''' <returns>String</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/triggername.htm</remarks>
    Public ReadOnly Property TriggerName(ByVal TriggerNumber As Integer) As String Implements HomeSeerAPI.IPlugInAPI.TriggerName
        Get
            Return plugin.TriggerName(TriggerNumber)
        End Get
    End Property

    ''' <summary>
    ''' Process a post from the events web page when a user modifies any of the controls related to a plugin trigger. After processing the user selctions, create and return a strMultiReturn object.
    ''' </summary>
    ''' <param name="PostData">The PostData as NameValueCollection</param>
    ''' <param name="TrigInfoIn">The trigger information</param>
    ''' <returns>A structure, which is used in the Trigger and Action ProcessPostUI procedures, which not only communications trigger and action information through TrigActInfo which is strTrigActInfo , but provides an array of Byte where an updated/serialized trigger or action object from your plug-in can be stored.  See TriggerProcessPostUI and ActionProcessPostUI for more details.</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/triggerprocesspostui.htm</remarks>
    Public Function TriggerProcessPostUI(ByVal PostData As System.Collections.Specialized.NameValueCollection, ByVal TrigInfoIn As HomeSeerAPI.IPlugInAPI.strTrigActInfo) As HomeSeerAPI.IPlugInAPI.strMultiReturn Implements HomeSeerAPI.IPlugInAPI.TriggerProcessPostUI
        Return plugin.TriggerProcessPostUI(PostData, TrigInfoIn)
    End Function

    ''' <summary>
    ''' Return the number of triggers that the plugin supports.
    ''' </summary>
    ''' <returns>Integer</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/triggercount.htm</remarks>
    Public ReadOnly Property TriggerCount As Integer Implements HomeSeerAPI.IPlugInAPI.TriggerCount
        Get
            Return plugin.TriggerCount
        End Get
    End Property

    ''' <summary>
    ''' Return TRUE if your plug-in allows for configuration of your devices via the device utility page. This will allow you to generate some HTML controls that will be displayed to the user for modifying the device.
    ''' </summary>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/supportsconfigdevice.htm</remarks>
    Public Function SupportsConfigDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDevice
        Return True
    End Function

    ''' <summary>
    ''' If your plug-in manages all devices in the system, you can return TRUE from this function. Your configuration page will be available for all devices.
    ''' Should, in my opinion, be avoided if possible. -Moskus
    ''' </summary>
    ''' <returns>True/False</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/supportsconfigdeviceall.htm</remarks>
    Public Function SupportsConfigDeviceAll() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsConfigDeviceAll
        Return False
    End Function

    ''' <summary>
    ''' Return TRUE if the plugin supports the ability to add devices through the Add Device link on the device utility page. If TRUE a tab appears on the add device page that allows the user to configure specific options for the new device.
    ''' When ConfigDevice is called the newDevice parameter will be True if this is the first time the device config screen is being displayed and a new device is being created. See ConfigDevicePost  for more information.
    ''' </summary>
    ''' <returns>True/False</returns>
    ''' <remarks></remarks>
    Public Function SupportsAddDevice() As Boolean Implements HomeSeerAPI.IPlugInAPI.SupportsAddDevice
        Return False
    End Function

    ''' <summary>
    ''' This function is called when a user posts information from your plugin tab on the device utility page
    ''' </summary>
    ''' <param name="ref">The device reference</param>
    ''' <param name="data">query string data posted to the web server (name/value pairs from controls on the page)</param>
    ''' <param name="user">The user that is logged into the server and viewing the page</param>
    ''' <param name="userRights">The rights of the logged in user</param>
    ''' <returns>
    ''' DoneAndSave = 1            Any changes to the config are saved and the page is closed and the user it returned to the device utility page
    ''' DoneAndCancel = 2          Changes are not saved and the user is returned to the device utility page
    ''' DoneAndCancelAndStay = 3   No action is taken, the user remains on the plugin tab
    ''' CallbackOnce = 4           Your plugin ConfigDevice is called so your tab can be refereshed, the user stays on the plugin tab
    ''' CallbackTimer = 5          Your plugin ConfigDevice is called and a page timer is called so ConfigDevicePost is called back every 2 seconds
    ''' </returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/configdevicepost.htm</remarks>
    Function ConfigDevicePost(ByVal ref As Integer, ByVal data As String, ByVal user As String, ByVal userRights As Integer) As Enums.ConfigDevicePostReturn Implements IPlugInAPI.ConfigDevicePost
        Return plugin.ConfigDevicePost(ref, data, user, userRights)
    End Function

    ''' <summary>
    ''' If SupportsConfigDevice returns TRUE, this function will be called when the device properties are displayed for your device. This functions creates a tab for each plug-in that controls the device.
    ''' 
    ''' If the newDevice parameter is TRUE, the user is adding a new device from the HomeSeer user interface.
    ''' If you return TRUE from your SupportsAddDevice then ConfigDevice will be called when a user is creating a new device.
    ''' Your tab will appear and you can supply controls for the user to create a new device for your plugin. When your ConfigDevicePost is called you will need to get a reference to the device using the past ref number and then take ownership of the device by setting the interface property of the device to the name of your plugin. You can also set any other properties on the device as needed.
    ''' </summary>
    ''' <param name="ref">The device reference number</param>
    ''' <param name="user">The user that is logged into the server and viewing the page</param>
    ''' <param name="userRights">The rights of the logged in user</param>
    ''' <param name="newDevice">True if this a new device being created for the first time. In this case, the device configuration dialog may present different information than when simply editing an existing device.</param>
    ''' <returns>A string containing HTML to be displayed. Return an empty string if there is not configuration needed.</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/configdevice.htm</remarks>
    Function ConfigDevice(ByVal ref As Integer, ByVal user As String, ByVal userRights As Integer, newDevice As Boolean) As String Implements IPlugInAPI.ConfigDevice
        Return plugin.ConfigDevice(ref, user, userRights, newDevice)
    End Function

    ''' <summary>
    ''' This procedure will be called in your plug-in by HomeSeer whenever the user uses the search function of HomeSeer, and your plug-in is loaded and initialized.  Unlike ActionReferencesDevice and TriggerReferencesDevice, this search is not being specific to a device, it is meant to find a match anywhere in the resources managed by your plug-in.  This could include any textual field or object name that is utilized by the plug-in.
    ''' </summary>
    ''' <param name="SearchString">The string to search for</param>
    ''' <param name="RegEx">True/False</param>
    ''' <returns>HomeSeerAPI.SearchReturn. See links in Remarks.</returns>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/search.htm</remarks>
    Public Function Search(SearchString As String, RegEx As Boolean) As HomeSeerAPI.SearchReturn() Implements HomeSeerAPI.IPlugInAPI.Search
        Return Nothing
    End Function

    ''' <summary>
    ''' If your plug-in is registered as a Speak proxy plug-in, then when HomeSeer is asked to speak something, it will pass the speak information to your plug-in using this procedure.
    ''' </summary>
    ''' <param name="device">This is the device that is to be used for the speaking.  In older versions of HomeSeer, this value was used to indicate the sound card to use, and if it was over 100, then it indicated that it was speaking for HomeSeer Phone (device - 100 = phone line), or the WAV audio device to use.  Although this is still used for HomeSeer Phone, speaks for HomeSeer phone are never proxied and so values >= 100 should never been seen in the device parameter.  Pass the device parameter unchanged to SpeakProxy.</param>
    ''' <param name="txt">This is the text to be spoken, or if it is a WAV file to be played, then the characters ":\" will be found starting at position 2 of the string as playing a WAV file with the speak command in HomeSeer REQUIRES a fully qualified path and filename of the WAV file to play.</param>
    ''' <param name="w">Wait. This parameter tells HomeSeer whether to continue processing commands immediately or to wait until the speak command is finished - pass this parameter unchanged to SpeakProxy.</param>
    ''' <param name="host"> This is a list of host:instances to speak or play the WAV file on.  An empty string or a single asterisk (*) indicates all connected speaker clients on all hosts.  Normally this parameter is passed to SpeakProxy unchanged.</param>
    ''' <remarks>http://homeseer.com/support/homeseer/HS3/SDK/speakin.htm</remarks>
    Public Sub SpeakIn(device As Integer, txt As String, w As Boolean, host As String) Implements HomeSeerAPI.IPlugInAPI.SpeakIn

    End Sub

    Public Shared Sub SetDevice(DeviceID As Long, Command As String)

        Dim Capi As HomeSeerAPI.CAPI.CAPIControl
        Capi = hs.CAPIGetSingleControl(DeviceID, True, Command, False, True)
        If Capi IsNot Nothing Then
            Capi.Do_Update = False
            hs.CAPIControlHandler(Capi)
            'hs.WriteLog(GlobalVariables.ModuleName, "Set Device " + DeviceID.ToString + " to " + Command)
            Debug(4, "HSPI", "Set Device " + DeviceID.ToString + " to " + Command)
            GlobalVariables.KMTDevices.Find(Function(p) p.DeviceID = DeviceID).LastSeen = Now
            GlobalVariables.KMTDevices.Find(Function(p) p.DeviceID = DeviceID).Active = True
        Else
            Debug(4, "HSPI", "Device " + DeviceID.ToString + " cannot be found by CAPI")
            AddWarning(3, "Device " + DeviceID.ToString + " Command " + Command + " cannot be found by CAPI controller")
        End If
    End Sub

    Public Shared Sub SetDeviceValue(DeviceID As Long, Value As Double)
        hs.SetDeviceValueByRef(DeviceID, Value, True)

        'Capi.Do_Update = False
        'hs.CAPIControlHandler(Capi)
        'hs.WriteLog(GlobalVariables.ModuleName, "Set Device " + DeviceID.ToString + " to " + Command)
        Debug(4, "HSPI", "Set Device " + DeviceID.ToString + " to " + Value.ToString)

    End Sub

    Public Shared Function GetBinaryDevice(dRef)
        Dim IsON As Boolean = False
        IsON = hs.IsON(dRef)
        Return IsON
    End Function

#If PlugDLL Then
    ' These 2 functions for internal use only
    Public Property HSObj As HomeSeerAPI.IHSApplication Implements HomeSeerAPI.IPlugInAPI.HSObj
        Get
            Return hs
        End Get
        Set(value As HomeSeerAPI.IHSApplication)
            hs = value
        End Set
    End Property

    Public Property CallBackObj As HomeSeerAPI.IAppCallbackAPI Implements HomeSeerAPI.IPlugInAPI.CallBackObj
        Get
            Return callback
        End Get
        Set(value As HomeSeerAPI.IAppCallbackAPI)
            callback = value
        End Set
    End Property
#End If
End Class

