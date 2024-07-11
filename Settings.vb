''' A simple class for setting and getting different settings from the ini file
''' This can of course be done directly to/from the INI file, but having a custom class makes the programming simpler.
Public Class Settings

    Public Sub New()
        Me.Load()
    End Sub

    Public Sub Load()
        Me.IniDebugLevel = hs.GetINISetting("Settings", "DebugLevel", 5, INIFILE)
        Me.RawLogFile = hs.GetINISetting("Settings", "KMLogFile", "", INIFILE)
        GlobalVariables.USID = New Guid(hs.GetINISetting("Settings", "USID", Guid.Empty.ToString, INIFILE))
        GlobalVariables.Scale = hs.GetINISetting("Settings", "Scale", "Celsius", INIFILE)
        If GlobalVariables.USID = Guid.Empty Then
            Debug(4, "SETTINGS", "Missing USID Generating")
            GlobalVariables.USID = Guid.NewGuid()
            hs.SaveINISetting("Settings", "USID", GlobalVariables.USID.ToString, INIFILE)
        End If
        GetRelease()
    End Sub

    Public Sub Save()
        hs.SaveINISetting("Settings", "DebugLevel", Me.IniDebugLevel, INIFILE)
        hs.SaveINISetting("Settings", "KMLogFile", Me.RawLogFile, INIFILE)
        hs.SaveINISetting("Settings", "Scale", GlobalVariables.Scale, INIFILE)
    End Sub


    Private _IniDebugLevel As Int16
    Public Property IniDebugLevel() As Int16
        Get
            Return _IniDebugLevel
        End Get
        Set(ByVal value As Int16)
            _IniDebugLevel = value
        End Set
    End Property

    Private _RawLogFIle As String
    Public Property RawLogFile() As String
        Get
            Return _RawLogFIle
        End Get
        Set(ByVal value As String)
            _RawLogFIle = value
        End Set
    End Property

End Class
