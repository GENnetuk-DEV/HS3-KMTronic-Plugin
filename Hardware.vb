Public Class Hardware
    'This is where the hardware interface is done. 
    Public Shared Sub PollUnits()


    End Sub


    ' Read the bits from the unit and update KMTUnits & Homeseer 
    Public Shared Sub ReadUnit(UnitIP As String)
        Dim a, b, c, d, e
        Dim UnitActive As Boolean
        Dim UnitType As Integer = 0

        Dim Unit As UnitInfo = GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP)
        UnitActive = Unit.Active
        UnitType = Unit.Type

        Dim SS As String = ""

        Select Case UnitType
            Case 1
                Try
                    SS = New System.Net.WebClient().DownloadString("http://" + UnitIP.ToString + "/relays.cgi")
                    a = InStr(SS, "Status:")
                    If a > 0 Then
                        b = Mid(SS, a + 7, 20)
                        b = Replace(b, " ", "")
                        Debug(3, "Hardware", "Read [" + b.ToString + "] from Unit " + UnitIP.ToString)
                        GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).Map = b
                        GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).LastSeen = Now
                        GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).Active = True
                        If UnitActive = False Then
                            Debug(5, "Hardware", "Unit " + UnitIP.ToString + " is now online")
                            AddWarning(0, "Unit " + UnitIP.ToString + " is now online")
                            EnableDisableDevices(UnitIP, True)
                        End If
                    Else
                        Debug(4, "Hardware", "The response from Unit " + UnitIP.ToString + " Was incomplete")
                        AddWarning(3, "Unit " + UnitIP.ToString + " is sending invalid data")
                        GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).Active = False
                    End If
                    GlobalVariables.UnitReads = GlobalVariables.UnitReads + 1

                    For Each TheDevice As DeviceInfo In GlobalVariables.KMTDevices
                        If TheDevice.DeviceName = UnitIP Then
                            c = CInt(Mid(b, TheDevice.DeviceTarget, 1))
                            d = HSPI.GetBinaryDevice(TheDevice.DeviceID)
                            If d = True Then e = 1 Else e = 0
                            Debug(4, "Hardware", "Checking Device " + TheDevice.DeviceString + " HS3 = " + e.ToString + " Unit = " + c.ToString)
                            GlobalVariables.KMTDevices.Find(Function(p) p.DeviceID = TheDevice.DeviceID).LastSeen = Now
                            GlobalVariables.KMTDevices.Find(Function(p) p.DeviceID = TheDevice.DeviceID).Active = True
                            If c = 1 And e = 0 Or c = 0 And e = 1 Then
                                Debug(5, "Hardware", "Need to update HS for Device " + TheDevice.DeviceString + " from " + e.ToString + " to " + c.ToString)
                                GlobalVariables.eQUEUE.Enqueue("SETHS|" + TheDevice.DeviceID.ToString + "|" + c.ToString)
                            End If
                        End If
                    Next

                Catch ex As System.Net.WebException
                    'Failed to Contact Device
                    Debug(4, "Main", ex.Message.ToString)
                    SS = ""
                    AddWarning(3, "Unit " + UnitIP.ToString + " is not responding and is now offline")
                    Debug(5, "Hardware", "Lost Contact with Unit " + UnitIP.ToString)
                    GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).Active = False
                    EnableDisableDevices(UnitIP, False)
                End Try
            Case 2
                Try
                    SS = New System.Net.WebClient().DownloadString("http://" + Unit.DeviceIP.ToString + "/status.xml")

                    GlobalVariables.UnitReads = GlobalVariables.UnitReads + 1
                    'Loop through <temp> until </temp>
                    Dim u As Integer = 0
                    Dim v As Double = 0
                    Dim vs As String = ""
                    Dim sn As Integer = 0


                    a = InStr(SS, "<temp>")
                    While a > 0
                        sn = sn + 1
                        b = InStr(a, SS, "</temp>") 'ENd of string
                        vs = Mid(SS, a + 6, b - (a + 6))
                        SS = Right(SS, Len(SS) - b)  'truncate string to first read
                        Debug(4, "Hardware", "Read Sensor " + sn.ToString + " [" + vs + "] from " + Unit.DeviceIP.ToString)
                        If IsNumeric(vs) Then
                            Dim KMTD As DeviceInfo = GlobalVariables.KMTDevices.Find(Function(p) p.DeviceName = UnitIP And p.DeviceTarget = sn And p.Type = 2)
                            If KMTD IsNot Nothing Then
                                GlobalVariables.KMTDevices.Find(Function(p) p.DeviceName = UnitIP And p.DeviceTarget = sn And p.Type = 2).LastSeen = Now
                                GlobalVariables.KMTDevices.Find(Function(p) p.DeviceName = UnitIP And p.DeviceTarget = sn And p.Type = 2).Active = True
                                GlobalVariables.eQUEUE.Enqueue("SETHSV|" + KMTD.DeviceID.ToString + "|" + vs.ToString)
                            Else
                                Debug(4, "Hardware", "Cannot find matching KMTDevice for KMTUnit " + UnitIP + " Target " + sn.ToString + " Type 2")
                            End If
                        Else
                            Debug(4, "Hardware", "Non Numeric Value for " + UnitIP + " Target " + sn.ToString + " Type 2 - Ignoring")
                        End If
                        a = InStr(SS, "<temp>")
                    End While

                    If Unit.Active = False Then
                        Debug(5, "Hardware", "Unit " + Unit.DeviceIP.ToString + " is now online")
                        AddWarning(0, "Unit " + Unit.DeviceIP.ToString + " is now online")
                        EnableDisableDevices(Unit.DeviceIP, True)
                    End If
                    GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).LastSeen = Now
                    GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).Active = True

                Catch ex As System.Net.WebException
                    'Failed to Contact Device
                    Debug(4, "Main", ex.Message.ToString)
                    SS = ""
                    AddWarning(3, "Unit " + Unit.DeviceIP.ToString + " is not responding and is now offline")
                    Debug(5, "Hardware", "Lost Contact with Unit " + Unit.DeviceIP.ToString)
                    GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = Unit.DeviceIP).Active = False
                    EnableDisableDevices(Unit.DeviceIP, False)
                End Try
        End Select

    End Sub

    ' Set the unit with the applicable bits. 
    Public Shared Sub SetUnit(UnitIP As String, Bitmap As String)
        Dim UnitActive As Boolean

        UnitActive = GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).Active


        Dim BinHex As Integer = 0
        Dim b As Integer = 1
        Dim HexStr As String = ""

        For x = 1 To 8
            If Mid(Bitmap, x, 1) = "1" Then BinHex = BinHex + b
            b = b * 2
        Next
        HexStr = BinHex.ToString("X2")
        Debug(3, "Hardware", "Bitmap [" + Bitmap + "] binhex=[" + BinHex.ToString + "] = [" + HexStr + "]")
        Dim SS As String = ""
        Try
            SS = New System.Net.WebClient().DownloadString("http://" + UnitIP.ToString + "/FFE0" + HexStr)
            GlobalVariables.UnitWrites = GlobalVariables.UnitWrites + 1
            GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).Map = Bitmap
            GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).LastSeen = Now
            EnableDisableDevices(UnitIP, True)
        Catch ex As System.Net.WebException
            ''Failed to Contact Device
            Debug(4, "Main", ex.Message.ToString)
            SS = ""
            GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = UnitIP).Active = False
            AddWarning(3, "Unit " + UnitIP.ToString + " is not responding and is now offline")
            EnableDisableDevices(UnitIP, False)
        End Try

    End Sub


    Public Shared Sub EnableDisableDevices(UnitIP, EnableDisable)
        For Each TheDevice As DeviceInfo In GlobalVariables.KMTDevices
            If TheDevice.DeviceName = UnitIP Then
                TheDevice.Active = EnableDisable
                TheDevice.LastSeen = Now
            End If
        Next
    End Sub
End Class
