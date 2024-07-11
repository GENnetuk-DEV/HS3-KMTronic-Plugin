Imports System.ComponentModel
Imports System.Net
Imports System.Text

Public Class QueueProcessing
    Public Shared Sub ProcessQueue(ByVal State As Object)
        Dim DeviceRef As String = ""
        Dim Payload As String = ""
        Dim Command As String = ""
        Dim dRef As Double = 0
        Dim deQ As String = ""
        Dim CapiCommand As String = ""
        Dim Target As String = ""
        Dim a, b, c
        Dim TheKMTDevice As DeviceInfo
        Dim TheKMTUnit As UnitInfo
        Dim Map As String

        Dim dArray
        If GlobalVariables.EnableQueue Then

            Debug(1, "QPROC", "Queue is Currently " + GlobalVariables.eQUEUE.Count().ToString)
            While GlobalVariables.eQUEUE.Count() > 0
                GlobalVariables.StatsqProcessed = GlobalVariables.StatsqProcessed + 1
                deQ = GlobalVariables.eQUEUE.Dequeue
                dArray = Split(deQ, "|")
                If UBound(dArray) > 0 Then
                    Command = dArray(0).ToString
                    DeviceRef = dArray(1).ToString
                    Payload = dArray(2).ToString

                    Debug(2, "QPROC", "Command [" + Command + "] Device [" + DeviceRef + "] PayLoad [" + Payload + "]")
                    Select Case Command
                        Case "SETHS"   'Set HS Device to payload
                            If Payload = "1" Then
                                HSPI.SetDevice(CLng(DeviceRef), "ON")
                            Else
                                HSPI.SetDevice(CLng(DeviceRef), "OFF")
                            End If
                            Debug(4, "QPROC", "Updating HS " + DeviceRef + " to " + Payload.ToString)
                        Case "SETHSV"
                            Debug(4, "QPROC", "Updating HSValue " + DeviceRef + " to " + Payload.ToString)
                            Dim Temperature As Double = CDbl(Payload)
                            Select Case GlobalVariables.Scale
                                Case "Celsius" : Temperature = Temperature
                                Case "Fahrenheit" : Temperature = (Temperature * 1.8) + 32
                                Case "Rankine" : Temperature = (Temperature + 273.15) * 1.8
                                Case "Delisle" : Temperature = (Temperature * 1.5) - 100
                                Case "Newton" : Temperature = Temperature * 33 / 100
                                Case "Réaumur" : Temperature = Temperature * 4 / 5
                                Case "Rømer" : Temperature = (Temperature * 0.525) + 7.5
                                Case "Planck" : Temperature = (Temperature + 273.15) / (1.41683385 * 10 ^ 32)
                            End Select
                            Debug(4, "QPROC", "Updating HSValue " + DeviceRef + " to " + Payload.ToString + "c -> " + Temperature.ToString + " " + GlobalVariables.Scale)
                            HSPI.SetDeviceValue(CLng(DeviceRef), Temperature)
                        Case "SETUN"
                            TheKMTDevice = GlobalVariables.KMTDevices.Find(Function(p) p.DeviceID = DeviceRef)
                            TheKMTUnit = GlobalVariables.KMTUnits.Find(Function(p) p.DeviceIP = TheKMTDevice.DeviceName)
                            Map = TheKMTUnit.Map
                            Mid(Map, TheKMTDevice.DeviceTarget, 1) = Payload
                            Hardware.SetUnit(TheKMTUnit.DeviceIP, Map)
                            Debug(4, "QPROC", "Setting Device " + TheKMTDevice.DeviceName + "/" + TheKMTDevice.DeviceTarget.ToString + " to " + Payload.ToString + " [" + Map + "]")
                        Case Else
                            Debug(2, "QPROC", "No idea how to process " + Command)
                    End Select
                Else
                    Debug(1, "QPROC", "Invalid Split")
                End If

            End While
        Else
            Debug(4, "QPROC", "Queue Processing not active")
        End If

    End Sub

End Class
