Imports System.Net.NetworkInformation
Imports System.Text.RegularExpressions
Imports Microsoft.Win32

Module MainModule

  Sub Main()
    Try
      FixRegistry()
    Catch ex As Exception
      Console.WriteLine(ex.ToString)
      Console.Write("Press a key to exit")
      Console.ReadKey()
    End Try
  End Sub

  Private Sub FixRegistry()
    Dim adapters = NetworkInterface.GetAllNetworkInterfaces().ToDictionary(Function(item) item.Id, Function(item) item)

    Dim interfaces = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Services\NetBT\Parameters\Interfaces")
    Dim subKeyNames = interfaces.GetSubKeyNames

    Console.WriteLine("Network interfaces:")
    Console.WriteLine()

    Dim guids = New List(Of String)

    For index = 0 To subKeyNames.Count - 1
      Dim subKeyName = subKeyNames(index)
      Dim guid = UCase(Regex.Match(subKeyName, "{(.*?)}").Groups(0).Value)
      guids.Add(guid)
      If adapters.ContainsKey(guid) Then
        Console.WriteLine($"{index}: {guid} {adapters(guid).Description}")
      Else
        Console.WriteLine($"{index}: {guid}")
      End If
    Next
    Console.WriteLine()
    Console.Write("Please select index of interface to fix (E:exit, A:All, 1..N: single interface):")
    Dim input = Console.ReadLine

    If UCase(input) = "E" Then Return
    If UCase(input) = "A" Then
      Console.Write("Do you want to clear old values? (Y/N) ")
      Dim clear = Console.ReadLine
      If UCase(clear) = "Y" Then
        ClearOldValues()
      End If
      For index = 0 To subKeyNames.Count - 1
        FixInterface(guids(index))
      Next
    ElseIf Integer.TryParse(input, Nothing) Then
      Dim indexToFix = CInt(input)
      If indexToFix < 0 OrElse indexToFix > subKeyNames.Count - 1 Then
        Console.WriteLine()
        Console.Write($"Index {indexToFix} doesn't exist, press a key to exit")
        Console.ReadKey()
        Return
      End If
      FixInterface(guids(indexToFix))
    Else
      Console.WriteLine()
      Console.Write($"Wrong input, press a key to exit")
      Console.ReadKey()
      Return
    End If

    Console.WriteLine()
    Console.Write("Press a key to exit")
    Console.ReadKey()
  End Sub

  Private Sub ClearOldValues()
    ' Clear HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NetBT\Linkage
    Dim linkage = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Services\NetBT\Linkage", True)
    Console.WriteLine("Clearing key HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NetBT\Linkage")

    linkage.SetValue("Bind", New String() {}, RegistryValueKind.MultiString)
    Console.WriteLine("Cleared Bind")
    linkage.SetValue("Export", New String() {}, RegistryValueKind.MultiString)
    Console.WriteLine("Cleared Export")
    linkage.SetValue("Route", New String() {}, RegistryValueKind.MultiString)
    Console.WriteLine("Cleared Route")

    ' Clear HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Linkage
    linkage = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Services\LanmanServer\Linkage", True)
    Console.WriteLine("Clearing key HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Linkage")
    linkage.SetValue("Bind", New String() {}, RegistryValueKind.MultiString)
    Console.WriteLine("Cleared Bind")
    linkage.SetValue("Export", New String() {}, RegistryValueKind.MultiString)
    Console.WriteLine("Cleared Export")
    linkage.SetValue("Route", New String() {}, RegistryValueKind.MultiString)
    Console.WriteLine("Cleared Route")

    ' Clear HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Linkage
    linkage = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Linkage", True)
    Console.WriteLine("Clearing key HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Linkage")
    linkage.SetValue("Bind", New String() {}, RegistryValueKind.MultiString)
    Console.WriteLine("Cleared Bind")
    linkage.SetValue("Export", New String() {}, RegistryValueKind.MultiString)
    Console.WriteLine("Cleared Export")
    linkage.SetValue("Route", New String() {}, RegistryValueKind.MultiString)
    Console.WriteLine("Cleared Route")

  End Sub
  Private Sub FixInterface(interfaceGuid As String)

    Console.WriteLine()
    Console.WriteLine($"Fixing interface id: {interfaceGuid}")

    Dim newEntry = String.Empty

    ' Fix HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NetBT\Linkage
    Dim linkage = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Services\NetBT\Linkage", True)

    Console.WriteLine("Fixing key HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\NetBT\Linkage")

    ' Fix Bind Value
    Dim bind = CType(linkage.GetValue("Bind"), String()).ToList
    Dim bindCount = bind.Count

    newEntry = $"\Device\Tcpip_{interfaceGuid}"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\Tcpip6_{interfaceGuid}"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)

    If bindCount <> bind.Count Then
      linkage.SetValue("Bind", bind.ToArray, RegistryValueKind.MultiString)
      Console.WriteLine("Fixed Bind value")
    Else
      Console.WriteLine("Bind value already fixed")
    End If

    ' Fix Export Value
    Dim export = CType(linkage.GetValue("Export"), String()).ToList
    Dim exportCount = export.Count

    newEntry = $"\Device\NetBT_Tcpip_{interfaceGuid}"
    If Not export.Contains(newEntry) Then export.Add(newEntry)
    newEntry = $"\Device\NetBT_Tcpip6_{interfaceGuid}"
    If Not export.Contains(newEntry) Then export.Add(newEntry)

    If exportCount <> export.Count Then
      linkage.SetValue("Export", export.ToArray, RegistryValueKind.MultiString)
      Console.WriteLine("Fixed Export value")
    Else
      Console.WriteLine("Export value already fixed")
    End If

    ' Fix Route Value
    Dim route = CType(linkage.GetValue("Route"), String()).ToList
    Dim routeCount = route.Count

    newEntry = $"""Tcpip"" ""{interfaceGuid}"""
    If Not route.Contains(newEntry) Then route.Add(newEntry)
    newEntry = $"""Tcpip6"" ""{interfaceGuid}"""
    If Not route.Contains(newEntry) Then route.Add(newEntry)

    If routeCount <> route.Count Then
      linkage.SetValue("Route", route.ToArray, RegistryValueKind.MultiString)
      Console.WriteLine("Fixed Route value")
    Else
      Console.WriteLine("Route value already fixed")
    End If

    ' Fix HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Linkage

    Console.WriteLine("Fixing key HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanServer\Linkage")

    ' Fix Bind Value
    linkage = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Services\LanmanServer\Linkage", True)
    bind = CType(linkage.GetValue("Bind"), String()).ToList
    bindCount = bind.Count

    newEntry = $"\Device\NetbiosSmb"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\Tcpip_{interfaceGuid}"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\Tcpip6_{interfaceGuid}"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\NetBT_Tcpip_{interfaceGuid}"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\NetBT_Tcpip6_{interfaceGuid}"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)

    If bindCount <> bind.Count Then
      linkage.SetValue("Bind", bind.ToArray, RegistryValueKind.MultiString)
      Console.WriteLine("Fixed Bind value")
    Else
      Console.WriteLine("Bind value already fixed")
    End If

    ' Fix Export Value
    export = CType(linkage.GetValue("Export"), String()).ToList
    exportCount = export.Count

    newEntry = $"\Device\LanmanServer_NetbiosSmb"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\LanmanServer_Tcpip_{interfaceGuid}"
    If Not export.Contains(newEntry) Then export.Add(newEntry)
    newEntry = $"\Device\LanmanServer_Tcpip6_{interfaceGuid}"
    If Not export.Contains(newEntry) Then export.Add(newEntry)
    newEntry = $"\Device\LanmanServer_NetBT_Tcpip_{interfaceGuid}"
    If Not export.Contains(newEntry) Then export.Add(newEntry)
    newEntry = $"\Device\LanmanServer_NetBT_Tcpip6_{interfaceGuid}"
    If Not export.Contains(newEntry) Then export.Add(newEntry)

    If exportCount <> export.Count Then
      linkage.SetValue("Export", export.ToArray, RegistryValueKind.MultiString)
      Console.WriteLine("Fixed Export value")
    Else
      Console.WriteLine("Export value already fixed")
    End If

    ' Fix Route Value
    route = CType(linkage.GetValue("Route"), String()).ToList
    routeCount = route.Count

    newEntry = $"""NetbiosSmb"""
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"""Tcpip"" ""{interfaceGuid}"""
    If Not route.Contains(newEntry) Then route.Add(newEntry)
    newEntry = $"""Tcpip6"" ""{interfaceGuid}"""
    If Not route.Contains(newEntry) Then route.Add(newEntry)
    newEntry = $"""NetBT"" ""Tcpip"" ""{interfaceGuid}"""
    If Not route.Contains(newEntry) Then route.Add(newEntry)
    newEntry = $"""NetBT"" ""Tcpip6"" ""{interfaceGuid}"""
    If Not route.Contains(newEntry) Then route.Add(newEntry)

    If routeCount <> route.Count Then
      linkage.SetValue("Route", route.ToArray, RegistryValueKind.MultiString)
      Console.WriteLine("Fixed Route value")
    Else
      Console.WriteLine("Route value already fixed")
    End If

    ' Fix HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Linkage

    Console.WriteLine("Fixing key HKEY_LOCAL_MACHINE\SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Linkage")

    ' Fix Bind Value
    linkage = Registry.LocalMachine.OpenSubKey("SYSTEM\CurrentControlSet\Services\LanmanWorkstation\Linkage", True)
    bind = CType(linkage.GetValue("Bind"), String()).ToList
    bindCount = bind.Count

    newEntry = $"\Device\NetbiosSmb"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\Tcpip_{interfaceGuid}"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\Tcpip6_{interfaceGuid}"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\NetBT_Tcpip_{interfaceGuid}"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\NetBT_Tcpip6_{interfaceGuid}"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)

    If bindCount <> bind.Count Then
      linkage.SetValue("Bind", bind.ToArray, RegistryValueKind.MultiString)
      Console.WriteLine("Fixed Bind value")
    Else
      Console.WriteLine("Bind value already fixed")
    End If

    ' Fix Export Value
    export = CType(linkage.GetValue("Export"), String()).ToList
    exportCount = export.Count

    newEntry = $"\Device\LanmanWorkstation_NetbiosSmb"
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"\Device\LanmanWorkstation_Tcpip_{interfaceGuid}"
    If Not export.Contains(newEntry) Then export.Add(newEntry)
    newEntry = $"\Device\LanmanWorkstation_Tcpip6_{interfaceGuid}"
    If Not export.Contains(newEntry) Then export.Add(newEntry)
    newEntry = $"\Device\LanmanWorkstation_NetBT_Tcpip_{interfaceGuid}"
    If Not export.Contains(newEntry) Then export.Add(newEntry)
    newEntry = $"\Device\LanmanWorkstation_NetBT_Tcpip6_{interfaceGuid}"
    If Not export.Contains(newEntry) Then export.Add(newEntry)

    If exportCount <> export.Count Then
      linkage.SetValue("Export", export.ToArray, RegistryValueKind.MultiString)
      Console.WriteLine("Fixed Export value")
    Else
      Console.WriteLine("Export value already fixed")
    End If

    ' Fix Route Value
    route = CType(linkage.GetValue("Route"), String()).ToList
    routeCount = route.Count

    newEntry = $"""NetbiosSmb"""
    If Not bind.Contains(newEntry) Then bind.Add(newEntry)
    newEntry = $"""Tcpip"" ""{interfaceGuid}"""
    If Not route.Contains(newEntry) Then route.Add(newEntry)
    newEntry = $"""Tcpip6"" ""{interfaceGuid}"""
    If Not route.Contains(newEntry) Then route.Add(newEntry)
    newEntry = $"""NetBT"" ""Tcpip"" ""{interfaceGuid}"""
    If Not route.Contains(newEntry) Then route.Add(newEntry)
    newEntry = $"""NetBT"" ""Tcpip6"" ""{interfaceGuid}"""
    If Not route.Contains(newEntry) Then route.Add(newEntry)

    If routeCount <> route.Count Then
      linkage.SetValue("Route", route.ToArray, RegistryValueKind.MultiString)
      Console.WriteLine("Fixed Route value")
    Else
      Console.WriteLine("Route value already fixed")
    End If
  End Sub
End Module
