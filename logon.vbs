On Error Resume Next

Set objSysInfo = CreateObject("ADSystemInfo")
Set objNetwork = CreateObject("Wscript.Network")

strUserPath = "LDAP://" & objSysInfo.UserName
Set objUser = GetObject(strUserPath)

colGroups = objUser.GetEx("memberOf")
For Each strGroup in colGroups
strGroupPath = "LDAP://" & strGroup
Set objGroup = GetObject(strGroupPath)
strGroupName = objGroup.CN

if strGroupName = "_adm_led" or strGroupName = "_adm_med" or strGroupName = "_ceo" then
      objNetwork.MapNetworkDrive "S:", "\\daan\data\faelles\administration"
End If

if strGroupName = "_prod_led" or strGroupName = "_adm_led" or strGroupName = "_ceo" then
      objNetwork.MapNetworkDrive "P:", "\\daan\data\faelles\leder"
End If

if strGroupName = "_prod_led" or strGroupName = "_prod_med" or strGroupName = "_ceo" then
      objNetwork.MapNetworkDrive "L:", "\\daan\data\faelles\produktion"
End If

Next

wscript.echo "Brian & Anders IT Solutions ApS"
