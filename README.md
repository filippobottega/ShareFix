# ShareFix
Command line tool to fix NetBios and SMB registry bindings on Windows 10

Please refer to:

[Cannot connect to Windows share via local network IP address, but can by localhost](https://superuser.com/questions/1240213/cannot-connect-to-windows-share-via-local-network-ip-address-but-can-by-localho?utm_medium=organic&utm_source=google_rich_qa&utm_campaign=google_rich_qa)

[Netbios over TcpIP broken since upgrade to creators upgrade (or hyper-V feature installed)](https://answers.microsoft.com/en-us/windows/forum/windows_10-networking-winpc/netbios-over-tcpip-broken-since-upgrade-to/fd141dd9-8500-419d-b8e7-ac7255f44ec0?messageId=64e9e7d7-ba23-4e49-a4ef-901cd60f1c7b)

## ShareFix user interface

Download https://github.com/filippobottega/ShareFix/releases/latest, unzip and run ShareFix.exe.

![1524003739414](https://github.com/filippobottega/ShareFix/blob/master/ScreenShots/1524003739414.png)

The program shows the network interfaces and let you choose to:

E: exit from the program

A: rebuild bindings of all interfaces

1..N: select an interface for witch you need to rebuild the registry bindings

If you select A the program wll ask you if you want to delete the old values before adding new ones.

## ShareFix silent mode

Run 

> ShareFix.exe --silent 

to skip all user interactions.

Commands executed are:

- A: rebuild bindings of all interfaces
- Delete the old values before adding new ones