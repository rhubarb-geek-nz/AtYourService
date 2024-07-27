# rhubarb-geek-nz/AtYourService

Demonstration of a COM object in a service.

The class is registered when the service starts. The `CLSID` is found using the `ProgID`.

`CoCreateInstance` will start the service if not already started.

The service runs as the local system user and uses impersonation in the method handler.

```
Windows Registry Editor Version 5.00

[HKEY_LOCAL_MACHINE\SOFTWARE\Classes\RhubarbGeekNz.AtYourService\CLSID]
@="{D52EC170-7501-421A-B4C0-0F88988D323A}"

[HKEY_LOCAL_MACHINE\SOFTWARE\Classes\CLSID\{D52EC170-7501-421A-B4C0-0F88988D323A}]
"AppID"="{FCE27BC8-8959-4F8D-83C1-0C63E43677E0}"

[HKEY_LOCAL_MACHINE\SOFTWARE\Classes\AppID\{FCE27BC8-8959-4F8D-83C1-0C63E43677E0}]
"LocalService"="AtYourService"
```

[dispsvc.idl](dispsvc/dispsvc.idl) defines the dual-interface for a simple local server.

[dispsvc.c](dispsvc/dispsvc.c) implements the interface in the service.

[dispapp.cpp](dispapp/dispapp.cpp) creates an instance with [CoCreateInstance](https://learn.microsoft.com/en-us/windows/win32/api/combaseapi/nf-combaseapi-cocreateinstance) and uses it to get a message to display.

[dispnet.cs](dispnet/dispnet.cs) demonstrates using import library.

[package.ps1](package.ps1) is used to automate the building of multiple architectures.

[dispps1.ps1](dispps1/dispps1.ps1) PowerShell creating and invoking Hello World
