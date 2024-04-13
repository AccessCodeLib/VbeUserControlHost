# VbeUserControlHost
.net UserControl host for VBIDE (VBE Add-ins)

COM registration of VbeUserControlHost required

    ClassGuid: 030A1F2F-4E0B-4041-A7F5-C4C0B94BAF07
    ProgId: AccLib.VbeUserControlHost
    FullClassName: AccessCodeLib.Common.VBIDETools.VbeUserControlHost

 Using the library
 
```
var vbeControl = new VbeUserControl<TypeOfUserControlToHost>(
                        AddIn,
                        "VBE Window Caption",
                        UserControlToHost.PositionGuid,
                        InstanceOfUserControlToHost);
```
            