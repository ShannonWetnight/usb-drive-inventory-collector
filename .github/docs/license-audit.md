# License attribution audit

Audited on 2026-10-02 against the `net8.0-windows`, self-contained `win-x64` release. The collector's MIT license and Shannon Wetnight copyright are unchanged. Selected upstream notices in LICENSE retain their original text.

The inventory comes from a Release publish with SDK 8.0.425, the resolved `project.assets.json`, the published `.deps.json`, and the files in the publish directory. Publishing with `PublishSingleFile=false` exposes the components that the release workflow otherwise bundles into its executable. No trimming is enabled. Runtime patch versions can change when CI installs a newer 8.0 SDK; repeat this audit when dependencies or release packaging change.

| Published dependency | Resolved version | License source |
| --- | --- | --- |
| DocumentFormat.OpenXml and DocumentFormat.OpenXml.Framework | 3.3.0 | [Open XML SDK MIT](https://github.com/dotnet/Open-XML-SDK/blob/v3.3.0/LICENSE) |
| System.Management, System.Text.Encoding.CodePages, System.CodeDom | 8.0.0 | Package MIT licenses and applicable runtime component notices |
| System.IO.Packaging | 8.0.1 | Package MIT license and applicable runtime component notices |
| Microsoft.NETCore.App.Runtime.win-x64 | 8.0.31 | [Runtime MIT](https://github.com/dotnet/runtime/blob/v8.0.31/LICENSE.TXT), selected [component notices](https://github.com/dotnet/runtime/blob/v8.0.31/THIRD-PARTY-NOTICES.TXT), and Windows binary terms |
| Microsoft.WindowsDesktop.App.Runtime.win-x64 | 8.0.31 | Windows Forms and WPF source licenses, selected component notices, and Windows binary terms |

smartctl is installed separately and invoked as an external program. The collector does not redistribute it. Its project attribution and GPL-2.0-or-later reference remain in LICENSE.

## Excluded notices

Exclusions require evidence beyond whether the collector calls a library. Copied algorithms in a shipped dependency still need their notices.

| Notice removed | Evidence |
| --- | --- |
| Open XML SDK: Newtonsoft.Json | [Generator model project](https://github.com/dotnet/Open-XML-SDK/blob/v3.3.0/gen/DocumentFormat.OpenXml.Generator.Models/DocumentFormat.OpenXml.Generator.Models.csproj) uses it with `PrivateAssets="all"`. The generator is a build tool; neither it nor Newtonsoft.Json is a published dependency. |
| Open XML SDK: StyleCopAnalyzers | [Directory.Build.targets](https://github.com/dotnet/Open-XML-SDK/blob/v3.3.0/Directory.Build.targets) references the analyzer for upstream builds. No analyzer binaries are published with the collector. |
| Runtime: Json.NET | References are in build tasks, tests, and Mono WebAssembly tooling. The library-tree reference is the [runtime graph generator](https://github.com/dotnet/runtime/blob/v8.0.31/src/libraries/Microsoft.NETCore.Platforms/src/GenerateRuntimeGraph.cs), a build tool. Neither that generator nor Newtonsoft.Json is published. |
| Runtime: MessagePack-CSharp, lz4net, Nerdbank.Streams, ldap4net | None appears in the resolved dependency graph or publish inventory. Searching the runtime source finds no implementation references; MessagePack appears only in JSON test resource data. |
| Runtime: Jb Evain | Attributed code is in [ILLink](https://github.com/dotnet/runtime/tree/v8.0.31/src/tools/illink/src/linker). This release is untrimmed and does not publish linker binaries. |
| Runtime: Greg Parker, libunwind | Greg Parker's `fakepoll.h` belongs to CoreCLR's Unix PAL. [CoreCLR CMake configuration](https://github.com/dotnet/runtime/blob/v8.0.31/src/coreclr/CMakeLists.txt) selects PAL/libunwind for Unix builds. These are absent from the Windows runtime. |
| Runtime: Apple header files | Attributed headers are under [System.Native/ios](https://github.com/dotnet/runtime/tree/v8.0.31/src/native/libs/System.Native/ios); they do not participate in this Windows publish. |
| Runtime: JavaScript queues | [queue.ts](https://github.com/dotnet/runtime/blob/v8.0.31/src/mono/wasm/runtime/queue.ts) is part of Mono WebAssembly. This release publishes CoreCLR for Windows, with no JavaScript runtime. |
| Windows Forms: Library of Congress | Covers `Resources/media.mpg` in the upstream [WinformsControlsTest integration-test project](https://github.com/dotnet/winforms/blob/v8.0.31/src/System.Windows.Forms/tests/IntegrationTests/WinformsControlsTest/WinformsControlsTest.csproj). Neither the test app nor video is published. |
| WPF: Json.NET | No Newtonsoft.Json implementation or dependency appears in WPF's shipped source projects, the collector's resolved graph, or the Windows Desktop publish inventory. |

## Retained runtime notices

The release includes coreclr, the BCL, compression and networking libraries, Windows Forms, WPF managed assemblies, and WPF native binaries. WPF notices remain because these binaries are distributed even though the collector does not use WPF for its interface. Windows Forms' Ookii dialog notice and WPF's zlib notice remain.

Runtime notices also cover algorithms copied into other libraries. For example, [NativeMemory.cs](https://github.com/dotnet/runtime/blob/v8.0.31/src/libraries/System.Private.CoreLib/src/System/Runtime/InteropServices/NativeMemory.cs) incorporates mimalloc overflow logic, and [SpanHelpers.Char.cs](https://github.com/dotnet/runtime/blob/v8.0.31/src/libraries/System.Private.CoreLib/src/System/SpanHelpers.Char.cs) credits Wojciech Mula's string-search work. Removing these notices based solely on missing standalone libraries would lose applicable credits. Notices without sufficient evidence for exclusion remain.

Microsoft's [Windows licensing inventory](https://github.com/dotnet/core/blob/main/license-information-windows.md) identifies the single-file runtime and several bundled native binaries as covered by the [.NET Library License](https://dotnet.microsoft.com/dotnet_library_license.htm). `D3DCompiler_47_cor3.dll` is covered by the [Windows SDK License](https://learn.microsoft.com/legal/windows-sdk/license). LICENSE and Version Information link these terms separately from the collector's MIT license.

## Repeat the inventory

From a Windows environment with the .NET 8 SDK:

```powershell
dotnet publish WindowsApp/USBDriveInventoryCollector.csproj -c Release -r win-x64 --self-contained true -p:PublishSingleFile=false -p:DebugType=none -p:DebugSymbols=false -o audit-publish
Get-ChildItem audit-publish -File | Select-Object Name
Get-Content WindowsApp/obj/project.assets.json
Get-Content audit-publish/*.deps.json
```

Compare the inventory with the package licenses and upstream source at the resolved versions before adding or removing notices. Keep copyright, permission, and disclaimer text intact for every retained notice.
