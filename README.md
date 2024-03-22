# TBControlPanel
twinBASIC Control Panel Applet Demo

![image](https://github.com/fafalone/TBControlPanel/assets/7834493/96c9f525-bc23-49a9-a617-e44194b28095)

This was always the intended goal of my [Property Sheet Demo project](https://github.com/fafalone/PropsheetDemo), but I wanted to learn how to use those in an easier way first, and then got sidetracked by so many other projects. But now, I've finished this project up and can now show you a working Control Panel applet made in twinBASIC!

First of all, I'm not going to cover the basics of setting up property sheets and displaying them; that's what the first project is for. This readme will only cover invoking them through a .cpl control panel applet. This project uses the same dialog resources and image resources as the property sheet demo, with all the same IDs, it's just called different. I've only tested compiling as 64bit; I'm not sure about WOW64 applets.

## Project config
You'll need to create a Standard DLL project for this, then manually set the build output path to use a .cpl extension. The resources and modPropsheet all come from out Property Sheet Demo project; they're just added in as is, the minor modifications described below.


## The basic setup: CPlApplet entry point

When Windows find a .cpl file in System32 (or one is registered in another location), it looks for an exported function named `CPlApplet`, if it finds it, it's handled as the standard applet type we're using here. This is done in twinBASIC by creating a Standard Dll project, and labeling the function with the `[DllExport]` attribute. This is the core of the applet:

```vba
[DllExport]
Public Function CPlApplet(ByVal hWnd As LongPtr, ByVal uMsg As Long, ByVal lParam1 As LongPtr, ByVal lParam2 As LongPtr) As Long

    Select Case uMsg
        Case CPL_INIT
              Init
            Return CTRUE
        
        Case CPL_GETCOUNT 'Number of applets (*not* pages on our single applet)
            Return 1
            
        Case CPL_INQUIRE
            Return OnInquire(CLng(lParam1), lParam2)
            
        Case CPL_NEWINQUIRE
            Return OnNewInquire(CLng(lParam1), lParam2)
            
        Case CPL_DBLCLK
            Return OnDoubleClick(hWnd, lParam1, lParam2)
            
        Case CPL_STARTWPARMS
            Return OnDoubleClick(hWnd, lParam1, lParam2)
            
        Case CPL_STOP
            Return S_OK
        
        Case CPL_EXIT
            If hCtx Then ReleaseActCtx(hCtx)
    End Select
End Function
```

## Supplying the name and icon

We'll look at our custom Init function later, but we always want to return `TRUE` there. For the count, we return one-- as noted, it's the number of whole applets, not the number of pages. `CPL_INQUIRE` and `CPL_NEWINQUIRE` are where it gets the info for the name and tooltip:

```vba
Private Function OnInquire(ByVal uAppletNumber As Long, pInfo As CPLINFO) As Long
      pInfo.idIcon = IDI_CPL
      pInfo.idName = IDS_TITLE
      pInfo.idInfo = IDS_INFO
      pInfo.lData = 0
      Return 0
End Function
Private Function OnNewInquire(ByVal uAppletNumber As Long, pInfo As NEWCPLINFO) As Long
    pInfo.dwSize = LenB(Of NEWCPLINFO)
    LoadStringW 0, IDS_TITLE, VarPtr(pInfo.szName(0)), 32
    LoadStringW 0, IDS_INFO, VarPtr(pInfo.szInfo(0)), 64
    pInfo.dwFlags = 0
    pInfo.dwHelpContext = 0
    pInfo.lData = 0
    pInfo.szHelpFile(0) = 0
    Return 0
End Function
```

You only need to respond to one; I've done both to show the different techniques. These specify resource IDs within your CPL file (really, just a DLL with a different name). 

We don't take any command line arguments, so DBLCLK and STARTWPARMS just go to the same place, where we finally show the applet when clicked in the Control Panel. Here I ran into an issue- at first it didn't work. I realized, while the Control Panel knows to look inside for the info we provided above, the API calls we make wouldn't--- previously we used `GetModuleHandle()`, but that points to the resources in the hosting exe-- the Control Panel, not our dll. So for the property sheet APIs to access our resources, we have to load a reference to our CPL; there's a module-level variable we set as `hMod = LoadLibrary("TBCtlPanelDemo.cpl")`. Passing that to our existing ShowPropsheet function got the paes to show, but then I ran into a more serious issue... the main image was rendered in all the wrong colors, other images didn't work, and the 2nd page couldn't be displayed at all; it vanished and caused graphical glitching. Based on the format of the images and the fact the 2nd page used a comctl6-only control, I figured out that was the culprit-- visual styles weren't applied, so it's clear the Control Panel wasn't manifested to enabled ComCtl6 for us. Having a manifest in the CPL isn't enough, because it matters only what the parent exe had. 

It was very fortunate just the other day I had read about a potential solution to this... about how you use activation contexts to apply ComCtl6 manifests per-control. Since some other CPLs had visual styles, I figured this was a good approach. So the `Init` function became:

```vba
Private Function Init() As Long
    hMod = LoadLibrary("TBCtlPanelDemo.cpl")
    szSys = Environ$("WINDIR") & "\System32"
    Dim ctx As ACTCTX
    ctx.cbSize = LenB(Of ACTCTX)
    ctx.dwFlags = ACTCTX_FLAG_RESOURCE_NAME_VALID Or ACTCTX_FLAG_ASSEMBLY_DIRECTORY_VALID Or ACTCTX_FLAG_HMODULE_VALID
    ctx.hModule = hMod
    ctx.lpResourceName = 1
    ctx.lpAssemblyDirectory = StrPtr(szSys)
    hCtx = CreateActCtx(ctx)
End Function
```

Note that you wouldn't hard-core the names in a regular app. The APIs here dynamically load a manifest so it can be activated and deactivated at will-- note that here too we need to manually specify the CPL's module handle so it loads our manifest, not the parent exe which doesn't have one. When it's active, if it specifies ComCtl6, all calls will use it. The manifest stored in our CPL is "#1", so 1 is pointing to it. `hCtx` is module level, so that we can activate it around our property sheet call:

```vba
Private Function OnDoubleClick(hWnd As LongPtr, lParam1 As LongPtr, lParam2 As LongPtr) As Long
    Dim lCookie As LongPtr
    If hCtx Then
        ActivateActCtx(hCtx, lCookie)
    End If
    ShowPropsheet hWnd, hMod, pdtModal
    DeactivateActCtx(0, lCookie)
End Function
```

The only other minor issue was DPI awareness... you don't get to control whether it's enabled, so need to support it. So the font sizes are adjusted.

The end result of clicking the icon we saw at the top is just the property sheets we saw in the initial project, of course modified to note they're a Control Panel Applet now!

![image](https://github.com/fafalone/TBControlPanel/assets/7834493/3d4dd582-769b-476a-a30a-468b24f3f775)

TADA!

