'$INCLUDE:'New-WPFMessageBox.ps1.bin.bas'
'=============================================
TYPE Params
    Content AS STRING 'mandatory
    Title AS STRING 'not mandatory
    ButtonType AS STRING 'defaults to "OK"
    'Valid ButtonType values are "OK", "OK-Cancel", "Abort-Retry-Ignore", "Yes-No-Cancel", "Yes-No", "Retry-Cancel", "Cancel-TryAgain-Continue", "None" ("None" is used when defining custom buttons. This QB64 program allows for 3)
    ContentFontSize AS STRING 'defaults to "14"
    TitleFontSize AS STRING 'defaults to "14"
    BorderThickness AS STRING 'defaults to "0"
    CornerRadius AS STRING 'defaults to "8"
    ShadowDepth AS STRING 'defaults to "3"
    BlurRadius AS STRING 'defaults to "20"
    'WindowHost AS STRING
    Timeout AS STRING 'not mandatory
    'OnLoaded AS STRING 'not mandatory
    'OnClosed AS STRING 'not mandatory
    ContentBackground AS STRING 'not mandatory, defaults to "White"
    FontFamily AS STRING 'not mandatory, defaults to "Segoe UI"
    TitleFontWeight AS STRING 'not mandatory, defaults to "Normal"
    ContentFontWeight AS STRING 'not mandatory, defaults to "Normal"
    ContentTextForeground AS STRING 'not mandatory, defaults to "Black"
    TitleTextForeground AS STRING 'not mandatory, defaults to "Black"
    BorderBrush AS STRING 'not mandatory, defaults to "Black"
    TitleBackground AS STRING 'not mandatory, defaults to "White"
    ButtonTextForeground AS STRING 'not mandatory, defaults to "Black"
    'Valid colors are listed at https://docs.microsoft.com/en-us/dotnet/api/system.windows.media.colors?view=netframework-4.8
    Sound AS STRING 'not mandatory, defaults to "Windows Notify System Generic"
    'Valid sounds are "Windows Background","Windows Balloon","Windows Battery Critical","Windows Battery Low","Windows Critical Stop","Windows Default","Windows Ding","Windows Error","Windows Exclamation","Windows Feed Discovered","Windows Foreground","Windows Hardware Fail","Windows Hardware Insert","Windows Hardware Remove","Windows Information Bar","Windows Logoff Sound","Windows Logon","Windows Menu Command","Windows Message Nudge","Windows Minimize","Windows Navigation Start","Windows Notify Calendar","Windows Notify Email","Windows Notify Messaging","Windows Notify System Generic","Windows Notify","Windows Pop-up Blocked","Windows Print complete","Windows Proximity Connection","Windows Proximity Notification","Windows Recycle","Windows Restore","Windows Ringin","Windows Ringout","Windows Shutdown","Windows Startup","Windows Unlock","Windows User Account Control"
    CustomButton1 AS STRING 'not mandatory, only used if ButtonType = "None"
    CustomButton2 AS STRING 'not mandatory, only used if ButtonType = "None"
    CustomButton3 AS STRING 'not mandatory, only used if ButtonType = "None"
END TYPE
DIM SHARED Params AS Params
TYPE ButtonClicked
    OK AS INTEGER
    Cancel AS INTEGER
    Abort AS INTEGER
    Retry AS INTEGER
    Ignore AS INTEGER
    Yes AS INTEGER
    No AS INTEGER
    TryAgain AS INTEGER
    Continue AS INTEGER
END TYPE
DIM SHARED ButtonClicked AS ButtonClicked
ButtonClicked.OK = 1
ButtonClicked.Cancel = 2
ButtonClicked.Abort = 3
ButtonClicked.Retry = 4
ButtonClicked.Ignore = 5
ButtonClicked.Yes = 6
ButtonClicked.No = 7
ButtonClicked.TryAgain = 8
ButtonClicked.Continue = 9
