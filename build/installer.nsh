; Custom NSIS installer script for Excel Data Sync Pro
; Räumt alte python-embed Reste auf bevor die neue Version installiert wird

!macro customInit
  ; Nichts beim Init
!macroend

!macro preInit
  ; Nichts beim preInit
!macroend

!macro customInstall
  ; Nach der Installation: Alte python-embed Verzeichnisse aufräumen
  ; Falls ein altes mac-arm64 Verzeichnis im Windows-Build existiert, entfernen
  ${If} ${FileExists} "$INSTDIR\resources\app.asar.unpacked\python-embed\mac-arm64\*.*"
    RMDir /r "$INSTDIR\resources\app.asar.unpacked\python-embed\mac-arm64"
  ${EndIf}
  
  ; Alte __pycache__ Verzeichnisse entfernen (können veralteten Bytecode enthalten)
  ${If} ${FileExists} "$INSTDIR\resources\app.asar.unpacked\python\__pycache__\*.*"
    RMDir /r "$INSTDIR\resources\app.asar.unpacked\python\__pycache__"
  ${EndIf}
!macroend

!macro customUnInstall
  ; Bei Deinstallation: python-embed komplett aufräumen
  ${If} ${FileExists} "$INSTDIR\resources\app.asar.unpacked\python-embed\*.*"
    RMDir /r "$INSTDIR\resources\app.asar.unpacked\python-embed"
  ${EndIf}
  ${If} ${FileExists} "$INSTDIR\resources\app.asar.unpacked\python\*.*"
    RMDir /r "$INSTDIR\resources\app.asar.unpacked\python"
  ${EndIf}
!macroend
