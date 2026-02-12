# Windows 11 -- Configurazione Intune CSP (Start + Taskbar Office Only)

## Obiettivo

-   Rimuovere pubblicità e sezione Consigliati
-   Allineare Start a sinistra
-   Rimuovere Task View
-   Mostrare solo app Office nel menu Start
-   Bloccare Word, Excel, Edge e Outlook nella Taskbar

------------------------------------------------------------------------

# 1) Rimuovere sezione "Consigliati"

**OMA-URI**

    ./Device/Vendor/MSFT/Policy/Config/Start/HideRecommendedSection

Data type: Integer\
Value: 1

------------------------------------------------------------------------

# 2) Disabilitare Consumer Features (pubblicità)

**OMA-URI**

    ./Device/Vendor/MSFT/Policy/Config/Experience/AllowWindowsConsumerFeatures

Data type: Integer\
Value: 0

------------------------------------------------------------------------

# 3) Allineare Start a sinistra

**OMA-URI**

    ./Device/Vendor/MSFT/Policy/Config/Start/StartLayoutAlignment

Data type: Integer\
Value: 0

------------------------------------------------------------------------

# 4) Rimuovere Task View

**OMA-URI**

    ./Device/Vendor/MSFT/Policy/Config/Taskbar/ShowTaskViewButton

Data type: Integer\
Value: 0

------------------------------------------------------------------------

# 5) Start Layout (Office Only) -- JSON

**OMA-URI**

    ./Device/Vendor/MSFT/Policy/Config/Start/StartLayout

Data type: String

``` json
{
  "pinnedList": [
    { "desktopAppId": "Microsoft.Office.WINWORD.EXE.15" },
    { "desktopAppId": "Microsoft.Office.EXCEL.EXE.15" },
    { "desktopAppId": "Microsoft.Office.POWERPNT.EXE.15" },
    { "desktopAppId": "Microsoft.Office.OUTLOOK.EXE.15" },
    { "desktopAppId": "Microsoft.Office.ONENOTE.EXE.15" }
  ]
}
```

------------------------------------------------------------------------

# 6) Taskbar Layout -- Word, Excel, Edge, Outlook

**OMA-URI**

    ./Device/Vendor/MSFT/Policy/Config/Start/StartLayout

Data type: String

``` json
{
  "pinnedList": [
    { "desktopAppId": "Microsoft.Office.WINWORD.EXE.15" },
    { "desktopAppId": "Microsoft.Office.EXCEL.EXE.15" },
    { "desktopAppId": "Microsoft.Office.OUTLOOK.EXE.15" },
    { "desktopAppId": "MSEdge" }
  ]
}
```

------------------------------------------------------------------------

# Note Importanti

-   Funziona su Windows 11 22H2+
-   Le app devono essere già installate (Microsoft 365 Apps)
-   Il layout viene applicato al primo accesso dell'utente
-   Per ambienti Enterprise è consigliato assegnare il profilo come
    Device configuration
info- 
https://learn.microsoft.com/en-us/windows/client-management/mdm/policy-csp-start#hiderecommendedsection
https://learn.microsoft.com/en-us/windows/configuration/taskbar/pinned-apps?tabs=intune&pivots=windows-11#taskbar-layout-example
