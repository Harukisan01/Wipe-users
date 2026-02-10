---
title: "Personalizzazione della Barra delle Applicazioni di Windows 11 con Intune"
author: "Ayoub Sekoum-Photography"
date: "10 Febbraio 2026"
---

# Personalizzazione della Barra delle Applicazioni di Windows 11 tramite Intune

Questa guida fornisce le istruzioni dettagliate per personalizzare e bloccare la barra delle applicazioni di **Windows 1 Això11** utilizzando **Microsoft Intune** e le **Configuration Service Provider (CSP)**.

---

## 📌 Introduzione
Le policy CSP consentono di configurare e bloccare il layout della barra delle applicazioni, nascondere icone indesiderate e definire app predefinite come Word, Outlook, Excel e Edge.

---

## 🔧 Policy CSP per la Personalizzazione

### 1. Nascondere la Sezione "Raccomandati"
Per nascondere la sezione "Raccomandati" nel menu Start:

| CSP | Valore | Tipo |
|-----|--------|------|
| `./Device/Vendor/MSFT/Policy/Config/Start/HideRecommendedSection` | `1` | Integer |

---

### 2. Personalizzare il Layout della Barra delle Applicazioni
Per bloccare app specifiche (Word, Outlook, Excel, Edge), crea un file XML con il seguente contenuto:

```xml
<LayoutModificationTemplate
    xmlns="http://schemas.microsoft.com/Start/2014/LayoutModification"
    xmlns\:defaultlayout="http://schemas.microsoft.com/Start/2014/FullDefaultLayout"
    xmlns\:start="http://schemas.microsoft.com/Start/2014/StartLayout"
    xmlns\:taskbar="http://schemas.microsoft.com/Start/2014/TaskbarLayout"
    Version="1">

  <CustomTaskbarLayoutCollection>
    <defaultlayout\:TaskbarLayout>
      <taskbar\:TaskbarPinList>
        <taskbar\:UWA AppUserModelID="Microsoft.MicrosoftEdge_8wekyb3d8bbwe!MicrosoftEdge" />
        <taskbar\:DesktopApp DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\Excel.lnk" />
        <taskbar\:DesktopApp DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\Outlook.lnk" />
        <taskbar\:DesktopApp DesktopApplicationLinkPath="%APPDATA%\Microsoft\Windows\Start Menu\Programs\Word.lnk" />
      </taskbar\:TaskbarPinList>
    </defaultlayout\:TaskbarLayout>
  </CustomTaskbarLayoutCollection>
</LayoutModificationTemplate>
