---
title: LaunchEvents в файле манифеста
description: Элемент LaunchEvents настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 02/02/2022
ms.localizationpriority: medium
---

# <a name="launchevents-element"></a>Элемент LaunchEvents

Настраивает надстройка для активации на основе поддерживаемых событий. Ребенок элемента [`<ExtensionPoint>`](extensionpoint.md) . Дополнительные сведения см. в Outlook [надстройки для активации на основе событий](../../outlook/autolaunch.md).

**Тип надстройки:** почтовая

**Допустимо только в этих схемах VersionOverrides**:

- Почта 1.1

Дополнительные сведения см. [в переопределениях Версии в манифесте](../../develop/add-in-manifests.md#version-overrides-in-the-manifest).

## <a name="syntax"></a>Синтаксис

```XML
<ExtensionPoint xsi:type="LaunchEvent">
  <LaunchEvents>
    <LaunchEvent Type="OnNewMessageCompose" FunctionName="onMessageComposeHandler"/>
    <LaunchEvent Type="OnNewAppointmentOrganizer" FunctionName="onAppointmentComposeHandler"/>
  </LaunchEvents>
  <!-- Identifies the runtime to be used (also referenced by the Runtime element). -->
  <SourceLocation resid="WebViewRuntime.Url"/>
</ExtensionPoint>
```

## <a name="contained-in"></a>Содержится в

[ExtensionPoint](extensionpoint.md) (**надстройка для почты LaunchEvent** )

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Да |  Карта поддерживаемого события для его функции в файле JavaScript для активации надстройки. |

## <a name="see-also"></a>См. также

- [LaunchEvent](launchevent.md)
