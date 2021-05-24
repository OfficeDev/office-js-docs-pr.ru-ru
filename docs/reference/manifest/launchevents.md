---
title: LaunchEvents в файле манифеста
description: Элемент LaunchEvents настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 16d721ca6d9402d2bd5d19787707e146358044f0
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590919"
---
# <a name="launchevents-element"></a>Элемент LaunchEvents

Настраивает надстройка для активации на основе поддерживаемых событий. Ребенок [`<ExtensionPoint>`](extensionpoint.md) элемента. Дополнительные сведения см. в Outlook [надстройки](../../outlook/autolaunch.md)для активации на основе событий.

**Тип надстройки:** почтовая

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

[ExtensionPoint](extensionpoint.md) **(Надстройка для почты LaunchEvent)**

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Да |  Карта поддерживаемого события для его функции в файле JavaScript для активации надстройки. |

## <a name="see-also"></a>См. также

- [LaunchEvent](launchevent.md)
