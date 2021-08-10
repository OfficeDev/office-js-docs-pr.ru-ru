---
title: LaunchEvents в файле манифеста
description: Элемент LaunchEvents настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: c6714c4f52bdc1ed9d7a75a42100df8d3fe046c504575295880ff614fe4a447f
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089751"
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
