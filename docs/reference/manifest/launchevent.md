---
title: LaunchEvent в файле манифеста
description: Элемент LaunchEvent настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/12/2021
localization_priority: Normal
ms.openlocfilehash: 763ee6b0776cce0118df35dbbad3488a54fd7599cc774de240d4f1130ed021cb
ms.sourcegitcommit: 4f2c76b48d15e7d03c5c5f1f809493758fcd88ec
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/07/2021
ms.locfileid: "57089771"
---
# <a name="launchevent-element"></a>Элемент LaunchEvent

Настраивает надстройка для активации на основе поддерживаемых событий. Ребенок [`<LaunchEvents>`](launchevents.md) элемента. Дополнительные сведения см. в Outlook [надстройки](../../outlook/autolaunch.md)для активации на основе событий.

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Тип**  |  Да  | Указывает поддерживаемый тип события. Для набора поддерживаемых типов см. в Outlook надстройку для [активации на](../../outlook/autolaunch.md#supported-events)основе событий. |
|  **FunctionName**  |  Да  | Указывает имя функции JavaScript для обработки события, указанного в `Type` атрибуте. |

## <a name="see-also"></a>См. также

- [LaunchEvents](launchevents.md)
