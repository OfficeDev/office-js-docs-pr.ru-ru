---
title: LaunchEvent в файле манифеста
description: Элемент LaunchEvent настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/12/2021
ms.localizationpriority: medium
ms.openlocfilehash: 23615424e194917a15b20ea4afbf7d9c5b8017e9
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59153886"
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

## <a name="see-also"></a>Дополнительные материалы

- [LaunchEvents](launchevents.md)
