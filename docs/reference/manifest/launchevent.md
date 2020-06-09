---
title: Лаунчевент в файле манифеста (Предварительная версия)
description: Элемент Лаунчевент настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/15/2020
localization_priority: Normal
ms.openlocfilehash: 4874b9f4c14e3a999f41ec3fa20a15393b031ea6
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44611780"
---
# <a name="launchevent-element-preview"></a>Элемент Лаунчевент (Preview)

Настраивает надстройку для активации на основе поддерживаемых событий. Дочерний [`<LaunchEvents>`](launchevents.md) элемент. Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).

**Тип надстройки:** почтовая

> [!IMPORTANT]
> Активация на основе событий в настоящее время находится [в режиме предварительной версии](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) и доступна только в Outlook в Интернете. Дополнительные сведения см. [в статье Просмотр функции активации на основе событий](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature).

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
|  **Тип**  |  Да  | Указывает поддерживаемый тип события. Доступны типы `OnNewMessageCompose` и `OnNewAppointmentOrganizer` . |
|  **FunctionName**  |  Да  | Задает имя функции JavaScript для обработки события, указанного в `Type` атрибуте. |

## <a name="see-also"></a>См. также

- [LaunchEvents](launchevents.md)
