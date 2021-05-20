---
title: LaunchEvent в файле манифеста (предварительный просмотр)
description: Элемент LaunchEvent настраивает надстройки для активации на основе поддерживаемых событий.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: 7283e9aba9ca57793019ffe027a7f4d6e3243aa8
ms.sourcegitcommit: 693d364616b42eea66977eef47530adabc51a40f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/19/2021
ms.locfileid: "52555313"
---
# <a name="launchevent-element-preview"></a>Элемент LaunchEvent (предварительный просмотр)

Настраивает надстройки для активации на основе поддерживаемых событий. Дитя [`<LaunchEvents>`](launchevents.md) элемента. Для получения дополнительной информации [см Outlook.](../../outlook/autolaunch.md)

**Тип надстройки:** почтовая

> [!IMPORTANT]
> Активация на основе событий в [настоящее время находится](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) в предварительном просмотре и доступна только Outlook веб-сайтах и Windows. Для получения дополнительной информации [узнайте, как просмотреть функцию активации на основе событий.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)

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
|  **Тип**  |  Да  | Определяет поддерживаемый тип события. Для набора поддерживаемых типов см. Как [просмотреть функцию активации на основе событий.](../../outlook/autolaunch.md#supported-events) |
|  **FunctionName**  |  Да  | Указывается название функции JavaScript для обработки события, указанного в `Type` атрибуте. |

## <a name="see-also"></a>См. также

- [LaunchEvents](launchevents.md)
