---
title: Лаунчевентс в файле манифеста (Предварительная версия)
description: Элемент Лаунчевентс настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 05/18/2020
localization_priority: Normal
ms.openlocfilehash: 2e1ad56d405fca0f85fad500a113fba7d0448caf
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278557"
---
# <a name="launchevents-element-preview"></a>Элемент Лаунчевентс (Preview)

Настраивает надстройку для активации на основе поддерживаемых событий. Дочерний [`<ExtensionPoint>`](extensionpoint.md) элемент. Дополнительные сведения см. [в разделе Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md).

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

[ExtensionPoint](extensionpoint.md) (почтовые надстройки**лаунчевент** )

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [лаунчевент](launchevent.md) | Да |  Сопоставление поддерживаемого события с функцией в файле JavaScript для активации надстройки. |

## <a name="see-also"></a>См. также

- [лаунчевент](launchevent.md)
