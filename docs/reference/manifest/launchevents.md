---
title: LaunchEvents в файле манифеста (предварительная версия)
description: Элемент LaunchEvents настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 02/01/2021
localization_priority: Normal
ms.openlocfilehash: 9df059879018d79a61f1c900888c8d197e0b9880
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104814"
---
# <a name="launchevents-element-preview"></a>Элемент LaunchEvents (предварительная версия)

Настраивает надстройки для активации на основе поддерживаемых событий. Child of the [`<ExtensionPoint>`](extensionpoint.md) element. Дополнительные сведения см. в настройке [надстройки Outlook для активации на основе событий.](../../outlook/autolaunch.md)

**Тип надстройки:** почтовая

> [!IMPORTANT]
> Активация на основе событий в настоящее время находится [в предварительной](../../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md) версии и доступна только в Outlook в Интернете и Windows. Дополнительные сведения см. в [предварительном просмотре функции активации на основе событий.](../../outlook/autolaunch.md#how-to-preview-the-event-based-activation-feature)

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

[ExtensionPoint](extensionpoint.md) ( Почтовая надстройка **LaunchEvent)**

## <a name="child-elements"></a>Дочерние элементы

|  Элемент |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
| [LaunchEvent](launchevent.md) | Да |  Соейте поддерживаемые события с его функцией в файле JavaScript для активации надстройки. |

## <a name="see-also"></a>См. также

- [LaunchEvent](launchevent.md)
