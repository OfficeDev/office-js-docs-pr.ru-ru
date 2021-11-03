---
title: LaunchEvent в файле манифеста
description: Элемент LaunchEvent настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 11/01/2021
ms.localizationpriority: medium
ms.openlocfilehash: a8ab75633d87284e02e9db9b1a71f7a8436f7daf
ms.sourcegitcommit: 23ce57b2702aca19054e31fcb2d2f015b4183ba1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2021
ms.locfileid: "60681711"
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
|  **SendMode** (предварительный просмотр) |  Нет  | Необходимые для `OnMessageSend` и `OnAppointmentSend` события. Указывает параметры, доступные пользователю, если надстройка останавливает отправление элемента. Для доступных параметров обратитесь к [доступным вариантам SendMode.](#available-sendmode-options-preview) |

## <a name="available-sendmode-options-preview"></a>Доступные параметры SendMode (предварительный просмотр)

При включив событие или событие в манифест, необходимо также задать `OnMessageSend` `OnAppointmentSend` свойство **SendMode.** Ниже параметров. В зависимости от условий, которые ищет ваша надстройка, пользователь получает предупреждение, если ваша надстройка находит проблему в отправленных элементах.

| Параметр SendMode | Описание |
|---|---|
|`PromptUser`|В оповещении пользователь может выбрать отправку в любом случае **или** решить проблему, а затем попытаться отправить элемент снова.|
|`SoftBlock`|Пользователь должен устранить проблему, прежде чем снова отправить элемент.|

## <a name="see-also"></a>См. также

- [LaunchEvents](launchevents.md)
- [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md#supported-events)
- [Использование смарт-оповещений и события OnMessageSend в Outlook надстройки](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
