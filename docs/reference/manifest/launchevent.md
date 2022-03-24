---
title: LaunchEvent в файле манифеста
description: Элемент LaunchEvent настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 03/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 71469693bff7213455582a3247778cabf92c2aa3
ms.sourcegitcommit: 968d637defe816449a797aefd930872229214898
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/23/2022
ms.locfileid: "63745818"
---
# <a name="launchevent-element"></a>Элемент LaunchEvent

Настраивает надстройка для активации на основе поддерживаемых событий. Ребенок элемента [`<LaunchEvents>`](launchevents.md) . Дополнительные сведения см. в Outlook [надстройки для активации на основе событий](../../outlook/autolaunch.md).

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

- [LaunchEvents](launchevents.md)

## <a name="attributes"></a>Атрибуты

|  Атрибут  |  Обязательный  |  Описание  |
|:-----|:-----|:-----|
|  **Тип**  |  Да  | Указывает поддерживаемый тип события. Для набора поддерживаемых типов см. в Outlook надстройку для [активации](../../outlook/autolaunch.md#supported-events) на основе событий. |
|  **FunctionName**  |  Да  | Указывает имя функции JavaScript для обработки события, указанного в атрибуте `Type` . |
|  **SendMode** (предварительный просмотр) |  Нет  | Используется и `OnMessageSend` события `OnAppointmentSend` . Указывает параметры, доступные пользователю, если надстройка останавливает отправление элемента или если надстройка недоступна. Если свойство **SendMode** не включено, параметр `SoftBlock` задан по умолчанию. Для доступных параметров обратитесь к [доступным вариантам SendMode](#available-sendmode-options-preview). |

## <a name="available-sendmode-options-preview"></a>Доступные параметры SendMode (предварительный просмотр)

При включив `OnMessageSend` событие или `OnAppointmentSend` событие в манифест, необходимо также задать свойство **SendMode** . Если свойство **SendMode** не включено, параметр `SoftBlock` задан по умолчанию. Ниже параметров. В зависимости от условий, которые ищет ваша надстройка, пользователь получает предупреждение, если ваша надстройка находит проблему в отправленных элементах.

| Параметр SendMode | Описание |
|---|---|
|`PromptUser`|Если элемент не соответствует условиям надстройки, пользователь может выбрать отправку в оповещении или решить проблему, а затем попытаться отправить элемент снова. Если надстройка обрабатывает элемент длительное время, пользователю будет предложена возможность прекратить запуск надстройки и выбрать **отправку в любом случае**. Если надстройка недоступна (например, существует ошибка загрузки надстройки), элемент будет отправлен.|
|`SoftBlock`|Параметр По умолчанию, если **свойство SendMode** не включено. Пользователь получает предупреждение о том, что отправляемый им элемент не соответствует условиям надстройки, и он должен решить проблему, прежде чем снова отправить элемент. Однако, если надстройка недоступна (например, существует ошибка загрузки надстройки), элемент будет отправлен.|
|`Block`|Элемент не отправляется, если возникают какие-либо из следующих ситуаций.<br>— Элемент не соответствует условиям надстройки.<br>— надстройка не может подключиться к серверу.<br>- Существует ошибка загрузки надстройки.|

## <a name="see-also"></a>См. также

- [LaunchEvents](launchevents.md)
- [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md#supported-events)
- [Использование смарт-оповещений и события OnMessageSend в Outlook надстройки](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
