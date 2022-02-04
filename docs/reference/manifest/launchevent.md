---
title: LaunchEvent в файле манифеста
description: Элемент LaunchEvent настраивает надстройку для активации на основе поддерживаемых событий.
ms.date: 02/02/2022
ms.localizationpriority: medium
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
|  **SendMode** (предварительный просмотр) |  Нет  | Необходимые для и `OnMessageSend` `OnAppointmentSend` события. Указывает параметры, доступные пользователю, если надстройка останавливает отправление элемента. Для доступных параметров обратитесь к [доступным вариантам SendMode](#available-sendmode-options-preview). |

## <a name="available-sendmode-options-preview"></a>Доступные параметры SendMode (предварительный просмотр)

При включив событие `OnMessageSend` или `OnAppointmentSend` событие в манифест, необходимо также задать свойство **SendMode** . Ниже параметров. В зависимости от условий, которые ищет ваша надстройка, пользователь получает предупреждение, если ваша надстройка находит проблему в отправленных элементах.

| Параметр SendMode | Description |
|---|---|
|`PromptUser`|В оповещении пользователь может выбрать отправку в любом **случае или решить** проблему, а затем попытаться отправить элемент снова.|
|`SoftBlock`|Пользователь должен устранить проблему, прежде чем снова отправить элемент.|

## <a name="see-also"></a>См. также

- [LaunchEvents](launchevents.md)
- [Настройка надстройки Outlook для активации на основе событий](../../outlook/autolaunch.md#supported-events)
- [Использование смарт-оповещений и события OnMessageSend в Outlook надстройки](../../outlook/smart-alerts-onmessagesend-walkthrough.md)
