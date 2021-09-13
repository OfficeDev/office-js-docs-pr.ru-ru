---
title: Управление состоянием и настройками Outlook надстройки
description: Узнайте, как сохранить состояние надстройки и параметры для Outlook надстройки.
ms.date: 05/17/2021
ms.localizationpriority: medium
ms.openlocfilehash: dee9d49c50df610957cb009c73451c58507adbc5
ms.sourcegitcommit: 1306faba8694dea203373972b6ff2e852429a119
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2021
ms.locfileid: "59150662"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a>Управление состоянием и настройками Outlook надстройки

> [!NOTE]
> Перед [чтением этой](../develop/persisting-add-in-state-and-settings.md) статьи просмотрите состояние и параметры сохраняющихся надстройок в разделе **Основные** концепции этой документации.

Для Outlook надстройки Office API JavaScript предоставляет объекты [RoamingSettings](/javascript/api/outlook/office.roamingsettings) и [CustomProperties](/javascript/api/outlook/office.customproperties) для сохранения состояния надстройки во всех сеансах, как описано в следующей таблице. Во всех случаях сохраненные значения параметров связаны с [Id](../reference/manifest/id.md) создавшей их надстройки.

|**Object**|**Расположение хранилища**|
|:-----|:-----|
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Почтовый ящик пользователя на сервере Exchange, на котором установлена надстройка. Поскольку эти параметры хранятся в почтовом ящике сервера пользователя, они могут "перемещаться" с пользователем и доступны надстройке, когда она запущена в контексте любого поддерживаемого клиента, доступ к почтовому ящику этого пользователя.<br/><br/> Параметры перемещения надстройки Outlook доступны только создавшей их надстройке и только в том почтовом ящике, в котором она установлена.|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Элемент сообщения, встречи, запроса на собрание для которого была запущена надстройка. Пользовательские свойства элемента надстройки Outlook доступны только для создавшей их надстройки и только в элементе, в котором они сохранены.|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Сохранение параметров в почтовом ящике пользователя для надстроек Outlook в качестве параметров перемещения

Надстройка Outlook может использовать объект [RoamingSettings](/javascript/api/outlook/office.roamingsettings) для сохранения сведений о состоянии и параметров надстройки, относящихся к почтовому ящику пользователя. Эти данные доступны только этой надстройке Outlook, запущенной от имени пользователя. Эти данные хранятся в почтовом ящике пользователя на сервере Exchange Server и становятся доступны, когда пользователь войдет в свою учетную запись и запустит надстройку Outlook.

### <a name="loading-roaming-settings"></a>Загрузка параметров перемещения

В следующем примере кода JavaScript показано, как выполняется загрузка существующих параметров перемещения.

```js
var _settings = Office.context.roamingSettings;
```

### <a name="creating-or-assigning-a-roaming-setting"></a>Создание или назначение параметра перемещения

Развивая предыдущий пример, следующая функция  `setAppSetting`, показывает, как использовать метод [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set_name__value_) для определения или обновления заданного параметра `cookie` с указанием сегодняшнего числа. Затем он позволяет заново сохранить все параметры перемещения на сервере Exchange при помощи метода [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveAsync_callback_).

```js
// Set an add-in setting.
function setAppSetting() {
    _settings.set("cookie", Date());
    _settings.saveAsync(saveMyAppSettingsCallback);
}

// Saves all roaming settings.
function saveMyAppSettingsCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

Метод **saveAsync** сохраняет параметры перемещения асинхронно и получает дополнительную функцию обратного вызова. Данный пример кода передает функцию обратного вызова `saveMyAppSettingsCallback` в метод **saveAsync**. После возврата асинхронного вызова параметр _asyncResult_ функции `saveMyAppSettingsCallback` предоставляет доступ к объекту [AsyncResult](/javascript/api/office/office.asyncresult), который можно использовать для определения успешного или неудачного выполнения операции при помощи свойства **AsyncResult.status**.

### <a name="removing-a-roaming-setting"></a>Удаление параметра перемещения

Предыдущие примеры дополняет следующая функция  `removeAppSetting`, демонстрирующая применение метода [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove_name_) для удаления параметра `cookie` и повторного сохранения всех параметров перемещения на сервере Exchange.

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a>Сохранение параметров для каждого элемента надстройки Outlook в качестве пользовательских свойств

Пользовательские свойства позволяют надстройке Outlook сохранять сведения об элементе, который она использует. Например, если в надстройке Outlook создается встреча на основе приглашения на собрание в сообщении, с помощью пользовательских свойств можно сохранить сведения о факте создания собрания. Это гарантирует, что надстройка не предложит создать встречу еще раз при повторном открытии сообщения.

Перед использованием пользовательских свойств для определенного сообщения, встречи или элемента приглашения на собрание, необходимо загрузить свойства в память путем вызова метода [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) объекта **Item**. Если какие-либо пользовательские свойства уже заданы для текущего элемента, на этом этапе они загружаются с сервера Exchange. После загрузки свойств можно использовать методы [set](/javascript/api/outlook/office.customproperties#set_name__value_) и [get](/javascript/api/outlook/office.roamingsettings) объекта **CustomProperties** для добавления, обновления и получения свойств в памяти. Чтобы сохранить любые изменения, внесенные в пользовательские свойства элемента, необходимо использовать метод [saveAsync](/javascript/api/outlook/office.customproperties#saveAsync_callback__asyncContext_) для сохранения изменений в элементе на сервере Exchange.

### <a name="custom-properties-example"></a>Пример пользовательских свойств

В следующем примере демонстрируется упрощенный набор функций для надстройки Outlook, применяющей пользовательские свойства. Этот пример можно использовать в качестве отправной точки для работы с такой надстройкой Outlook.

Надстройка Outlook, использующая эти функции, получает любые пользовательские свойства, вызывая метод **get** для переменной `_customProps`, как показано в приведенном ниже примере.

```js
var property = _customProps.get("propertyName");
```

В этом примере содержатся следующие функции.

|**Имя функции**|**Описание**|
|:-----|:-----|
| `Office.initialize`|Инициализирует надстройку и загружает пользовательские свойства текущего элемента с сервера Exchange.|
| `customPropsCallback`|Получает пользовательские свойства, возвращенные сервером Exchange, и сохраняет их для дальнейшего использования.|
| `updateProperty`|Задает или обновляет определенное свойство, а затем сохраняет изменение на сервер Exchange.|
| `removeProperty`|Удаляет определенное свойство и сохраняет факт удаления на сервере Exchange.|
| `saveCallback`|Обратный вызов метода **saveAsync** в функциях `updateProperty` и `removeProperty`.|

```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
    _mailbox = Office.context.mailbox;
    _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
    });
}

// Get the item's custom properties from the server and save for later use.
function customPropsCallback(asyncResult) {
    _customProps = asyncResult.value;
}

// Sets or updates the specified property, and then saves the change
// to the server.
function updateProperty(name, value) {
    _customProps.set(name, value);
    _customProps.saveAsync(saveCallback);
}

// Removes the specified property, and then persists the removal
// to the server.
function removeProperty(name) {
   _customProps.remove(name);
   _customProps.saveAsync(saveCallback);
}

// Callback for calls to saveAsync method.
function saveCallback(asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        // Handle the failure.
    }
}
```

### <a name="platform-behavior-in-emails"></a>Поведение платформы в сообщениях электронной почты

В следующей таблице обобщается сохраненное поведение пользовательских свойств в сообщениях электронной почты для Outlook клиентов.

|Сценарий|Windows|Web|Mac|
|---|---|---|---|
|Новое сочинение|null|null|null|
|Ответ, ответ все|null|null|null|
|Перенаправление|Загружает свойства родителей|null|null|
|Отправленный элемент из новой композиции|null|null|null|
|Отправленный элемент из ответа или ответа|null|null|null|
|Отправленный элемент из вперед|Удаляет свойства родителей, если их не сохранить|null|null|

Для обработки ситуации на Windows:

1. Проверьте существующие свойства при инициализации надстройки и храните их или очищайте по мере необходимости.
1. При настройке настраиваемого свойства включайте дополнительное свойство, чтобы указать, были ли добавлены настраиваемые свойства во время чтения сообщения или в режиме чтения надстройки. Это поможет вам различать, было ли свойство создано во время создания или унаследовано от родителя.
1. Чтобы проверить, перенаносит ли пользователь сообщение электронной почты или отвечает, можно использовать [item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getComposeTypeAsync_options__callback_) (доступно из набора требований 1.10).

## <a name="see-also"></a>Дополнительные материалы

- [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md)
- [Инициализация надстройки Office](../develop/initialize-add-in.md)
