---
title: Управление состоянием и параметрами надстройки Outlook
description: Узнайте, как сохранить состояние и параметры надстройки Outlook.
ms.date: 07/08/2022
ms.localizationpriority: medium
ms.openlocfilehash: 59349e4b23182bf53b5863430d3d847563188b08
ms.sourcegitcommit: b6a3815a1ad17f3522ca35247a3fd5d7105e174e
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/22/2022
ms.locfileid: "66958813"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a>Управление состоянием и параметрами надстройки Outlook

> [!NOTE]
> Прежде чем [прочитать эту](../develop/persisting-add-in-state-and-settings.md) статью, просмотрите сведения о состоянии  и параметрах сохранения надстройки в разделе основных понятий этой документации.

Для надстроек Outlook API JavaScript для Office предоставляет объекты [RoamingSettings](/javascript/api/outlook/office.roamingsettings) и [CustomProperties](/javascript/api/outlook/office.customproperties) для сохранения состояния надстройки в сеансах, как описано в следующей таблице. Во всех случаях сохраненные значения параметров связаны с [Id](/javascript/api/manifest/id) создавшей их надстройки.

|**Object**|**Расположение хранилища**|
|:-----|:-----|
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Почтовый ящик пользователя на сервере Exchange, на котором установлена надстройка. Так как эти параметры хранятся в почтовом ящике сервера пользователя, они могут "перемещаться" вместе с пользователем и доступны надстройке, когда она выполняется в контексте любого поддерживаемого клиента, который имеет доступ к почтовому ящику этого пользователя.<br/><br/> Параметры перемещения надстройки Outlook доступны только создавшей их надстройке и только в том почтовом ящике, в котором она установлена.|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Элемент сообщения, встречи, запроса на собрание для которого была запущена надстройка. Пользовательские свойства элемента надстройки Outlook доступны только для создавшей их надстройки и только в элементе, в котором они сохранены.|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Сохранение параметров в почтовом ящике пользователя для надстроек Outlook в качестве параметров перемещения

Надстройка Outlook может использовать объект [RoamingSettings](/javascript/api/outlook/office.roamingsettings) для сохранения сведений о состоянии и параметров надстройки, относящихся к почтовому ящику пользователя. Эти данные доступны только этой надстройке Outlook, запущенной от имени пользователя. Эти данные хранятся в почтовом ящике пользователя на сервере Exchange Server и становятся доступны, когда пользователь войдет в свою учетную запись и запустит надстройку Outlook.

### <a name="loading-roaming-settings"></a>Загрузка параметров перемещения

В следующем примере кода JavaScript показано, как выполняется загрузка существующих параметров перемещения.

```js
const _settings = Office.context.roamingSettings;
```

### <a name="creating-or-assigning-a-roaming-setting"></a>Создание или назначение параметра перемещения

Развивая предыдущий пример, следующая функция  `setAppSetting`, показывает, как использовать метод [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-set-member(1)) для определения или обновления заданного параметра `cookie` с указанием сегодняшнего числа. Затем он позволяет заново сохранить все параметры перемещения на сервере Exchange при помощи метода [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-saveasync-member(1)).

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

Предыдущие примеры дополняет следующая функция  `removeAppSetting`, демонстрирующая применение метода [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#outlook-office-roamingsettings-remove-member(1)) для удаления параметра `cookie` и повторного сохранения всех параметров перемещения на сервере Exchange.

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

Перед использованием пользовательских свойств для определенного сообщения, встречи или элемента приглашения на собрание, необходимо загрузить свойства в память путем вызова метода [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) объекта **Item**. Если какие-либо пользовательские свойства уже заданы для текущего элемента, на этом этапе они загружаются с сервера Exchange. После загрузки свойств можно использовать методы [set](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-set-member(1)) и [get](/javascript/api/outlook/office.roamingsettings) объекта **CustomProperties** для добавления, обновления и получения свойств в памяти. Чтобы сохранить любые изменения, внесенные в пользовательские свойства элемента, необходимо использовать метод [saveAsync](/javascript/api/outlook/office.customproperties#outlook-office-customproperties-saveasync-member(1)) для сохранения изменений в элементе на сервере Exchange.

### <a name="custom-properties-example"></a>Пример пользовательских свойств

В следующем примере демонстрируется упрощенный набор функций для надстройки Outlook, применяющей пользовательские свойства. Этот пример можно использовать в качестве отправной точки для работы с такой надстройкой Outlook.

Надстройка Outlook, использующая эти функции, получает любые пользовательские свойства, вызывая метод **get** для переменной `_customProps`, как показано в приведенном ниже примере.

```js
const property = _customProps.get("propertyName");
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
let _mailbox;
let _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
    // Checks for the DOM to load using the jQuery ready method.
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

В следующей таблице перечислены сохраненные пользовательские свойства в сообщениях электронной почты для различных клиентов Outlook.

|Сценарий|Windows|Web|Mac|
|---|---|---|---|
|Создание сообщения|null|null|null|
|Ответить, ответить всем|null|null|null|
|Перенаправление|Загружает свойства родительского элемента|null|null|
|Отправленный элемент из новой записи|null|null|null|
|Отправленный элемент из ответа или ответа всем|null|null|null|
|Отправленный элемент вперед|Удаляет свойства родительского элемента, если они не сохранены|null|null|

Для обработки ситуации в Windows:

1. Проверьте существующие свойства при инициализации надстройки и сохраните их или очистите по мере необходимости.
1. При настройке настраиваемых свойств включите дополнительное свойство, указывающее, были ли добавлены пользовательские свойства во время чтения сообщения или в режиме чтения надстройки. Это поможет определить, было ли свойство создано во время создания или унаследовано от родительского объекта.
1. Чтобы проверить, пересылает ли пользователь сообщение электронной почты или отвечает, можно использовать [item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#outlook-office-messagecompose-getcomposetypeasync-member(1)) (доступно из набора обязательных элементов 1.10).

## <a name="see-also"></a>Дополнительные ресурсы

- [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md)
- [Инициализация надстройки Office](../develop/initialize-add-in.md)
