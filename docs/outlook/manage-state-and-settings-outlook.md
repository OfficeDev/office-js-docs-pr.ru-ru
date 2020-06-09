---
title: Управление состоянием и параметрами для надстройки Outlook
description: Сведения о том, как хранить состояние и параметры надстройки для надстройки Outlook.
ms.date: 04/22/2020
localization_priority: Normal
ms.openlocfilehash: 796c7b38f8c85a5680c9b7de43297c754a0ebc1b
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609064"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a>Управление состоянием и параметрами для надстройки Outlook

> [!NOTE]
> Перед прочтением этой статьи ознакомьтесь с разделом [Сохранение состояния надстройки и параметров](../develop/persisting-add-in-state-and-settings.md) в разделе **Основные понятия** этой документации.

Для надстроек Outlook API JavaScript для Office предоставляет объекты [roamingSettings](/javascript/api/outlook/office.roamingsettings) и [CustomProperties](/javascript/api/outlook/office.customproperties) для сохранения состояния надстройки во всех сеансах, как описано в следующей таблице. Во всех случаях сохраненные значения параметров связаны с [Id](../reference/manifest/id.md) создавшей их надстройки.

|**Object**|**Расположение хранилища**|
|:-----|:-----|:-----|
|[RoamingSettings](/javascript/api/outlook/office.roamingsettings)|Почтовый ящик пользователя на сервере Exchange, на котором установлена надстройка. Поскольку параметры сохраняются на сервере почтового ящика пользователя, они могут "перемещаться" с пользователем и доступны надстройке при запуске в контексте любого поддерживаемого клиентского ведущего приложения или браузера с получением доступа к почтовому ящику нужного пользователя.<br/><br/> Параметры перемещения надстройки Outlook доступны только для создавшей их надстройки и только в почтовом ящике, в котором она установлена.|
|[CustomProperties](/javascript/api/outlook/office.customproperties)|Элемент сообщения, встречи, запроса на собрание для которого была запущена надстройка. Пользовательские свойства элемента надстройки Outlook доступны только для создавшей их надстройки и только в элементе, в котором они сохранены.|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a>Сохранение параметров в почтовом ящике пользователя для надстроек Outlook в качестве параметров перемещения

Надстройка Outlook может использовать объект [RoamingSettings](/javascript/api/outlook/office.roamingsettings) для сохранения сведений о состоянии и параметров надстройки, относящихся к почтовому ящику пользователя. Эти данные доступны только этой надстройке Outlook, запущенной от имени пользователя. Эти данные хранятся в почтовом ящике пользователя на сервере Exchange Server и становятся доступны, когда пользователь войдет в свою учетную запись и запустит надстройку Outlook.

### <a name="loading-roaming-settings"></a>Загрузка параметров перемещения

Надстройка Outlook обычно загружает параметры перемещения в обработчик событий [Office.initialize](/javascript/api/office). В следующем примере кода JavaScript показано, как выполняется загрузка существующих параметров перемещения.

```js
var _mailbox;
var _settings;

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
    // After the DOM is loaded, add-in-specific code can run.
   // Initialize instance variables to access API objects.
    _mailbox = Office.context.mailbox;
    _settings = Office.context.roamingSettings;
    });
}
```

### <a name="creating-or-assigning-a-roaming-setting"></a>Создание или назначение параметра перемещения

Развивая предыдущий пример, следующая функция  `setAppSetting`, показывает, как использовать метод [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) для определения или обновления заданного параметра `cookie` с указанием сегодняшнего числа. Затем он позволяет заново сохранить все параметры перемещения на сервере Exchange при помощи метода [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-).

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

Предыдущие примеры дополняет следующая функция  `removeAppSetting`, демонстрирующая применение метода [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) для удаления параметра `cookie` и повторного сохранения всех параметров перемещения на сервере Exchange.

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

Перед использованием пользовательских свойств для определенного сообщения, встречи или элемента приглашения на собрание, необходимо загрузить свойства в память путем вызова метода [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) объекта **Item**. Если какие-либо пользовательские свойства уже заданы для текущего элемента, на этом этапе они загружаются с сервера Exchange. После загрузки свойств можно использовать методы [set](/javascript/api/outlook/office.customproperties#set-name--value-) и [get](/javascript/api/outlook/office.roamingsettings) объекта **CustomProperties** для добавления, обновления и получения свойств в памяти. Чтобы сохранить любые изменения, внесенные в пользовательские свойства элемента, необходимо использовать метод [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) для сохранения изменений в элементе на сервере Exchange.

### <a name="custom-properties-example"></a>Пример пользовательских свойств

В следующем примере демонстрируется упрощенный набор функций для надстройки Outlook, применяющей пользовательские свойства. Этот пример можно использовать в качестве отправной точки для работы с такой надстройкой Outlook. 

Надстройка Outlook, использующая эти функции, получает любые пользовательские свойства, вызывая метод **get** для переменной `_customProps`, как показано в приведенном ниже примере.

```js
var property = _customProps.get("propertyName");
```

Этот пример включает следующие функции:

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

## <a name="see-also"></a>См. также

- [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md)
- [Инициализация надстройки Office](../develop/initialize-add-in.md)