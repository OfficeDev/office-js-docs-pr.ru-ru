---
title: Управление состоянием и настройками Outlook надстройки
description: Узнайте, как сохранить состояние надстройки и параметры для Outlook надстройки.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 8f43c7f105dc68c879f175beabcabb49715a75aa
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348505"
---
# <a name="manage-state-and-settings-for-an-outlook-add-in"></a><span data-ttu-id="0420d-103">Управление состоянием и настройками Outlook надстройки</span><span class="sxs-lookup"><span data-stu-id="0420d-103">Manage state and settings for an Outlook add-in</span></span>

> [!NOTE]
> <span data-ttu-id="0420d-104">Перед [чтением этой](../develop/persisting-add-in-state-and-settings.md) статьи просмотрите состояние и параметры сохраняющихся надстройок в разделе **Основные** концепции этой документации.</span><span class="sxs-lookup"><span data-stu-id="0420d-104">Please review [Persisting add-in state and settings](../develop/persisting-add-in-state-and-settings.md) in the **Core concepts** section of this documentation before reading this article.</span></span>

<span data-ttu-id="0420d-105">Для Outlook надстройки Office API JavaScript предоставляет объекты [RoamingSettings](/javascript/api/outlook/office.roamingsettings) и [CustomProperties](/javascript/api/outlook/office.customproperties) для сохранения состояния надстройки во всех сеансах, как описано в следующей таблице.</span><span class="sxs-lookup"><span data-stu-id="0420d-105">For Outlook add-ins, the Office JavaScript API provides [RoamingSettings](/javascript/api/outlook/office.roamingsettings) and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="0420d-106">Во всех случаях сохраненные значения параметров связаны с [Id](../reference/manifest/id.md) создавшей их надстройки.</span><span class="sxs-lookup"><span data-stu-id="0420d-106">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="0420d-107">**Object**</span><span class="sxs-lookup"><span data-stu-id="0420d-107">**Object**</span></span>|<span data-ttu-id="0420d-108">**Расположение хранилища**</span><span class="sxs-lookup"><span data-stu-id="0420d-108">**Storage location**</span></span>|
|:-----|:-----|
|[<span data-ttu-id="0420d-109">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="0420d-109">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="0420d-110">Почтовый ящик пользователя на сервере Exchange, на котором установлена надстройка.</span><span class="sxs-lookup"><span data-stu-id="0420d-110">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="0420d-111">Поскольку эти параметры хранятся в почтовом ящике сервера пользователя, они могут "перемещаться" с пользователем и доступны надстройке, когда она запущена в контексте любого поддерживаемого клиента, доступ к почтовому ящику этого пользователя.</span><span class="sxs-lookup"><span data-stu-id="0420d-111">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="0420d-112">Параметры перемещения надстройки Outlook доступны только создавшей их надстройке и только в том почтовом ящике, в котором она установлена.</span><span class="sxs-lookup"><span data-stu-id="0420d-112">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|
|[<span data-ttu-id="0420d-113">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="0420d-113">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="0420d-p103">Элемент сообщения, встречи, запроса на собрание для которого была запущена надстройка. Пользовательские свойства элемента надстройки Outlook доступны только для создавшей их надстройки и только в элементе, в котором они сохранены.</span><span class="sxs-lookup"><span data-stu-id="0420d-p103">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|

## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="0420d-116">Сохранение параметров в почтовом ящике пользователя для надстроек Outlook в качестве параметров перемещения</span><span class="sxs-lookup"><span data-stu-id="0420d-116">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>

<span data-ttu-id="0420d-117">Надстройка Outlook может использовать объект [RoamingSettings](/javascript/api/outlook/office.roamingsettings) для сохранения сведений о состоянии и параметров надстройки, относящихся к почтовому ящику пользователя.</span><span class="sxs-lookup"><span data-stu-id="0420d-117">An Outlook add-in can use the [RoamingSettings](/javascript/api/outlook/office.roamingsettings) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="0420d-118">Эти данные доступны только этой надстройке Outlook, запущенной от имени пользователя.</span><span class="sxs-lookup"><span data-stu-id="0420d-118">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="0420d-119">Эти данные хранятся в почтовом ящике пользователя на сервере Exchange Server и становятся доступны, когда пользователь войдет в свою учетную запись и запустит надстройку Outlook.</span><span class="sxs-lookup"><span data-stu-id="0420d-119">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>

### <a name="loading-roaming-settings"></a><span data-ttu-id="0420d-120">Загрузка параметров перемещения</span><span class="sxs-lookup"><span data-stu-id="0420d-120">Loading roaming settings</span></span>

<span data-ttu-id="0420d-121">В следующем примере кода JavaScript показано, как выполняется загрузка существующих параметров перемещения.</span><span class="sxs-lookup"><span data-stu-id="0420d-121">The following JavaScript code example shows how to load existing roaming settings.</span></span>

```js
var _settings = Office.context.roamingSettings;
```

### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="0420d-122">Создание или назначение параметра перемещения</span><span class="sxs-lookup"><span data-stu-id="0420d-122">Creating or assigning a roaming setting</span></span>

<span data-ttu-id="0420d-p105">Развивая предыдущий пример, следующая функция  `setAppSetting`, показывает, как использовать метод [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) для определения или обновления заданного параметра `cookie` с указанием сегодняшнего числа. Затем он позволяет заново сохранить все параметры перемещения на сервере Exchange при помощи метода [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-).</span><span class="sxs-lookup"><span data-stu-id="0420d-p105">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.roamingsettings#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](/javascript/api/outlook/office.roamingsettings#saveasync-callback-) method.</span></span>

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

<span data-ttu-id="0420d-125">Метод **saveAsync** сохраняет параметры перемещения асинхронно и получает дополнительную функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="0420d-125">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function.</span></span> <span data-ttu-id="0420d-126">Данный пример кода передает функцию обратного вызова `saveMyAppSettingsCallback` в метод **saveAsync**.</span><span class="sxs-lookup"><span data-stu-id="0420d-126">This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method.</span></span> <span data-ttu-id="0420d-127">После возврата асинхронного вызова параметр _asyncResult_ функции `saveMyAppSettingsCallback` предоставляет доступ к объекту [AsyncResult](/javascript/api/office/office.asyncresult), который можно использовать для определения успешного или неудачного выполнения операции при помощи свойства **AsyncResult.status**.</span><span class="sxs-lookup"><span data-stu-id="0420d-127">When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](/javascript/api/office/office.asyncresult) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>

### <a name="removing-a-roaming-setting"></a><span data-ttu-id="0420d-128">Удаление параметра перемещения</span><span class="sxs-lookup"><span data-stu-id="0420d-128">Removing a roaming setting</span></span>

<span data-ttu-id="0420d-129">Предыдущие примеры дополняет следующая функция  `removeAppSetting`, демонстрирующая применение метода [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) для удаления параметра `cookie` и повторного сохранения всех параметров перемещения на сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="0420d-129">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.roamingsettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>

```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```

## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="0420d-130">Сохранение параметров для каждого элемента надстройки Outlook в качестве пользовательских свойств</span><span class="sxs-lookup"><span data-stu-id="0420d-130">How to save settings per item for Outlook add-ins as custom properties</span></span>

<span data-ttu-id="0420d-p107">Пользовательские свойства позволяют надстройке Outlook сохранять сведения об элементе, который она использует. Например, если в надстройке Outlook создается встреча на основе приглашения на собрание в сообщении, с помощью пользовательских свойств можно сохранить сведения о факте создания собрания. Это гарантирует, что надстройка не предложит создать встречу еще раз при повторном открытии сообщения.</span><span class="sxs-lookup"><span data-stu-id="0420d-p107">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="0420d-p108">Перед использованием пользовательских свойств для определенного сообщения, встречи или элемента приглашения на собрание, необходимо загрузить свойства в память путем вызова метода [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) объекта **Item**. Если какие-либо пользовательские свойства уже заданы для текущего элемента, на этом этапе они загружаются с сервера Exchange. После загрузки свойств можно использовать методы [set](/javascript/api/outlook/office.customproperties#set-name--value-) и [get](/javascript/api/outlook/office.roamingsettings) объекта **CustomProperties** для добавления, обновления и получения свойств в памяти. Чтобы сохранить любые изменения, внесенные в пользовательские свойства элемента, необходимо использовать метод [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) для сохранения изменений в элементе на сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="0420d-p108">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](/javascript/api/outlook/office.mailbox) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](/javascript/api/outlook/office.customproperties#set-name--value-) and [get](/javascript/api/outlook/office.roamingsettings) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](/javascript/api/outlook/office.customproperties#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>

### <a name="custom-properties-example"></a><span data-ttu-id="0420d-138">Пример пользовательских свойств</span><span class="sxs-lookup"><span data-stu-id="0420d-138">Custom properties example</span></span>

<span data-ttu-id="0420d-p109">В следующем примере демонстрируется упрощенный набор функций для надстройки Outlook, применяющей пользовательские свойства. Этот пример можно использовать в качестве отправной точки для работы с такой надстройкой Outlook.</span><span class="sxs-lookup"><span data-stu-id="0420d-p109">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span>

<span data-ttu-id="0420d-141">Надстройка Outlook, использующая эти функции, получает любые пользовательские свойства, вызывая метод **get** для переменной `_customProps`, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="0420d-141">An Outlook add-in that uses these functions retrieves any custom properties by calling the **get** method on the `_customProps` variable, as shown in the following example.</span></span>

```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="0420d-142">В этом примере содержатся следующие функции.</span><span class="sxs-lookup"><span data-stu-id="0420d-142">This example includes the following functions.</span></span>

|<span data-ttu-id="0420d-143">**Имя функции**</span><span class="sxs-lookup"><span data-stu-id="0420d-143">**Function name**</span></span>|<span data-ttu-id="0420d-144">**Описание**</span><span class="sxs-lookup"><span data-stu-id="0420d-144">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="0420d-145">Инициализирует надстройку и загружает пользовательские свойства текущего элемента с сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="0420d-145">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="0420d-146">Получает пользовательские свойства, возвращенные сервером Exchange, и сохраняет их для дальнейшего использования.</span><span class="sxs-lookup"><span data-stu-id="0420d-146">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="0420d-147">Задает или обновляет определенное свойство, а затем сохраняет изменение на сервер Exchange.</span><span class="sxs-lookup"><span data-stu-id="0420d-147">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="0420d-148">Удаляет определенное свойство и сохраняет факт удаления на сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="0420d-148">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="0420d-149">Обратный вызов метода **saveAsync** в функциях `updateProperty` и `removeProperty`.</span><span class="sxs-lookup"><span data-stu-id="0420d-149">Callback for calls to the **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|

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

### <a name="platform-behavior-in-emails"></a><span data-ttu-id="0420d-150">Поведение платформы в сообщениях электронной почты</span><span class="sxs-lookup"><span data-stu-id="0420d-150">Platform behavior in emails</span></span>

<span data-ttu-id="0420d-151">В следующей таблице обобщается сохраненное поведение пользовательских свойств в сообщениях электронной почты для Outlook клиентов.</span><span class="sxs-lookup"><span data-stu-id="0420d-151">The following table summarizes saved custom properties behavior in emails for various Outlook clients.</span></span>

|<span data-ttu-id="0420d-152">Сценарий</span><span class="sxs-lookup"><span data-stu-id="0420d-152">Scenario</span></span>|<span data-ttu-id="0420d-153">Windows</span><span class="sxs-lookup"><span data-stu-id="0420d-153">Windows</span></span>|<span data-ttu-id="0420d-154">Web</span><span class="sxs-lookup"><span data-stu-id="0420d-154">Web</span></span>|<span data-ttu-id="0420d-155">Mac</span><span class="sxs-lookup"><span data-stu-id="0420d-155">Mac</span></span>|
|---|---|---|---|
|<span data-ttu-id="0420d-156">Новое сочинение</span><span class="sxs-lookup"><span data-stu-id="0420d-156">New compose</span></span>|<span data-ttu-id="0420d-157">null</span><span class="sxs-lookup"><span data-stu-id="0420d-157">null</span></span>|<span data-ttu-id="0420d-158">null</span><span class="sxs-lookup"><span data-stu-id="0420d-158">null</span></span>|<span data-ttu-id="0420d-159">null</span><span class="sxs-lookup"><span data-stu-id="0420d-159">null</span></span>|
|<span data-ttu-id="0420d-160">Ответ, ответ все</span><span class="sxs-lookup"><span data-stu-id="0420d-160">Reply, reply all</span></span>|<span data-ttu-id="0420d-161">null</span><span class="sxs-lookup"><span data-stu-id="0420d-161">null</span></span>|<span data-ttu-id="0420d-162">null</span><span class="sxs-lookup"><span data-stu-id="0420d-162">null</span></span>|<span data-ttu-id="0420d-163">null</span><span class="sxs-lookup"><span data-stu-id="0420d-163">null</span></span>|
|<span data-ttu-id="0420d-164">Перенаправление</span><span class="sxs-lookup"><span data-stu-id="0420d-164">Forward</span></span>|<span data-ttu-id="0420d-165">Загружает свойства родителей</span><span class="sxs-lookup"><span data-stu-id="0420d-165">Loads parent's properties</span></span>|<span data-ttu-id="0420d-166">null</span><span class="sxs-lookup"><span data-stu-id="0420d-166">null</span></span>|<span data-ttu-id="0420d-167">null</span><span class="sxs-lookup"><span data-stu-id="0420d-167">null</span></span>|
|<span data-ttu-id="0420d-168">Отправленный элемент из новой композиции</span><span class="sxs-lookup"><span data-stu-id="0420d-168">Sent item from new compose</span></span>|<span data-ttu-id="0420d-169">null</span><span class="sxs-lookup"><span data-stu-id="0420d-169">null</span></span>|<span data-ttu-id="0420d-170">null</span><span class="sxs-lookup"><span data-stu-id="0420d-170">null</span></span>|<span data-ttu-id="0420d-171">null</span><span class="sxs-lookup"><span data-stu-id="0420d-171">null</span></span>|
|<span data-ttu-id="0420d-172">Отправленный элемент из ответа или ответа</span><span class="sxs-lookup"><span data-stu-id="0420d-172">Sent item from reply or reply all</span></span>|<span data-ttu-id="0420d-173">null</span><span class="sxs-lookup"><span data-stu-id="0420d-173">null</span></span>|<span data-ttu-id="0420d-174">null</span><span class="sxs-lookup"><span data-stu-id="0420d-174">null</span></span>|<span data-ttu-id="0420d-175">null</span><span class="sxs-lookup"><span data-stu-id="0420d-175">null</span></span>|
|<span data-ttu-id="0420d-176">Отправленный элемент из вперед</span><span class="sxs-lookup"><span data-stu-id="0420d-176">Sent item from forward</span></span>|<span data-ttu-id="0420d-177">Удаляет свойства родителей, если их не сохранить</span><span class="sxs-lookup"><span data-stu-id="0420d-177">Removes parent's properties if not saved</span></span>|<span data-ttu-id="0420d-178">null</span><span class="sxs-lookup"><span data-stu-id="0420d-178">null</span></span>|<span data-ttu-id="0420d-179">null</span><span class="sxs-lookup"><span data-stu-id="0420d-179">null</span></span>|

<span data-ttu-id="0420d-180">Для обработки ситуации на Windows:</span><span class="sxs-lookup"><span data-stu-id="0420d-180">To handle the situation on Windows:</span></span>

1. <span data-ttu-id="0420d-181">Проверьте существующие свойства при инициализации надстройки и храните их или очищайте по мере необходимости.</span><span class="sxs-lookup"><span data-stu-id="0420d-181">Check for existing properties on initializing your add-in, and keep them or clear them as needed.</span></span>
1. <span data-ttu-id="0420d-182">При настройке настраиваемого свойства включайте дополнительное свойство, чтобы указать, были ли добавлены настраиваемые свойства во время чтения сообщения или в режиме чтения надстройки.</span><span class="sxs-lookup"><span data-stu-id="0420d-182">When setting custom properties, include an additional property to indicate whether the custom properties were added during message read or by Read mode of the add-in.</span></span> <span data-ttu-id="0420d-183">Это поможет вам различать, было ли свойство создано во время создания или унаследовано от родителя.</span><span class="sxs-lookup"><span data-stu-id="0420d-183">This will help you differentiate if the property was created during compose or inherited from the parent.</span></span>
1. <span data-ttu-id="0420d-184">Чтобы проверить, перенаносит ли пользователь сообщение электронной почты или отвечает, можно использовать [item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getComposeTypeAsync_options__callback_) (доступно из набора требований 1.10).</span><span class="sxs-lookup"><span data-stu-id="0420d-184">To check if the user is forwarding an email or replying, you can use [item.getComposeTypeAsync](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getComposeTypeAsync_options__callback_) (available from requirement set 1.10).</span></span>

## <a name="see-also"></a><span data-ttu-id="0420d-185">См. также</span><span class="sxs-lookup"><span data-stu-id="0420d-185">See also</span></span>

- [<span data-ttu-id="0420d-186">Persisting add-in state and settings</span><span class="sxs-lookup"><span data-stu-id="0420d-186">Persisting add-in state and settings</span></span>](../develop/persisting-add-in-state-and-settings.md)
- [<span data-ttu-id="0420d-187">Инициализация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="0420d-187">Initialize your Office Add-in</span></span>](../develop/initialize-add-in.md)
