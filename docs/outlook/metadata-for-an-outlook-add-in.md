---
title: Просмотр и изменение метаданных элемента в надстройке Outlook
description: Управление пользовательскими данными в надстройке Outlook с помощью параметров перемещения или настраиваемых свойств.
ms.date: 10/31/2019
localization_priority: Normal
ms.openlocfilehash: 3bf19f56b11b524ea2ee722e2997465bbd36d55c
ms.sourcegitcommit: 5d29801180f6939ec10efb778d2311be67d8b9f1
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/27/2020
ms.locfileid: "42324935"
---
# <a name="get-and-set-add-in-metadata-for-an-outlook-add-in"></a><span data-ttu-id="2f985-103">Просмотр и изменение метаданных для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="2f985-103">Get and set add-in metadata for an Outlook add-in</span></span>

<span data-ttu-id="2f985-104">Для управления пользовательскими данными в настройке Outlook можно использовать следующее:</span><span class="sxs-lookup"><span data-stu-id="2f985-104">You can manage custom data in your Outlook add-in by using either of the following:</span></span>

- <span data-ttu-id="2f985-105">параметры перемещения, которые управляют пользовательскими данными для почтового ящика пользователя;</span><span class="sxs-lookup"><span data-stu-id="2f985-105">Roaming settings, which manage custom data for a user's mailbox.</span></span>
- <span data-ttu-id="2f985-106">настраиваемые свойства, которые управляют пользовательскими данными для элемента в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="2f985-106">Custom properties, which manage custom data for an item in a user's mailbox.</span></span>

<span data-ttu-id="2f985-p101">Оба этих способа предоставляют доступ к пользовательским данным, доступным только надстройке Outlook, но каждый метод хранит данные отдельно от остальных. Другими словами, данные, хранящиеся с помощью параметров перемещения, недоступны настраиваемым свойствам и наоборот. Данные хранятся на сервере этого почтового ящика и доступны в последующих сеансах Outlook на всех поддерживаемых надстройкой форм-факторах.</span><span class="sxs-lookup"><span data-stu-id="2f985-p101">Both of these give access to custom data that is only accessible by your Outlook add-in, but each method stores the data separately from the other. That is, the data stored through roaming settings is not accessible by custom properties, and vice versa. The data is stored on the server for that mailbox, and is accessible in subsequent Outlook sessions on all the form factors that the add-in supports.</span></span>

## <a name="custom-data-per-mailbox-roaming-settings"></a><span data-ttu-id="2f985-110">Пользовательские данные на один почтовый ящик: параметры перемещения</span><span class="sxs-lookup"><span data-stu-id="2f985-110">Custom data per mailbox: roaming settings</span></span>

<span data-ttu-id="2f985-p102">Вы можете указать данные, специфичные для пользователя почтового ящика Exchange, с помощью объекта [RoamingSettings](/javascript/api/outlook/office.RoamingSettings). Примерами таких данных являются личные данные и предпочтения пользователя. Ваша почтовая надстройка может получить доступ к параметрам перемещения, когда перемещение происходит на любом из устройств, предназначенных для работы (настольный ПК, планшет или смартфон).</span><span class="sxs-lookup"><span data-stu-id="2f985-p102">You can specify data specific to a user's Exchange mailbox using the [RoamingSettings](/javascript/api/outlook/office.RoamingSettings) object. Examples of such data include the user's personal data and preferences. Your mail add-in can access roaming settings when it roams on any device it's designed to run on (desktop, tablet, or smartphone).</span></span>

<span data-ttu-id="2f985-p103">Изменения этих данных хранятся в памяти текущего сеанса Outlook. После изменения все параметры перемещения следует сохранить, чтобы они были доступны, когда пользователь откроет надстройку на том же или другом поддерживаемом устройстве в следующий раз.</span><span class="sxs-lookup"><span data-stu-id="2f985-p103">Changes to this data are stored on an in-memory copy of those settings for the current Outlook session. You should explicitly save all the roaming settings after updating them so that they will be available the next time the user opens your add-in, on the same or any other supported device.</span></span>


### <a name="roaming-settings-format"></a><span data-ttu-id="2f985-116">Формат параметров перемещения</span><span class="sxs-lookup"><span data-stu-id="2f985-116">Roaming settings format</span></span>

<span data-ttu-id="2f985-117">Данные в объекте **RoamingSettings** хранятся в виде сериализованной строки нотации объектов JavaScript (JSON).</span><span class="sxs-lookup"><span data-stu-id="2f985-117">The data in a **RoamingSettings** object is stored as a serialized JavaScript Object Notation (JSON) string.</span></span> 

<span data-ttu-id="2f985-118">Ниже приведен пример структуры для трех определенных параметров перемещения с именами `add-in_setting_name_0`, `add-in_setting_name_1`, и `add-in_setting_name_2`.</span><span class="sxs-lookup"><span data-stu-id="2f985-118">The following is an example of the structure, assuming there are three defined roaming settings named `add-in_setting_name_0`,  `add-in_setting_name_1`, and  `add-in_setting_name_2`.</span></span>


```json
{
  "add-in_setting_name_0": "add-in_setting_value_0",
  "add-in_setting_name_1": "add-in_setting_value_1",
  "add-in_setting_name_2": "add-in_setting_value_2"
}
```


### <a name="loading-roaming-settings"></a><span data-ttu-id="2f985-119">Загрузка параметров перемещения</span><span class="sxs-lookup"><span data-stu-id="2f985-119">Loading roaming settings</span></span>

<span data-ttu-id="2f985-120">Надстройка почты обычно загружает параметры перемещения в обработчик событий [Office.initialize](/javascript/api/office#office-initialize-reason-).</span><span class="sxs-lookup"><span data-stu-id="2f985-120">A mail add-in typically loads roaming settings in the [Office.initialize](/javascript/api/office#office-initialize-reason-) event handler.</span></span> <span data-ttu-id="2f985-121">В следующем примере кода JavaScript показано, как выполняется загрузка существующих параметров перемещения и получение значений 2 параметров **customerName** и **customerBalance**:</span><span class="sxs-lookup"><span data-stu-id="2f985-121">The following JavaScript code example shows how to load existing roaming settings and get the values of 2 settings, **customerName** and **customerBalance**:</span></span>


```js
var _mailbox;
var _settings;
var _customerName;
var _customerBalance;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  // Initialize instance variables to access API objects.
  _mailbox = Office.context.mailbox;
  _settings = Office.context.roamingSettings;
  _customerName = _settings.get("customerName");
  _customerBalance = _settings.get("customerBalance");
}

```


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="2f985-122">Создание или назначение параметра перемещения</span><span class="sxs-lookup"><span data-stu-id="2f985-122">Creating or assigning a roaming setting</span></span>

<span data-ttu-id="2f985-123">Развивая предыдущий пример, следующая функция JavaScript `setAddInSetting` показывает, как использовать метод [RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) для определения заданного параметра `cookie` с указанием сегодняшнего числа, и как сохраненить данных с помощью метода [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-), чтобы сохранить все параметры перемещения на сервере.</span><span class="sxs-lookup"><span data-stu-id="2f985-123">Continuing with the preceding example, the following JavaScript function,  `setAddInSetting`, shows how to use the [RoamingSettings.set](/javascript/api/outlook/office.RoamingSettings) method to set a setting named `cookie` with today's date, and persist the data by using the [RoamingSettings.saveAsync](/javascript/api/outlook/office.RoamingSettings#saveasync-callback-) method to save all the roaming settings back to the server.</span></span>

<span data-ttu-id="2f985-124">`set` Метод создает этот параметр, если он еще не существует, и присваивает этому параметру указанное значение.</span><span class="sxs-lookup"><span data-stu-id="2f985-124">The `set` method creates the setting if the setting does not already exist, and assigns the setting to the specified value.</span></span> <span data-ttu-id="2f985-125">`saveAsync` Метод сохраняет параметры перемещения асинхронно.</span><span class="sxs-lookup"><span data-stu-id="2f985-125">The `saveAsync` method saves roaming settings asynchronously.</span></span> <span data-ttu-id="2f985-126">Этот пример кода передает метод обратного вызова `saveMyAddInSettingsCallback`, `saveAsync` при завершении `saveMyAddInSettingsCallback` асинхронного вызова, вызывается с помощью одного параметра _asyncResult_.</span><span class="sxs-lookup"><span data-stu-id="2f985-126">This code sample passes a callback method, `saveMyAddInSettingsCallback`, to `saveAsync` When the asynchronous call finishes,  `saveMyAddInSettingsCallback` is called by using one parameter, _asyncResult_.</span></span> <span data-ttu-id="2f985-127">Этот параметр является объектом [AsyncResult](/javascript/api/office/office.asyncresult), который содержит результат и все сведения об асинхронном вызове.</span><span class="sxs-lookup"><span data-stu-id="2f985-127">This parameter is an [AsyncResult](/javascript/api/office/office.asyncresult) object that contains the result of and any details about the asynchronous call.</span></span> <span data-ttu-id="2f985-128">Необязательный параметр _userContext_ можно использовать для передачи сведений о состоянии из асинхронного вызова в функцию обратного звонка.</span><span class="sxs-lookup"><span data-stu-id="2f985-128">You can use the optional _userContext_ parameter to pass any state information from the asynchronous call to the callback function.</span></span>

```js
// Set a roaming setting.
function setAddInSetting() {
  _settings.set("cookie", Date());
  // Save roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}

// Callback method after saving custom roaming settings.
function saveMyAddInSettingsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="2f985-129">Удаление параметра перемещения</span><span class="sxs-lookup"><span data-stu-id="2f985-129">Removing a roaming setting</span></span>

<span data-ttu-id="2f985-130">Кроме того, в расширениях предыдущих примеров следующая функция JavaScript —  `removeAddInSetting` — показывает, как метод [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove-name-) используется для удаления параметра `cookie` и сохранения всех параметров перемещения обратно в Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="2f985-130">Also extending the preceding examples, the following JavaScript function,  `removeAddInSetting`, shows how to use the [RoamingSettings.remove](/javascript/api/outlook/office.RoamingSettings#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


```js
// Remove an add-in setting.
function removeAddInSetting()
{
  _settings.remove("cookie");
  // Save changes to the roaming settings for the mailbox
  // to the server so that they will be available
  // in the next session.
  _settings.saveAsync(saveMyAddInSettingsCallback);
}
```


## <a name="custom-data-per-item-in-a-mailbox-custom-properties"></a><span data-ttu-id="2f985-131">Пользовательские данные для каждого элемента в почтовом ящике: пользовательские свойства</span><span class="sxs-lookup"><span data-stu-id="2f985-131">Custom data per item in a mailbox: custom properties</span></span>

<span data-ttu-id="2f985-p106">Вы также можете указать данные, характерные для элемента в почтовом ящике пользователя, используя объект [CustomProperties](/javascript/api/outlook/office.CustomProperties). Например, ваша почтовая надстройка могла бы категоризировать некоторые сообщения и отмечать категорию с помощью настраиваемого свойства `messageCategory`. Либо, если ваша почтовая надстройка создает встречи из сообщений с предложениями о собрании, вы можете использовать настраиваемое свойство, чтобы отслеживать каждую из этих встреч. Это гарантирует, что если пользователь вновь откроет сообщение, ваша почтовая надстройка не станет во второй раз предлагать создать встречу.</span><span class="sxs-lookup"><span data-stu-id="2f985-p106">You can specify data specific to an item in the user's mailbox using the [CustomProperties](/javascript/api/outlook/office.CustomProperties) object. For example, your mail add-in could categorize certain messages and note the category using a custom property `messageCategory`. Or, if your mail add-in creates appointments from meeting suggestions in a message, you can use a custom property to track each of these appointments. This ensures that if the user opens the message again, your mail add-in doesn't offer to create the appointment a second time.</span></span>

<span data-ttu-id="2f985-p107">Аналогично параметрам перемещения, изменения настраиваемых свойств хранятся в копии контейнера свойств для текущего сеанса Outlook. Чтобы эти настраиваемые свойства были доступны при следующем сеансе, используйте [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-).</span><span class="sxs-lookup"><span data-stu-id="2f985-p107">Similar to roaming settings, changes to custom properties are stored on in-memory copies of the properties for the current Outlook session. To make sure these custom properties will be available in the next session, use [CustomProperties.saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-).</span></span>

<span data-ttu-id="2f985-138">Эти специальные свойства, относящиеся к определенному элементу, доступны только с помощью `CustomProperties` объекта.</span><span class="sxs-lookup"><span data-stu-id="2f985-138">These add-in-specific, item-specific custom properties can only be accessed by using the `CustomProperties` object.</span></span> <span data-ttu-id="2f985-139">Эти свойства отличаются от настраиваемых, основанных на MAPI [UserProperties](/office/vba/api/Outlook.UserProperties) в объектной модели Outlook и расширенных свойств в веб-службах Exchange (EWS).</span><span class="sxs-lookup"><span data-stu-id="2f985-139">These properties are different from the custom, MAPI-based [UserProperties](/office/vba/api/Outlook.UserProperties) in the Outlook object model, and extended properties in Exchange Web Services (EWS).</span></span> <span data-ttu-id="2f985-140">Вы не можете напрямую `CustomProperties` получить доступ с помощью объектной модели Outlook, EWS или REST.</span><span class="sxs-lookup"><span data-stu-id="2f985-140">You cannot directly access `CustomProperties` by using the Outlook object model, EWS, or REST.</span></span> <span data-ttu-id="2f985-141">Чтобы узнать, как получить `CustomProperties` доступ с помощью EWS или REST, ознакомьтесь с разделом [Получение настраиваемых свойств с помощью EWS или REST](#get-custom-properties-using-ews-or-rest).</span><span class="sxs-lookup"><span data-stu-id="2f985-141">To learn how to access `CustomProperties` using EWS or REST, see the section [Get custom properties using EWS or REST](#get-custom-properties-using-ews-or-rest).</span></span>

### <a name="using-custom-properties"></a><span data-ttu-id="2f985-142">Использование настраиваемых свойств</span><span class="sxs-lookup"><span data-stu-id="2f985-142">Using custom properties</span></span>

<span data-ttu-id="2f985-143">Перед использованием настраиваемых свойств необходимо загрузить их, вызвав метод [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods).</span><span class="sxs-lookup"><span data-stu-id="2f985-143">Before you can use custom properties, you must load them by calling the [loadCustomPropertiesAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.item.md#methods) method.</span></span> <span data-ttu-id="2f985-144">После создания контейнера свойств можно использовать методы [set](/javascript/api/outlook/office.CustomProperties#set-name--value-) и [get](/javascript/api/outlook/office.CustomProperties) для добавления и извлечения настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="2f985-144">After you have created the property bag, you can use the [set](/javascript/api/outlook/office.CustomProperties#set-name--value-) and [get](/javascript/api/outlook/office.CustomProperties) methods to add and retrieve custom properties.</span></span> <span data-ttu-id="2f985-145">Чтобы сохранить любые изменения, внесенные в контейнер свойств, необходимо использовать метод [saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-).</span><span class="sxs-lookup"><span data-stu-id="2f985-145">You must use the [saveAsync](/javascript/api/outlook/office.CustomProperties#saveasync-callback--asynccontext-) method to save any changes that you make to the property bag.</span></span>


 > [!NOTE]
 > <span data-ttu-id="2f985-146">Так как Outlook для Mac не кэширует настраиваемые свойства, в случае перебоев в работе сети пользователя почтовые надстройки в Outlook для Mac не смогут получить доступ к их настраиваемым свойствам.</span><span class="sxs-lookup"><span data-stu-id="2f985-146">Because Outlook on Mac doesn't cache custom properties, if the user's network goes down, mail add-ins in Outlook on Mac would not be able to access their custom properties.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="2f985-147">Пример пользовательских свойств</span><span class="sxs-lookup"><span data-stu-id="2f985-147">Custom properties example</span></span>


<span data-ttu-id="2f985-p110">Следующий пример показывает простой набор методов для надстройки Outlook, использующей настраиваемые свойства. Этот пример можно использовать в качестве отправной точки для создания надстройки, использующей настраиваемые свойства.</span><span class="sxs-lookup"><span data-stu-id="2f985-p110">The following example shows a simplified set of methods for an Outlook add-in that uses custom properties. You can use this example as a starting point for your add-in that uses custom properties.</span></span>

<span data-ttu-id="2f985-150">Этот пример содержит следующие методы:</span><span class="sxs-lookup"><span data-stu-id="2f985-150">This example includes the following methods:</span></span>


- <span data-ttu-id="2f985-151">[Office.initialize](/javascript/api/office#office-initialize-reason-): инициализирует надстройку и загружает контейнер настраиваемых свойств с сервера Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="2f985-151">[Office.initialize](/javascript/api/office#office-initialize-reason-) -- Initializes the add-in and loads the custom property bag from the Exchange server.</span></span>

- <span data-ttu-id="2f985-152">**customPropsCallback**: получает контейнер настраиваемых свойств, возвращенный с сервера, и сохраняет его для дальнейшего использования.</span><span class="sxs-lookup"><span data-stu-id="2f985-152">**customPropsCallback** -- Gets the custom property bag that is returned from the server and saves it for later use.</span></span>

- <span data-ttu-id="2f985-153">**updateProperty**: задает или обновляет определенное свойство, а затем сохраняет изменения на сервере.</span><span class="sxs-lookup"><span data-stu-id="2f985-153">**updateProperty** -- Sets or updates a specific property, and then saves the change to the server.</span></span>

- <span data-ttu-id="2f985-154">**removeProperty**: удаляет определенное свойство из контейнера свойств, а затем сохраняет удаление на сервере.</span><span class="sxs-lookup"><span data-stu-id="2f985-154">**removeProperty** -- Removes a specific property from the property bag, and then saves the removal to the server.</span></span>


```js
var _mailbox;
var _customProps;

// The initialize function is required for all add-ins.
Office.initialize = function () {
  _mailbox = Office.context.mailbox;
  _mailbox.item.loadCustomPropertiesAsync(customPropsCallback);
}

// Callback function from loading custom properties.
function customPropsCallback(asyncResult) {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
  else {
    // Successfully loaded custom properties,
    // can get them from the asyncResult argument.
    _customProps = asyncResult.value;
  }
}

// Get individual custom property.
function getProperty() {
  var myProp = _customProps.get("myProp");
}

// Set individual custom property.
function updateProperty(name, value) {
  _customProps.set(name, value);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Remove a custom property.
function removeProperty(name) {
  _customProps.remove(name);
  // Save all custom properties to server.
  _customProps.saveAsync(saveCallback);
}

// Callback function from saving custom properties.
function saveCallback() {
  if (asyncResult.status == Office.AsyncResultStatus.Failed) {
    // Handle the failure.
  }
}
```

### <a name="get-custom-properties-using-ews-or-rest"></a><span data-ttu-id="2f985-155">Просмотр настраиваемых свойств с помощью EWS или REST</span><span class="sxs-lookup"><span data-stu-id="2f985-155">Get custom properties using EWS or REST</span></span>

<span data-ttu-id="2f985-156">Чтобы получить объект **CustomProperties** с помощью EWS или REST, необходимо сначала определить имя его расширенного свойства, основанного на интерфейсе MAPI.</span><span class="sxs-lookup"><span data-stu-id="2f985-156">To get **CustomProperties** using EWS or REST, you should first determine the name of its MAPI-based extended property.</span></span> <span data-ttu-id="2f985-157">Затем можно получить это свойство способом, аналогичным используемому при получении любого расширенного свойства, основанного на интерфейсе MAPI.</span><span class="sxs-lookup"><span data-stu-id="2f985-157">You can then get that property in the same way you would get any MAPI-based extended property.</span></span>

#### <a name="how-custom-properties-are-stored-on-an-item"></a><span data-ttu-id="2f985-158">Способ хранения настраиваемых свойств в элементе</span><span class="sxs-lookup"><span data-stu-id="2f985-158">How custom properties are stored on an item</span></span>

<span data-ttu-id="2f985-159">Настраиваемые свойства, присвоенные надстройкой, отличаются от обычных свойств, основанных на интерфейсе MAPI.</span><span class="sxs-lookup"><span data-stu-id="2f985-159">Custom properties set by an add-in are not equivalent to normal MAPI-based properties.</span></span> <span data-ttu-id="2f985-160">API надстроек `CustomProperties` сериализуются всю надстройку как полезные данные JSON, а затем сохраняют их в отдельном расширенном свойстве на основе MAPI с именем `cecp-<app-guid>` (`<app-guid>` — идентификатором надстройки), а GUID набора свойств. `{00020329-0000-0000-C000-000000000046}`</span><span class="sxs-lookup"><span data-stu-id="2f985-160">Add-in APIs serialize all your add-in's `CustomProperties` as a JSON payload and then save them in a single MAPI-based extended property whose name is `cecp-<app-guid>` (`<app-guid>` is your add-in's ID) and property set GUID is `{00020329-0000-0000-C000-000000000046}`.</span></span> <span data-ttu-id="2f985-161">(Дополнительные сведения об этом объекте см. в статье [MS-OXCEXT 2.2.5 Настраиваемые свойства почтового приложения](https://msdn.microsoft.com/library/hh968549(v=exchg.80).aspx)). Затем можно использовать EWS или REST, чтобы получить это свойство, основанное на интерфейсе MAPI.</span><span class="sxs-lookup"><span data-stu-id="2f985-161">(For more information about this object, see [MS-OXCEXT 2.2.5 Mail App Custom Properties](https://msdn.microsoft.com/library/hh968549(v=exchg.80).aspx).) You can then use EWS or REST to get this MAPI-based property.</span></span>

#### <a name="get-custom-properties-using-ews"></a><span data-ttu-id="2f985-162">Просмотр настраиваемых свойств с помощью EWS</span><span class="sxs-lookup"><span data-stu-id="2f985-162">Get custom properties using EWS</span></span>

<span data-ttu-id="2f985-163">Почтовая надстройка может получить расширенное `CustomProperties` свойство на основе MAPI с помощью операции [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) .</span><span class="sxs-lookup"><span data-stu-id="2f985-163">Your mail add-in can get the `CustomProperties` MAPI-based extended property by using the EWS [GetItem](/exchange/client-developer/web-service-reference/getitem-operation) operation.</span></span> <span data-ttu-id="2f985-164">Доступ `GetItem` на стороне сервера с помощью маркера обратного вызова или на стороне клиента с помощью метода [Mailbox. makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) .</span><span class="sxs-lookup"><span data-stu-id="2f985-164">Access `GetItem` on the server side by using a callback token, or on the client side by using the [mailbox.makeEwsRequestAsync](../reference/objectmodel/preview-requirement-set/office.context.mailbox.md#methods) method.</span></span> <span data-ttu-id="2f985-165">В `GetItem` запросе укажите свойство на основе `CustomProperties` MAPI в наборе свойств, используя сведения, приведенные в предыдущем разделе, [как пользовательские свойства хранятся в элементе](#how-custom-properties-are-stored-on-an-item).</span><span class="sxs-lookup"><span data-stu-id="2f985-165">In the `GetItem` request, specify the `CustomProperties` MAPI-based property in its property set using the details provided in the preceding section [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).</span></span>

<span data-ttu-id="2f985-166">В приведенном ниже примере показано, как получить элемент и его настраиваемые свойства.</span><span class="sxs-lookup"><span data-stu-id="2f985-166">The following example shows how to get an item and its custom properties.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2f985-167">В приведенном ниже примере замените `<app-guid>` идентификатором своей надстройки.</span><span class="sxs-lookup"><span data-stu-id="2f985-167">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

```typescript
let request_str =
    '<?xml version="1.0" encoding="utf-8"?>' +
    '<soap:Envelope xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"' +
                   'xmlns:m="http://schemas.microsoft.com/exchange/services/2006/messages"' +
                   'xmlns:t="http://schemas.microsoft.com/exchange/services/2006/types"' +
                   'xmlns:soap="http://schemas.xmlsoap.org/soap/envelope/">' +
        '<soap:Header xmlns:wsse="http://docs.oasis-open.org/wss/2004/01/oasis-200401-wss-wssecurity-secext-1.0.xsd"' +
                     'xmlns:wsa="http://www.w3.org/2005/08/addressing">' +
            '<t:RequestServerVersion Version="Exchange2010_SP1"/>' +
        '</soap:Header>' +
        '<soap:Body>' +
            '<m:GetItem>' +
                '<m:ItemShape>' +
                    '<t:BaseShape>AllProperties</t:BaseShape>' +
                    '<t:IncludeMimeContent>true</t:IncludeMimeContent>' +
                    '<t:AdditionalProperties>' +
                        '<t:ExtendedFieldURI ' +
                          'DistinguishedPropertySetId="PublicStrings" ' +
                          'PropertyName="cecp-<app-guid>"' +
                          'PropertyType="String" ' +
                        '/>' +
                    '</t:AdditionalProperties>' +
                '</m:ItemShape>' +
                '<m:ItemIds>' +
                    '<t:ItemId Id="' +
                      Office.context.mailbox.item.itemId +
                    '"/>' +
                '</m:ItemIds>' +
            '</m:GetItem>' +
        '</soap:Body>' +
    '</soap:Envelope>';

Office.context.mailbox.makeEwsRequestAsync(
    request_str,
    function(asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded) {
            console.log(asyncResult.value);
        }
        else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

<span data-ttu-id="2f985-168">Также можно получить дополнительные настраиваемые свойства, если указать их в строке запроса как другие элементы [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri).</span><span class="sxs-lookup"><span data-stu-id="2f985-168">You can also get more custom properties if you specify them in the request string as other [ExtendedFieldURI](/exchange/client-developer/web-service-reference/extendedfielduri) elements.</span></span>

#### <a name="get-custom-properties-using-rest"></a><span data-ttu-id="2f985-169">Просмотр настраиваемых свойств с помощью REST</span><span class="sxs-lookup"><span data-stu-id="2f985-169">Get custom properties using REST</span></span>

<span data-ttu-id="2f985-170">В своей надстройке можно создать запрос REST для получения сообщений и событий, уже имеющих настраиваемые свойства.</span><span class="sxs-lookup"><span data-stu-id="2f985-170">In your add-in, you can construct your REST query against messages and events to get the ones that already have custom properties.</span></span> <span data-ttu-id="2f985-171">В запрос нужно включить расширенное свойство на основе интерфейса MAPI **CustomProperties** и его набор свойств с помощью сведений, указанных в разделе [Способ хранения настраиваемых свойств в элементе](#how-custom-properties-are-stored-on-an-item).</span><span class="sxs-lookup"><span data-stu-id="2f985-171">In your query, you should include the **CustomProperties** MAPI-based property and its property set using the details provided in the section [How custom properties are stored on an item](#how-custom-properties-are-stored-on-an-item).</span></span>

<span data-ttu-id="2f985-172">В приведенном ниже примере показано, как получить все события, которые содержат любые настраиваемые свойства, присвоенные вашей надстройкой, и обеспечить наличие в отклике значения свойства, чтобы в дальнейшем можно было применить логику фильтрации.</span><span class="sxs-lookup"><span data-stu-id="2f985-172">The following example shows how to get all events that have any custom properties set by your add-in and ensure that the response includes the value of the property so you can apply further filtering logic.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2f985-173">В приведенном ниже примере замените `<app-guid>` идентификатором своей надстройки.</span><span class="sxs-lookup"><span data-stu-id="2f985-173">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

```rest
GET https://outlook.office.com/api/v2.0/Me/Events?$filter=SingleValueExtendedProperties/Any
  (ep: ep/PropertyId eq 'String {00020329-0000-0000-C000-000000000046}
  Name cecp-<app-guid>' and ep/Value ne null)
  &$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String
  {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')
```

<span data-ttu-id="2f985-174">Другие примеры использования REST для получения однозначного расширенного свойства, основанного на интерфейсе MAPI, см. в статье [Получение объекта singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0).</span><span class="sxs-lookup"><span data-stu-id="2f985-174">For other examples that use REST to get single-value MAPI-based extended properties, see [Get singleValueExtendedProperty](/graph/api/singlevaluelegacyextendedproperty-get?view=graph-rest-1.0).</span></span>

<span data-ttu-id="2f985-175">В приведенном ниже примере показано, как получить элемент и его настраиваемые свойства.</span><span class="sxs-lookup"><span data-stu-id="2f985-175">The following example shows how to get an item and its custom properties.</span></span> <span data-ttu-id="2f985-176">В функции обратного вызова для метода `done` объект `item.SingleValueExtendedProperties` содержит список требуемых настраиваемых свойств.</span><span class="sxs-lookup"><span data-stu-id="2f985-176">In the callback function for the `done` method, `item.SingleValueExtendedProperties` contains a list of the requested custom properties.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2f985-177">В приведенном ниже примере замените `<app-guid>` идентификатором своей надстройки.</span><span class="sxs-lookup"><span data-stu-id="2f985-177">In the following example, replace `<app-guid>` with your add-in's ID.</span></span>

```typescript
Office.context.mailbox.getCallbackTokenAsync(
    {
        isRest: true
    },
    function (asyncResult) {
        if (asyncResult.status === Office.AsyncResultStatus.Succeeded
            && asyncResult.value !== "") {
            let item_rest_id = Office.context.mailbox.convertToRestId(
                Office.context.mailbox.item.itemId,
                Office.MailboxEnums.RestVersion.v2_0);
            let rest_url = Office.context.mailbox.restUrl +
                           "/v2.0/me/messages('" +
                           item_rest_id +
                           "')";
            rest_url += "?$expand=SingleValueExtendedProperties($filter=PropertyId eq 'String {00020329-0000-0000-C000-000000000046} Name cecp-<app-guid>')";

            let auth_token = asyncResult.value;
            $.ajax(
                {
                    url: rest_url,
                    dataType: 'json',
                    headers:
                        {
                            "Authorization":"Bearer " + auth_token
                        }
                }
                ).done(
                    function (item) {
                        console.log(JSON.stringify(item));
                    }
                ).fail(
                    function (error) {
                        console.log(JSON.stringify(error));
                    }
                );
        } else {
            console.log(JSON.stringify(asyncResult));
        }
    }
);
```

## <a name="see-also"></a><span data-ttu-id="2f985-178">См. также</span><span class="sxs-lookup"><span data-stu-id="2f985-178">See also</span></span>

- [<span data-ttu-id="2f985-179">Обзор свойств MAPI</span><span class="sxs-lookup"><span data-stu-id="2f985-179">MAPI Property Overview</span></span>](/office/client-developer/outlook/mapi/mapi-property-overview)
- [<span data-ttu-id="2f985-180">Обзор свойств Outlook</span><span class="sxs-lookup"><span data-stu-id="2f985-180">Outlook Properties Overview</span></span>](/office/vba/outlook/How-to/Navigation/properties-overview)  
- [<span data-ttu-id="2f985-181">Вызов REST API Outlook из надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="2f985-181">Call Outlook REST APIs from an Outlook add-in</span></span>](use-rest-api.md)
- [<span data-ttu-id="2f985-182">Вызов веб-служб из надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="2f985-182">Call web services from an Outlook add-in</span></span>](web-services.md)
- [<span data-ttu-id="2f985-183">Свойства и расширенные свойства в веб-службах Exchange</span><span class="sxs-lookup"><span data-stu-id="2f985-183">Properties and extended properties in EWS in Exchange</span></span>](/exchange/client-developer/exchange-web-services/properties-and-extended-properties-in-ews-in-exchange)
- [<span data-ttu-id="2f985-184">Наборы свойств и формы ответа в веб-службах Exchange</span><span class="sxs-lookup"><span data-stu-id="2f985-184">Property sets and response shapes in EWS in Exchange</span></span>](/exchange/client-developer/exchange-web-services/property-sets-and-response-shapes-in-ews-in-exchange)
