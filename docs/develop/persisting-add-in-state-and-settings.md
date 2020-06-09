---
title: Сохранение состояния и параметров надстройки
description: Сведения о том, как хранить данные в веб-приложениях надстройки Office, работающих в среде без сохранения состояния элемента управления браузера.
ms.date: 05/08/2020
localization_priority: Normal
ms.openlocfilehash: 81f149bdff540b236252a02a0c368799a11fed10
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609403"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="69b9a-103">Сохранение состояния и параметров надстройки</span><span class="sxs-lookup"><span data-stu-id="69b9a-103">Persisting add-in state and settings</span></span>

[!include[information about the common API](../includes/alert-common-api-info.md)]

<span data-ttu-id="69b9a-p101">Надстройки Office, по сути, представляют собой веб-приложения, которые выполняются в среде без сведений о состоянии элемента управления браузером. Вследствие этого надстройке может потребоваться сохранять данные для обеспечения непрерывности определенных операций или функций во время сеансов ее использования. Например, у надстройки могут быть настраиваемые параметры или другие значения, которые должны быть сохранены и повторно загружены при следующей инициализации, такие как выбранное пользователем представление или расположение по умолчанию. Это можно реализовать указанными ниже способами.</span><span class="sxs-lookup"><span data-stu-id="69b9a-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="69b9a-108">Используйте элементы API JavaScript для Office, которые хранят данные, как один из следующих:</span><span class="sxs-lookup"><span data-stu-id="69b9a-108">Use members of the Office JavaScript API that store data as either:</span></span>
    -  <span data-ttu-id="69b9a-109">пар имя-значение в контейнере свойств, расположение которого зависит от типа надстройки;</span><span class="sxs-lookup"><span data-stu-id="69b9a-109">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="69b9a-110">пользовательского кода XML в документе.</span><span class="sxs-lookup"><span data-stu-id="69b9a-110">Custom XML stored in the document.</span></span>

- <span data-ttu-id="69b9a-111">Использовать способы, предоставленные базовыми элементами управления браузером: cookie-файлы браузера или веб-хранилище HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) или [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span><span class="sxs-lookup"><span data-stu-id="69b9a-111">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>

<span data-ttu-id="69b9a-112">В этой статье рассказывается, как использовать API JavaScript для Office для сохранения состояния надстройки.</span><span class="sxs-lookup"><span data-stu-id="69b9a-112">This article focuses on how to use the Office JavaScript API to persist add-in state.</span></span> <span data-ttu-id="69b9a-113">Примеры использования файлов cookie браузера и веб-хранилища приведены в статье [Excel-Add-in-JavaScript-персисткустомсеттингс](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="69b9a-113">For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-office-javascript-api"></a><span data-ttu-id="69b9a-114">Сохранение состояния и параметров надстройки с помощью API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="69b9a-114">Persisting add-in state and settings with the Office JavaScript API</span></span>

<span data-ttu-id="69b9a-115">API JavaScript для Office предоставляет объекты [параметров](/javascript/api/office/office.settings), [roamingSettings](/javascript/api/outlook/office.roamingsettings)и [CustomProperties](/javascript/api/outlook/office.customproperties) для сохранения состояния надстройки во всех сеансах, как описано в следующей таблице.</span><span class="sxs-lookup"><span data-stu-id="69b9a-115">The Office JavaScript API provides the [Settings](/javascript/api/office/office.settings), [RoamingSettings](/javascript/api/outlook/office.roamingsettings), and [CustomProperties](/javascript/api/outlook/office.customproperties) objects for saving add-in state across sessions as described in the following table.</span></span> <span data-ttu-id="69b9a-116">Во всех случаях сохраненные значения параметров связаны с [Id](../reference/manifest/id.md) создавшей их надстройки.</span><span class="sxs-lookup"><span data-stu-id="69b9a-116">In all cases, the saved settings values are associated with the [Id](../reference/manifest/id.md) of the add-in that created them.</span></span>

|<span data-ttu-id="69b9a-117">**Объект**</span><span class="sxs-lookup"><span data-stu-id="69b9a-117">**Object**</span></span>|<span data-ttu-id="69b9a-118">**Поддерживаемый тип надстроек**</span><span class="sxs-lookup"><span data-stu-id="69b9a-118">**Add-in type support**</span></span>|<span data-ttu-id="69b9a-119">**Расположение хранилища**</span><span class="sxs-lookup"><span data-stu-id="69b9a-119">**Storage location**</span></span>|<span data-ttu-id="69b9a-120">**Поддержка ведущих приложений Office**</span><span class="sxs-lookup"><span data-stu-id="69b9a-120">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="69b9a-121">Параметры</span><span class="sxs-lookup"><span data-stu-id="69b9a-121">Settings</span></span>](/javascript/api/office/office.settings)|<span data-ttu-id="69b9a-122">Надстройки области задач и контентные надстройки</span><span class="sxs-lookup"><span data-stu-id="69b9a-122">content and task pane</span></span>|<span data-ttu-id="69b9a-123">Документ, электронная таблица или презентация, с которыми работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="69b9a-123">The document, spreadsheet, or presentation the add-in is working with.</span></span> <span data-ttu-id="69b9a-124">Параметры надстроек области задач и контентных надстроек доступны создавшей их надстройке в документе, в котором они сохранены.</span><span class="sxs-lookup"><span data-stu-id="69b9a-124">Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="69b9a-p105">**Внимание!** Не храните в объекте **Settings** пароли и другие конфиденциальные персональные данные. Сохраненные данные не видны пользователям, но содержатся документе, доступ к которому можно получить при прямом считывании. Необходимо ограничить использование надстройкой персональных данных и использовать для их хранения сервер, на котором эта надстройка размещена, как защищенный от пользователей ресурс.</span><span class="sxs-lookup"><span data-stu-id="69b9a-p105">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="69b9a-128">Word, Excel или PowerPoint</span><span class="sxs-lookup"><span data-stu-id="69b9a-128">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="69b9a-p106">**Примечание.** Надстройки области задач для Project 2013 не поддерживают API **Settings** для хранения данных о состоянии или параметров. Однако для надстроек, работающих в Project (а также в других ведущих приложениях Office), можно использовать cookie-файлы браузера или веб-хранилище. Дополнительные сведения об этих технологиях см. в статье [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="69b9a-p106">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="69b9a-132">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="69b9a-132">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings)|<span data-ttu-id="69b9a-133">Outlook</span><span class="sxs-lookup"><span data-stu-id="69b9a-133">Outlook</span></span>|<span data-ttu-id="69b9a-134">Почтовый ящик пользователя на сервере Exchange, на котором установлена надстройка.</span><span class="sxs-lookup"><span data-stu-id="69b9a-134">The user's Exchange server mailbox where the add-in is installed.</span></span> <span data-ttu-id="69b9a-135">Поскольку параметры сохраняются на сервере почтового ящика пользователя, они могут "перемещаться" с пользователем и доступны надстройке при запуске в контексте любого поддерживаемого клиентского ведущего приложения или браузера с получением доступа к почтовому ящику нужного пользователя.</span><span class="sxs-lookup"><span data-stu-id="69b9a-135">Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="69b9a-136">Параметры перемещения надстройки Outlook доступны только создавшей их надстройке и только в том почтовом ящике, в котором она установлена.</span><span class="sxs-lookup"><span data-stu-id="69b9a-136">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="69b9a-137">Outlook</span><span class="sxs-lookup"><span data-stu-id="69b9a-137">Outlook</span></span>|
|[<span data-ttu-id="69b9a-138">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="69b9a-138">CustomProperties</span></span>](/javascript/api/outlook/office.customproperties)|<span data-ttu-id="69b9a-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="69b9a-139">Outlook</span></span>|<span data-ttu-id="69b9a-p108">Элемент сообщения, встречи, запроса на собрание для которого была запущена надстройка. Пользовательские свойства элемента надстройки Outlook доступны только для создавшей их надстройки и только в элементе, в котором они сохранены.</span><span class="sxs-lookup"><span data-stu-id="69b9a-p108">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="69b9a-142">Outlook</span><span class="sxs-lookup"><span data-stu-id="69b9a-142">Outlook</span></span>|
|[<span data-ttu-id="69b9a-143">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="69b9a-143">CustomXmlParts</span></span>](/javascript/api/office/office.customxmlparts)|<span data-ttu-id="69b9a-144">Надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="69b9a-144">task pane</span></span>|<span data-ttu-id="69b9a-p109">Документ, электронная таблица или презентация, с которыми работает надстройка. Параметры надстроек области задач доступны создавшей их надстройке в том документе, где они сохранены.</span><span class="sxs-lookup"><span data-stu-id="69b9a-p109">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="69b9a-p110">**Внимание!** Не храните пароли и другие конфиденциальные личные сведения в пользовательской части XML. Сохраненные данные не видны пользователям, но содержатся в документе, доступ к которому можно получить при прямом считывании формата файла. Необходимо ограничить использование надстройкой личных сведений и хранить их только на том сервере, где размещена эта надстройка, так как этот ресурс защищен от пользователей.</span><span class="sxs-lookup"><span data-stu-id="69b9a-p110">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="69b9a-150">Word (с использованием общего API JavaScript для Office), Excel (с использованием специального API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="69b9a-150">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="69b9a-151">Данные параметров обрабатываются в памяти во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="69b9a-151">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="69b9a-p111">В следующих двух разделах рассматриваются параметры в контексте общего API JavaScript для Office. Специальный API JavaScript для Excel также предоставляет доступ к настраиваемым параметрам. Интерфейсы API Excel и шаблоны программирования слегка отличаются. Дополнительные сведения см. в статье [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).</span><span class="sxs-lookup"><span data-stu-id="69b9a-p111">The following two sections discuss settings in the context of the Office Common JavaScript API. The host-specific Excel JavaScript API also provides access to the custom settings. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel SettingCollection](/javascript/api/excel/excel.settingcollection).</span></span>

<span data-ttu-id="69b9a-156">Внутренние данные в контейнере свойств, доступ к которым осуществляется с `Settings` помощью `CustomProperties` объектов, или `RoamingSettings` объектов, хранятся как сериализованный объект нотации объектов JavaScript (JSON), содержащий пары "имя-значение".</span><span class="sxs-lookup"><span data-stu-id="69b9a-156">Internally, the data in the property bag accessed with the `Settings`, `CustomProperties`, or `RoamingSettings` objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs.</span></span> <span data-ttu-id="69b9a-157">Имя (ключ) для каждого значения должно иметь значение `string` , а хранимое значение может быть JavaScript,, `string` `number` `date` или `object` , но не **функцией**.</span><span class="sxs-lookup"><span data-stu-id="69b9a-157">The name (key) for each value must be a `string`, and the stored value can be a JavaScript `string`, `number`, `date`, or `object`, but not a **function**.</span></span>

<span data-ttu-id="69b9a-158">Пример структуры контейнера свойств, содержащего три определенных **строковых** значения с именами `firstName`, `location` и `defaultView`.</span><span class="sxs-lookup"><span data-stu-id="69b9a-158">This example of the property bag structure contains three defined **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="69b9a-159">После сохранения контейнера свойств параметров во время предыдущего сеанса надстройки он может быть загружен при инициализации надстройки или в любое время после этого в течение текущего сеанса надстройки.</span><span class="sxs-lookup"><span data-stu-id="69b9a-159">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session.</span></span> <span data-ttu-id="69b9a-160">Во время сеанса параметры полностью управляются в памяти с помощью `get` `set` методов, и `remove` объекта, соответствующего типу создаваемых параметров (**Settings**, **CustomProperties**или **roamingSettings**).</span><span class="sxs-lookup"><span data-stu-id="69b9a-160">During the session, the settings are managed in entirely in memory using the `get`, `set`, and `remove` methods of the object that corresponds to the kind of settings you are creating (**Settings**, **CustomProperties**, or **RoamingSettings**).</span></span>


> [!IMPORTANT]
> <span data-ttu-id="69b9a-161">Для сохранения добавлений, обновлений или удалений, внесенных в текущем сеансе надстройки, в место хранения необходимо вызвать `saveAsync` метод соответствующего объекта, который используется для работы с этими параметрами.</span><span class="sxs-lookup"><span data-stu-id="69b9a-161">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the `saveAsync` method of the corresponding object used to work with that kind of settings.</span></span> <span data-ttu-id="69b9a-162">`get`Методы, `set` и, Кроме того, `remove` работают только в копии контейнера свойств параметров, нашедшегося в памяти.</span><span class="sxs-lookup"><span data-stu-id="69b9a-162">The `get`, `set`, and `remove` methods operate only on the in-memory copy of the settings property bag.</span></span> <span data-ttu-id="69b9a-163">Если ваша надстройка закрывается без вызова `saveAsync` , любые изменения, внесенные в параметры во время этого сеанса, будут потеряны.</span><span class="sxs-lookup"><span data-stu-id="69b9a-163">If your add-in is closed without calling `saveAsync`, any changes made to settings during that session will be lost.</span></span>


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="69b9a-164">Сохранение состояния надстройки и параметров документа для контентных надстроек и надстроек области задач</span><span class="sxs-lookup"><span data-stu-id="69b9a-164">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="69b9a-165">Чтобы сохранить состояние или пользовательские параметры в контентной надстройке или надстройке области задач в Word, Excel или PowerPoint, следует использовать объект [Settings](/javascript/api/office/office.settings) и его методы.</span><span class="sxs-lookup"><span data-stu-id="69b9a-165">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](/javascript/api/office/office.settings) object and its methods.</span></span> <span data-ttu-id="69b9a-166">Контейнер свойств, созданный с помощью методов объекта, `Settings` доступен только для экземпляра созданной и созданной надстройкой области задач и только из документа, в котором она сохранена.</span><span class="sxs-lookup"><span data-stu-id="69b9a-166">The property bag created with the methods of the `Settings` object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="69b9a-167">`Settings`Объект автоматически загружается как часть объекта [Document](/javascript/api/office/office.document) и становится доступным при активации надстройки области задач или контентной надстройки.</span><span class="sxs-lookup"><span data-stu-id="69b9a-167">The `Settings` object is automatically loaded as part of the [Document](/javascript/api/office/office.document) object, and is available when the task pane or content add-in is activated.</span></span> <span data-ttu-id="69b9a-168">После `Document` создания экземпляра объекта можно получить доступ к `Settings` объекту с помощью свойства [Settings](/javascript/api/office/office.document#settings) `Document` объекта.</span><span class="sxs-lookup"><span data-stu-id="69b9a-168">After the `Document` object is instantiated, you can access the `Settings` object with the [settings](/javascript/api/office/office.document#settings) property of the `Document` object.</span></span> <span data-ttu-id="69b9a-169">Во время существования сеанса можно просто использовать `Settings.get` `Settings.set` методы, и и `Settings.remove` для чтения, записи или удаления сохраненных параметров и состояния надстройки из копии контейнера свойств в памяти.</span><span class="sxs-lookup"><span data-stu-id="69b9a-169">During the lifetime of the session, you can just use the `Settings.get`, `Settings.set`, and `Settings.remove` methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="69b9a-170">Поскольку методы "set" и "remove" работают только в копии контейнера свойств параметров, содержащейся в памяти, для сохранения новых или измененных параметров документа, с которым сопоставлена надстройка, необходимо вызвать метод [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-).</span><span class="sxs-lookup"><span data-stu-id="69b9a-170">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="69b9a-171">Создание или обновление значения параметра</span><span class="sxs-lookup"><span data-stu-id="69b9a-171">Creating or updating a setting value</span></span>

<span data-ttu-id="69b9a-p117">Следующий пример кода демонстрирует использование метода [Settings.set](/javascript/api/office/office.settings#set-name--value-) для создания параметра с именем `'themeColor'`, имеющий значение  `'green'`. Первый параметр этого метода — это зависящий от регистра идентификатор  _name_ параметра, который следует определить или создать. Второй параметр — это _value_ параметра.</span><span class="sxs-lookup"><span data-stu-id="69b9a-p117">The following code example shows how to use the [Settings.set](/javascript/api/office/office.settings#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="69b9a-175">Создается параметр с указанным именем, если таковой еще не существует или обновляется значение, если параметр существует.</span><span class="sxs-lookup"><span data-stu-id="69b9a-175">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist.</span></span> <span data-ttu-id="69b9a-176">Используйте `Settings.saveAsync` метод для сохранения новых или обновленных параметров в документе.</span><span class="sxs-lookup"><span data-stu-id="69b9a-176">Use the `Settings.saveAsync` method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="69b9a-177">Получение значения параметра</span><span class="sxs-lookup"><span data-stu-id="69b9a-177">Getting the value of a setting</span></span>

<span data-ttu-id="69b9a-178">В следующем примере показано, как использовать метод [Settings.get](/javascript/api/office/office.settings#get-name-) для получения значения параметра "themeColor".</span><span class="sxs-lookup"><span data-stu-id="69b9a-178">The following example shows how use the [Settings.get](/javascript/api/office/office.settings#get-name-) method to get the value of a setting called "themeColor".</span></span> <span data-ttu-id="69b9a-179">Единственный параметр `get` метода — это _имя_ параметра с учетом регистра.</span><span class="sxs-lookup"><span data-stu-id="69b9a-179">The only parameter of the `get` method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="69b9a-180">`get`Метод возвращает значение, сохраненное ранее для _имени_ параметра, которое было передано.</span><span class="sxs-lookup"><span data-stu-id="69b9a-180">The `get` method returns the value that was previously saved for the setting _name_ that was passed in.</span></span> <span data-ttu-id="69b9a-181">Если параметр не существует, метод возвращает **null**.</span><span class="sxs-lookup"><span data-stu-id="69b9a-181">If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="69b9a-182">Удаление параметра</span><span class="sxs-lookup"><span data-stu-id="69b9a-182">Removing a setting</span></span>

<span data-ttu-id="69b9a-183">В следующем примере показано, как использовать метод [Settings.remove](/javascript/api/office/office.settings#remove-name-) для удаления параметра с именем "themeColor".</span><span class="sxs-lookup"><span data-stu-id="69b9a-183">The following example shows how to use the [Settings.remove](/javascript/api/office/office.settings#remove-name-) method to remove a setting with the name "themeColor".</span></span> <span data-ttu-id="69b9a-184">Единственный параметр `remove` метода — это _имя_ параметра с учетом регистра.</span><span class="sxs-lookup"><span data-stu-id="69b9a-184">The only parameter of the `remove` method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="69b9a-185">Если параметр не существует, ничего не произойдет.</span><span class="sxs-lookup"><span data-stu-id="69b9a-185">Nothing will happen if the setting does not exist.</span></span> <span data-ttu-id="69b9a-186">Используйте `Settings.saveAsync` метод для сохранения удаления параметра из документа.</span><span class="sxs-lookup"><span data-stu-id="69b9a-186">Use the `Settings.saveAsync` method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="69b9a-187">Сохранение параметров</span><span class="sxs-lookup"><span data-stu-id="69b9a-187">Saving your settings</span></span>

<span data-ttu-id="69b9a-188">Чтобы сохранить любые добавления, изменения или удаления, внесенные надстройкой в копию контейнера свойств параметров, хранящуюся в памяти, во время текущего сеанса надстройки, необходимо вызвать метод [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) для их сохранения в документе.</span><span class="sxs-lookup"><span data-stu-id="69b9a-188">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](/javascript/api/office/office.settings#saveasync-options--callback-) method to store them in the document.</span></span> <span data-ttu-id="69b9a-189">Единственный параметр `saveAsync` метода — _обратный вызов_, который является функцией обратного вызова с одним параметром.</span><span class="sxs-lookup"><span data-stu-id="69b9a-189">The only parameter of the `saveAsync` method is _callback_, which is a callback function with a single parameter.</span></span> 


```js
Office.context.document.settings.saveAsync(function (asyncResult) {
    if (asyncResult.status == Office.AsyncResultStatus.Failed) {
        write('Settings save failed. Error: ' + asyncResult.error.message);
    } else {
        write('Settings saved.');
    }
});
// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

<span data-ttu-id="69b9a-190">Анонимная функция, передаваемая в `saveAsync` метод в качестве параметра _callback_ , выполняется по завершении операции.</span><span class="sxs-lookup"><span data-stu-id="69b9a-190">The anonymous function passed into the `saveAsync` method as the _callback_ parameter is executed when the operation is completed.</span></span> <span data-ttu-id="69b9a-191">Параметр _asyncResult_ обратного вызова предоставляет доступ к `AsyncResult` объекту, который содержит состояние операции.</span><span class="sxs-lookup"><span data-stu-id="69b9a-191">The _asyncResult_ parameter of the callback provides access to an `AsyncResult` object that contains the status of the operation.</span></span> <span data-ttu-id="69b9a-192">В этом примере функция проверяет `AsyncResult.status` свойство, чтобы убедиться, что операция сохранения выполнена успешно или не выполнена, а затем отображает результат на странице надстройки.</span><span class="sxs-lookup"><span data-stu-id="69b9a-192">In the example, the function checks the `AsyncResult.status` property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="69b9a-193">Сохранение пользовательского кода XML в документе</span><span class="sxs-lookup"><span data-stu-id="69b9a-193">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="69b9a-p125">В этом разделе рассматриваются пользовательские части XML в контексте общего API JavaScript для Office, поддерживаемого в Word. Специальный API JavaScript для Excel также предоставляет доступ к пользовательским частям XML. Интерфейсы API Excel и шаблоны программирования слегка отличаются. Дополнительные сведения см. в статье [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span><span class="sxs-lookup"><span data-stu-id="69b9a-p125">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word. The host-specific Excel JavaScript API also provides access to the custom XML parts. The Excel APIs and programming patterns are somewhat different. For more information, see [Excel CustomXmlPart](/javascript/api/excel/excel.customxmlpart).</span></span>

<span data-ttu-id="69b9a-198">Если требуется сохранить данные, размер которых превышает ограничения для параметров документа, или структурированные данные, то используется дополнительный параметр хранения.</span><span class="sxs-lookup"><span data-stu-id="69b9a-198">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character.</span></span> <span data-ttu-id="69b9a-199">Вы можете сохранять пользовательскую разметку XML в надстройке области задач для Word (а также для Excel, но следует учитывать примечание в начале этого раздела).</span><span class="sxs-lookup"><span data-stu-id="69b9a-199">You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section).</span></span> <span data-ttu-id="69b9a-200">В Word можно использовать объект [CustomXmlPart](/javascript/api/office/office.customxmlpart) и его методы (еще раз, см. примечание для Excel выше).</span><span class="sxs-lookup"><span data-stu-id="69b9a-200">In Word, you use the [CustomXmlPart](/javascript/api/office/office.customxmlpart) object and its methods (again, see the note above for Excel).</span></span> <span data-ttu-id="69b9a-201">В приведенном ниже коде создается пользовательская часть XML, после чего в разделителях на странице отображается сначала ее ИД, а затем ее содержимое.</span><span class="sxs-lookup"><span data-stu-id="69b9a-201">The following code creates a custom XML part and displays its ID and then its content in divs on the page.</span></span> <span data-ttu-id="69b9a-202">Обратите внимание, что в строке XML должен быть указан атрибут `xmlns`.</span><span class="sxs-lookup"><span data-stu-id="69b9a-202">Note that there must be an `xmlns` attribute in the XML string.</span></span>

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.value.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

<span data-ttu-id="69b9a-p127">Чтобы получить пользовательскую часть XML, используйте метод [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-). Однако ИД — это GUID, генерируемый при создании части XML, поэтому его невозможно узнать во время написания кода. По этой причине при создании части XML рекомендуется сразу сохранить ее ИД в виде параметра с запоминающимся идентификатором. Ниже показано, как это сделать. В предыдущих разделах этой статьи вы найдете подробные сведения и рекомендации по работе с настраиваемыми параметрами.</span><span class="sxs-lookup"><span data-stu-id="69b9a-p127">To retrieve a custom XML part, you use the [getByIdAsync](/javascript/api/office/office.customxmlparts#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is. For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key. The following method shows how to do this. (But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

 ```js
function createCustomXmlPartAndStoreId() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            Office.context.document.settings.set('ReviewersID', asyncResult.id);
            Office.context.document.settings.saveAsync();
        }
    );
}
```

<span data-ttu-id="69b9a-207">В приведенном ниже коде показано, как получить часть XML, сначала получив ее ИД из параметра.</span><span class="sxs-lookup"><span data-stu-id="69b9a-207">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID');
    Office.context.document.customXmlParts.getByIdAsync(reviewersXmlId,
        (asyncResult) => {
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);
                }
            );
        }
    );
}
```

## <a name="how-to-save-settings-in-an-outlook-add-in"></a><span data-ttu-id="69b9a-208">Сохранение параметров в надстройке Outlook</span><span class="sxs-lookup"><span data-stu-id="69b9a-208">How to save settings in an Outlook add-in</span></span>

<span data-ttu-id="69b9a-209">Сведения о том, как сохранить параметры в надстройке Outlook, можно узнать в статье [Управление состоянием и настройками надстройки Outlook](../outlook/manage-state-and-settings-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="69b9a-209">For information about how to save settings in an Outlook add-in, see [Manage state and settings for an Outlook add-in](../outlook/manage-state-and-settings-outlook.md).</span></span>


## <a name="see-also"></a><span data-ttu-id="69b9a-210">См. также</span><span class="sxs-lookup"><span data-stu-id="69b9a-210">See also</span></span>

- [<span data-ttu-id="69b9a-211">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="69b9a-211">Understanding the Office JavaScript API</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="69b9a-212">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="69b9a-212">Outlook add-ins</span></span>](../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="69b9a-213">Управление состоянием и параметрами для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="69b9a-213">Manage state and settings for an Outlook add-in</span></span>](../outlook/manage-state-and-settings-outlook.md)
- [<span data-ttu-id="69b9a-214">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="69b9a-214">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
