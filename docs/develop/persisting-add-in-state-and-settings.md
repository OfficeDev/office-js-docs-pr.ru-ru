---
title: Сохранение состояния и параметров надстройки
description: ''
ms.date: 12/04/2017
ms.openlocfilehash: 7ea35f00809fbe960155137c7cdae3f6dfd60b90
ms.sourcegitcommit: 30435939ab8b8504c3dbfc62fd29ec6b0f1a7d22
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/12/2018
ms.locfileid: "23945384"
---
# <a name="persisting-add-in-state-and-settings"></a><span data-ttu-id="2b2f3-102">Сохранение состояния и параметров надстройки</span><span class="sxs-lookup"><span data-stu-id="2b2f3-102">Persisting add-in state and settings</span></span>

<span data-ttu-id="2b2f3-p101">Надстройки Office, по сути, представляют собой веб-приложения, которые выполняются в среде без сведений о состоянии элемента управления браузером. Вследствие этого надстройке может потребоваться сохранять данные для обеспечения непрерывности определенных операций или функций во время сеансов ее использования. Например, у надстройки могут быть настраиваемые параметры или другие значения, которые должны быть сохранены и повторно загружены при следующей инициализации, такие как выбранное пользователем представление или расположение по умолчанию. Это можно реализовать указанными ниже способами.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p101">Office Add-ins are essentially web applications running in the stateless environment of a browser control. As a result, your add-in may need to persist data to maintain the continuity of certain operations or features across sessions of using your add-in. For example, your add-in may have custom settings or other values that it needs to save and reload the next time it's initialized, such as a user's preferred view or default location. To do that, you can:</span></span>

- <span data-ttu-id="2b2f3-107">Использовать элементы API JavaScript для Office, чтобы хранить данные в виде:</span><span class="sxs-lookup"><span data-stu-id="2b2f3-107">Use members of the JavaScript API for Office that store data as either:</span></span>
    -  <span data-ttu-id="2b2f3-108">пар имя-значение в контейнере свойств, расположение которого зависит от типа надстройки;</span><span class="sxs-lookup"><span data-stu-id="2b2f3-108">Name/value pairs in a property bag stored in a location that depends on add-in type.</span></span>
    -  <span data-ttu-id="2b2f3-109">пользовательского кода XML в документе.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-109">Custom XML stored in the document.</span></span>
    
- <span data-ttu-id="2b2f3-110">Использовать способы, предоставленные базовыми элементами управления браузером: cookie-файлы браузера или веб-хранилище HTML5 ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) или [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span><span class="sxs-lookup"><span data-stu-id="2b2f3-110">Use techniques provided by the underlying browser control: browser cookies, or HTML5 web storage ([localStorage](https://developer.mozilla.org/docs/Web/API/Window/localStorage) or [sessionStorage](https://developer.mozilla.org/docs/Web/API/Window/sessionStorage)).</span></span>
    
<span data-ttu-id="2b2f3-p102">Эта статья содержит сведения об использовании API JavaScript для сохранения состояния надстройки. Примеры использования cookie-файлов браузера и веб-хранилища см. в примере кода [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p102">This article focuses on how to use the JavaScript API for Office to persist add-in state. For examples of using browser cookies and web storage, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span>

## <a name="persisting-add-in-state-and-settings-with-the-javascript-api-for-office"></a><span data-ttu-id="2b2f3-113">Сохранение состояния и параметров надстройки с помощью JavaScript API для Office</span><span class="sxs-lookup"><span data-stu-id="2b2f3-113">Persisting add-in state and settings with the JavaScript API for Office</span></span>

<span data-ttu-id="2b2f3-p103">API JavaScript для Office предоставляет объекты [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js), [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) и [CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js) для сохранения состояния надстройки во время сеансов, как показано в следующей таблице. Во всех случаях сохраненные значения параметров связаны с [Id](https://docs.microsoft.com/javascript/office/manifest/id?view=office-js) создавшей их надстройки.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p103">The JavaScript API for Office provides the [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js), [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js), and [CustomProperties](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js) objects for saving add-in state across sessions as described in the following table. In all cases, the saved settings values are associated with the [Id](https://docs.microsoft.com/javascript/office/manifest/id?view=office-js) of the add-in that created them.</span></span>

|<span data-ttu-id="2b2f3-116">**Объект**</span><span class="sxs-lookup"><span data-stu-id="2b2f3-116">**Object**</span></span>|<span data-ttu-id="2b2f3-117">**Поддерживаемый тип надстроек**</span><span class="sxs-lookup"><span data-stu-id="2b2f3-117">**Add-in type support**</span></span>|<span data-ttu-id="2b2f3-118">**Расположение хранилища**</span><span class="sxs-lookup"><span data-stu-id="2b2f3-118">**Storage location**</span></span>|<span data-ttu-id="2b2f3-119">**Поддержка ведущих приложений Office**</span><span class="sxs-lookup"><span data-stu-id="2b2f3-119">**Office host support**</span></span>|
|:-----|:-----|:-----|:-----|
|[<span data-ttu-id="2b2f3-120">Параметры</span><span class="sxs-lookup"><span data-stu-id="2b2f3-120">Settings</span></span>](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js)|<span data-ttu-id="2b2f3-121">Надстройки области задач и контентные надстройки</span><span class="sxs-lookup"><span data-stu-id="2b2f3-121">content and task pane</span></span>|<span data-ttu-id="2b2f3-122">Документ, электронная таблица или презентация, с которой работает надстройка. Параметры надстроек области задач и контентных надстроек доступны создавшей их надстройке в том документе, где они сохранены.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-122">The document, spreadsheet, or presentation the add-in is working with.Content and task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="2b2f3-p104">**Внимание!** Не храните в объекте **Settings** пароли и другие конфиденциальные персональные данные. Сохраненные данные не видны пользователям, но содержатся документе, доступ к которому можно получить при прямом считывании. Необходимо ограничить использование надстройкой персональных данных и использовать для их хранения сервер, на котором эта надстройка размещена, как защищенный от пользователей ресурс.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p104">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) with the **Settings** object. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="2b2f3-126">Word, Excel или PowerPoint</span><span class="sxs-lookup"><span data-stu-id="2b2f3-126">Word, Excel, or PowerPoint</span></span><br/><br/> <span data-ttu-id="2b2f3-p105">**Примечание.** Надстройки области задач для Project 2013 не поддерживают API **Settings** для хранения данных о состоянии или параметров. Однако для надстроек, работающих в Project (а также в других ведущих приложениях Office), можно использовать cookie-файлы браузера или веб-хранилище. Дополнительные сведения об этих технологиях см. в статье [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p105">**Note:** Task pane add-ins for Project 2013 don't support the **Settings** API for storing add-in state or settings. However, for add-ins running in Project (as well as other Office host applications) you can use techniques such as browser cookies or web storage. For more information on these techniques, see the [Excel-Add-in-JavaScript-PersistCustomSettings](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings).</span></span> |
|[<span data-ttu-id="2b2f3-130">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="2b2f3-130">RoamingSettings</span></span>](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js)|<span data-ttu-id="2b2f3-131">Outlook</span><span class="sxs-lookup"><span data-stu-id="2b2f3-131">Outlook</span></span>|<span data-ttu-id="2b2f3-132">Почтовый ящик пользователя на сервере Exchange Server, где установлена надстройка. Так как параметры сохраняются на сервере почтового ящика пользователя, они могут "перемещаться" с пользователем и доступны надстройке при запуске в контексте любого поддерживаемого клиентского ведущего приложения или браузера, получающего доступ к почтовому ящику этого пользователя.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-132">The user's Exchange server mailbox where the add-in is installed.Because these settings are stored in the user's server mailbox, they can "roam" with the user and are available to the add-in when it is running in the context of any supported client host application or browser accessing that user's mailbox.</span></span><br/><br/> <span data-ttu-id="2b2f3-133">Параметры перемещения надстройки Outlook доступны только создавшей их надстройке и только в том почтовом ящике, в котором она установлена.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-133">Outlook add-in roaming settings are available only to the add-in that created them, and only from the mailbox where the add-in is installed.</span></span>|<span data-ttu-id="2b2f3-134">Outlook</span><span class="sxs-lookup"><span data-stu-id="2b2f3-134">Outlook</span></span>|
|[<span data-ttu-id="2b2f3-135">CustomProperties</span><span class="sxs-lookup"><span data-stu-id="2b2f3-135">CustomProperties</span></span>](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js)|<span data-ttu-id="2b2f3-136">Outlook</span><span class="sxs-lookup"><span data-stu-id="2b2f3-136">Outlook</span></span>|<span data-ttu-id="2b2f3-p106">Элемент сообщения, встречи, запроса на собрание для которого была запущена надстройка. Пользовательские свойства элемента надстройки Outlook доступны только для создавшей их надстройки и только в элементе, в котором они сохранены.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p106">The message, appointment, or meeting request item the add-in is working with. Outlook add-in item custom properties are available only to the add-in that created them, and only from the item where they are saved.</span></span>|<span data-ttu-id="2b2f3-139">Outlook</span><span class="sxs-lookup"><span data-stu-id="2b2f3-139">Outlook</span></span>|
|[<span data-ttu-id="2b2f3-140">CustomXmlParts</span><span class="sxs-lookup"><span data-stu-id="2b2f3-140">customXmlParts</span></span>](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js)|<span data-ttu-id="2b2f3-141">Надстройки области задач</span><span class="sxs-lookup"><span data-stu-id="2b2f3-141">task pane</span></span>|<span data-ttu-id="2b2f3-p107">Документ, электронная таблица или презентация, с которыми работает надстройка. Параметры надстроек области задач доступны создавшей их надстройке в том документе, где они сохранены.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p107">The document, spreadsheet, or presentation the add-in is working with. Task pane add-in settings are available to the add-in that created them from the document where they are saved.</span></span><br/><br/><span data-ttu-id="2b2f3-p108">**Внимание!** Не храните пароли и другие конфиденциальные личные сведения в пользовательской части XML. Сохраненные данные не видны пользователям, но содержатся в документе, доступ к которому можно получить при прямом считывании формата файла. Необходимо ограничить использование надстройкой личных сведений и хранить их только на том сервере, где размещена эта надстройка, так как этот ресурс защищен от пользователей.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p108">**Important:** Don't store passwords and other sensitive personally identifiable information (PII) in a custom XML part. The data saved isn't visible to end users, but it is stored as part of the document, which is accessible by reading the document's file format directly. You should limit your add-in's use of PII and store any PII required by your add-in only on the server hosting your add-in as a user-secured resource.</span></span>|<span data-ttu-id="2b2f3-147">Word (с использованием общего API JavaScript для Office), Excel (с использованием специального API JavaScript для Excel)</span><span class="sxs-lookup"><span data-stu-id="2b2f3-147">Word (using the Office JavaScript Common API) Excel (using the host-specific Excel JavaScript API</span></span>|

## <a name="settings-data-is-managed-in-memory-at-runtime"></a><span data-ttu-id="2b2f3-148">Данные параметров обрабатываются в памяти во время выполнения.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-148">Settings data is managed in memory at runtime</span></span>

> [!NOTE]
> <span data-ttu-id="2b2f3-149">В следующих двух разделах рассматриваются параметры в контексте общего API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-149">The following two sections discuss settings in the context of the Office Common JavaScript API.</span></span> <span data-ttu-id="2b2f3-150">Специальный API JavaScript для Excel также предоставляет доступ к настраиваемым параметрам.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-150">The host-specific Excel JavaScript API also provides access to the custom settings.</span></span> <span data-ttu-id="2b2f3-151">Интерфейсы API Excel и шаблоны программирования слегка отличаются.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-151">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="2b2f3-152">Дополнительные сведения см. в статье [Excel SettingCollection](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="2b2f3-152">For more information, see [Excel SettingCollection](https://docs.microsoft.com/javascript/api/excel/excel.settingcollection?view=office-js).</span></span>

<span data-ttu-id="2b2f3-p110">Для внутренних целей данные в контейнере свойств, доступных с помощью объектов**Settings**,**CustomProperties** или **RoamingSettings**, сохраняются в качестве сериализованного объекта JSON, содержащего пары "имя-значение". Имя (ключ) для каждого значения должно быть**string** и значение, сохраненное в свойстве, может быть JavaScript **string**,**number**, **date** или **object**, но не должно быть**function**.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p110">Internally, the data in the property bag accessed with the  **Settings**,  **CustomProperties**, or  **RoamingSettings** objects is stored as a serialized JavaScript Object Notation (JSON) object that contains name/value pairs. The name (key) for each value must be a **string**, and the stored value can be a JavaScript  **string**,  **number**,  **date**, or  **object**, but not a  **function**.</span></span>

<span data-ttu-id="2b2f3-155">Пример структуры контейнера свойств, содержащего три определенных значения  **string** с именами `firstName`,  `location` и `defaultView`.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-155">This example of the property bag structure contains three defined  **string** values named `firstName`,  `location`, and  `defaultView`.</span></span>

```json
{
    "firstName":"Erik",
    "location":"98052",
    "defaultView":"basic"
}
```

<span data-ttu-id="2b2f3-p111">После сохранения контейнера свойств параметров во время предыдущего сеанса надстройки, он может быть загружен при инициализации надстройки или в любое время после этого в течение текущего сеанса приложения. Во время сеанса параметры изменяются только в памяти с помощью методов объекта **get**,**set** и **remove**, соответствующего типу создаваемых параметров ( **Settings**,**CustomProperties** или **RoamingSettings**).</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p111">After the settings property bag is saved during the previous add-in session, it can be loaded when the add-in is initialized or at any point after that during the add-in's current session. During the session, the settings are managed in entirely in memory using the  **get**,  **set**, and  **remove** methods of the object that corresponds to the kind settings you are creating ( **Settings**,  **CustomProperties**, or  **RoamingSettings**).</span></span> 


> [!IMPORTANT]
> <span data-ttu-id="2b2f3-p112">Чтобы операции добавления, обновления и удаления, выполненные в текущем сеансе надстройки, не были отменены, необходимо вызвать метод **saveAsync** соответствующего объекта, используемого для работы с заданным типом параметров. Методы **get**, **set** и **remove** работают только в копии контейнера свойств параметров, содержащейся в памяти. Если закрыть надстройку, не вызывая метод **saveAsync**, то все изменения, внесенные в параметры во время сеанса, будут потеряны.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p112">To persist any additions, updates, or deletions made during the add-in's current session to the storage location, you must call the  **saveAsync** method of the corresponding object used to work with that kind of settings. The **get**,  **set**, and  **remove** methods operate only on the in-memory copy of the settings property bag. If your add-in is closed without calling **saveAsync**, any changes made to settings during that session will be lost.</span></span> 


## <a name="how-to-save-add-in-state-and-settings-per-document-for-content-and-task-pane-add-ins"></a><span data-ttu-id="2b2f3-161">Сохранение состояния надстройки и параметров документа для надстроек области задач и контентных надстроек</span><span class="sxs-lookup"><span data-stu-id="2b2f3-161">How to save add-in state and settings per document for content and task pane add-ins</span></span>


<span data-ttu-id="2b2f3-p113">Чтобы сохранить состояние или пользовательские параметры в контентной надстройке или надстройке области задач в Word, Excel или PowerPoint, следует использовать объект [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) и его методы. Контейнер свойств, созданный с помощью методов объекта **Settings**, доступен только тому экземпляру контентной надстройки или надстройки области задач, который создал этот контейнер, и только в том документе, где он сохранен.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p113">To persist state or custom settings of a content or task pane add-in for Word, Excel, or PowerPoint, you use the [Settings](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js) object and its methods. The property bag created with the methods of the **Settings** object are available only to the instance of the content or task pane add-in that created it, and only from the document in which it is saved.</span></span>

<span data-ttu-id="2b2f3-p114">Объект**Settings** автоматически загружается как часть объекта [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) и доступен при активации надстройки области задач или контентной надстройки. После создания экземпляра объекта **Document**, вы можете получить доступ к объекту **Settings** с помощью свойства [settings](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings) объекта **Document**. Во время действия сеанса можно использовать методы**Settings.get**,**Settings.set** и **Settings.remove** для чтения, записи или удаления сохраненных параметров и состояния надстройки из копии контейнера свойств, содержащейся в памяти.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p114">The  **Settings** object is automatically loaded as part of the [Document](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js) object, and is available when the task pane or content add-in is activated. After the **Document** object is instantiated, you can access the **Settings** object with the [settings](https://docs.microsoft.com/javascript/api/office/office.document?view=office-js#settings) property of the **Document** object. During the lifetime of the session, you can just use the **Settings.get**,  **Settings.set**, and  **Settings.remove** methods to read, write, or remove persisted settings and add-in state from the in-memory copy of the property bag.</span></span>

<span data-ttu-id="2b2f3-167">Поскольку методы "set" и "remove" работают только в копии контейнера свойств параметров, содержащейся в памяти, для сохранения новых или измененных параметров документа, с которым сопоставлена надстройка, необходимо вызвать метод [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-).</span><span class="sxs-lookup"><span data-stu-id="2b2f3-167">Because the set and remove methods operate against only the in-memory copy of the settings property bag, to save new or changed settings back to the document the add-in is associated with you must call the [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-) method.</span></span>


### <a name="creating-or-updating-a-setting-value"></a><span data-ttu-id="2b2f3-168">Создание или обновление значения параметра</span><span class="sxs-lookup"><span data-stu-id="2b2f3-168">Creating or updating a setting value</span></span>

<span data-ttu-id="2b2f3-p115">Следующий пример кода демонстрирует использование метода [Settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#set-name--value-) для создания параметра с именем `'themeColor'`, имеющий значение  `'green'`. Первый параметр этого метода — это зависящий от регистра идентификатор  _name_ параметра, который следует определить или создать. Второй параметр — это _value_ параметра.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p115">The following code example shows how to use the [Settings.set](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#set-name--value-) method to create a setting called `'themeColor'` with a value `'green'`. The first parameter of the set method is the case-sensitive  _name_ (Id) of the setting to set or create. The second parameter is the _value_ of the setting.</span></span>


```js
Office.context.document.settings.set('themeColor', 'green');
```

 <span data-ttu-id="2b2f3-p116">Создается параметр с указанным именем, если таковой еще не существует или обновляется значение, если параметр существует. Используйте метод **Settings.saveAsync** для сохранения новых или обновления существующих параметров документа.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p116">The setting with the specified name is created if it doesn't already exist, or its value is updated if it does exist. Use the **Settings.saveAsync** method to persist the new or updated settings to the document.</span></span>


### <a name="getting-the-value-of-a-setting"></a><span data-ttu-id="2b2f3-174">Получение значения параметра</span><span class="sxs-lookup"><span data-stu-id="2b2f3-174">Getting the value of a setting</span></span>

<span data-ttu-id="2b2f3-p117">В следующем примере показано, как использовать метод [Settings.get](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#get-name-) для получения значения параметра "themeColor". Единственным параметром метода **get** является зависящий от регистра параметр _name_.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p117">The following example shows how use the [Settings.get](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#get-name-) method to get the value of a setting called "themeColor". The only parameter of the **get** method is the case-sensitive _name_ of the setting.</span></span>


```js
write('Current value for mySetting: ' + Office.context.document.settings.get('themeColor'));

// Function that writes to a div with id='message' on the page.
function write(message){
    document.getElementById('message').innerText += message; 
}
```

 <span data-ttu-id="2b2f3-p118">Метод **get** возвращает значение, которое было ранее сохранено для переданного параметра _name_. Если параметр не существует, метод возвращает **null**.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p118">The **get** method returns the value that was previously saved for the setting _name_ that was passed in. If the setting doesn't exist, the method returns **null**.</span></span>


### <a name="removing-a-setting"></a><span data-ttu-id="2b2f3-179">Удаление параметра</span><span class="sxs-lookup"><span data-stu-id="2b2f3-179">Removing a setting</span></span>

<span data-ttu-id="2b2f3-p119">В следующем примере показано, как использовать метод [Settings.remove](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#remove-name-) для удаления параметра с именем "themeColor". Единственным параметром метода **remove** является зависящий от регистра параметр _name_.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p119">The following example shows how to use the [Settings.remove](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#remove-name-) method to remove a setting with the name "themeColor". The only parameter of the **remove** method is the case-sensitive _name_ of the setting.</span></span>


```js
Office.context.document.settings.remove('themeColor');
```

<span data-ttu-id="2b2f3-p120">Если параметр не существует, ничего не произойдет. Используйте метод**Settings.saveAsync** чтобы предотвратить удаление указанного параметра в документе.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p120">Nothing will happen if the setting does not exist. Use the  **Settings.saveAsync** method to persist removal of the setting from the document.</span></span>


### <a name="saving-your-settings"></a><span data-ttu-id="2b2f3-184">Сохранение параметров</span><span class="sxs-lookup"><span data-stu-id="2b2f3-184">Saving your settings</span></span>

<span data-ttu-id="2b2f3-p121">Чтобы сохранить любые добавления, изменения или удаления, внесенные надстройкой в копию контейнера свойств параметров, хранящуюся в памяти, во время текущего сеанса надстройки, необходимо вызвать метод [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-) для их сохранения в документе. Единственный параметр метода **saveAsync** — это _callback_, представляющий собой функцию обратного вызова с одним параметром.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p121">To save any additions, changes, or deletions your add-in made to the in-memory copy of the settings property bag during the current session, you must call the [Settings.saveAsync](https://docs.microsoft.com/javascript/api/office/office.settings?view=office-js#saveasync-options--callback-) method to store them in the document. The only parameter of the **saveAsync** method is _callback_, which is a callback function with a single parameter.</span></span> 


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

<span data-ttu-id="2b2f3-p122">Анонимная функция, переданная в метод**saveAsync** в качестве параметра _callback_, выполняется после завершения операции. Параметр обратного вызова _asyncResult_ предоставляет доступ к объекту **AsyncResult**, содержащему сведения о состоянии операции. В этом примере функция проверяет свойство  **AsyncResult.status** для проверки успешного или неудачного выполнения операции с последующим отображением результата на странице надстройки.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p122">The anonymous function passed into the  **saveAsync** method as the _callback_ parameter is executed when the operation is completed. The _asyncResult_ parameter of the callback provides access to an **AsyncResult** object that contains the status of the operation. In the example, the function checks the **AsyncResult.status** property to see if the save operation succeeded or failed, and then displays the result in the add-in's page.</span></span>

## <a name="how-to-save-custom-xml-to-the-document"></a><span data-ttu-id="2b2f3-190">Сохранение пользовательского кода XML в документе</span><span class="sxs-lookup"><span data-stu-id="2b2f3-190">How to save custom XML to the document</span></span>

> [!NOTE]
> <span data-ttu-id="2b2f3-191">В этом разделе рассматриваются пользовательские части XML в контексте общего API JavaScript для Office, поддерживаемого в Word.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-191">This section discusses custom XML parts in the context of the Office Common JavaScript API which is supported in Word.</span></span> <span data-ttu-id="2b2f3-192">Специальный API JavaScript для Excel также предоставляет доступ к пользовательским частям XML.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-192">The host-specific Excel JavaScript API also provides access to the custom XML parts.</span></span> <span data-ttu-id="2b2f3-193">Интерфейсы API Excel и шаблоны программирования слегка отличаются.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-193">The Excel APIs and programming patterns are somewhat different.</span></span> <span data-ttu-id="2b2f3-194">Дополнительные сведения см. в статье [Excel CustomXmlPart](https://docs.microsoft.com/javascript/api/excel/excel.customxmlpart?view=office-js).</span><span class="sxs-lookup"><span data-stu-id="2b2f3-194">For more information, see [Excel CustomXmlPart](https://docs.microsoft.com/javascript/api/excel/excel.customxmlpart?view=office-js).</span></span>

<span data-ttu-id="2b2f3-p124">Существует параметр дополнительного хранилища, когда вам необходимо сохранить сведения, превышающие размер ограничения параметров документа, или содержащие структурированный символ. Настраиваемую XML-разметку задач области надстройки можно сохранить для Word (а для Excel, см. примечание в верхней части этого раздела). В Word, используйте объект [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) и его методы (еще раз, см. примечание выше для Excel). Следующий код создает настраиваемую пользовательскую часть XML и отображает её идентификатор, а затем содержимое элементов DIV на странице. Обратите внимание, что должен быть атрибут`xmlns` в строке XML.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p124">There is an addtional storage option when you need to store information that exceeds the size limits of the document Settings or which has a structured character. You can persist custom XML markup in a task pane add-in for Word (and for Excel, but see the note at the top of this section). In Word, you use the [CustomXmlPart](https://docs.microsoft.com/javascript/api/office/office.customxmlpart?view=office-js) object and its methods (Again, see the note above for Excel.) The following code creates a custom XML part and displays its ID and then its content in divs on the page. Note that there must be an `xmlns` attribute in the XML string.</span></span>

```js
function createCustomXmlPart() {
    const xmlString = "<Reviewers xmlns='http://schemas.contoso.com/review/1.0'><Reviewer>Juan</Reviewer><Reviewer>Hong</Reviewer><Reviewer>Sally</Reviewer></Reviewers>";
    Office.context.document.customXmlParts.addAsync(xmlString,
        (asyncResult) => {
            $("#xml-id").text("Your new XML part's ID: " + asyncResult.id);
            asyncResult.value.getXmlAsync(
                (asyncResult) => {
                    $("#xml-blob").text(asyncResult.value);                    
                }
            );
        }
    );
}
```

<span data-ttu-id="2b2f3-200">Чтобы получить пользовательскую часть XML, используйте метод [getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#getbyidasync-id--options--callback-). Однако ИД — это GUID, генерируемый при создании части XML, поэтому его невозможно узнать во время написания кода.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-200">To retrieve a custom XML part, you use the [getByIdAsync](https://docs.microsoft.com/javascript/api/office/office.customxmlparts?view=office-js#getbyidasync-id--options--callback-) method, but the ID is a GUID that is generated when the XML part is created, so you can't know when coding what the ID is.</span></span> <span data-ttu-id="2b2f3-201">По этой причине при создании части XML рекомендуется сразу сохранить ее ИД в виде параметра с запоминающимся идентификатором.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-201">For that reason, it is a good practice when creating an XML part to immediately store the ID of the XML part as a setting and give it a memorable key.</span></span> <span data-ttu-id="2b2f3-202">Ниже показано, как это сделать.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-202">The following method shows how to do this.</span></span> <span data-ttu-id="2b2f3-203">В предыдущих разделах этой статьи вы найдете подробные сведения и рекомендации по работе с настраиваемыми параметрами.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-203">(But see earlier sections of this article for details and best practices when working with custom settings).</span></span>

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

<span data-ttu-id="2b2f3-204">В приведенном ниже коде показано, как получить часть XML, сначала получив ее ИД из параметра.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-204">The following code shows how to retrieve the XML part by first getting its ID from a setting.</span></span>

 ```js
function getReviewers() {
    const reviewersXmlId = Office.context.document.settings.get('ReviewersID'));
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


## <a name="how-to-save-settings-in-the-users-mailbox-for-outlook-add-ins-as-roaming-settings"></a><span data-ttu-id="2b2f3-205">Сохранение параметров в почтовом ящике пользователя для надстроек Outlook в качестве параметров перемещения</span><span class="sxs-lookup"><span data-stu-id="2b2f3-205">How to save settings in the user's mailbox for Outlook add-ins as roaming settings</span></span>


<span data-ttu-id="2b2f3-206">Надстройка Outlook может использовать [объект RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) для сохранения данных состояния надстройки и данных настроек, характерных для почтового ящика пользователя.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-206">An Outlook add-in can use the [RoamingSettings](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) object to save add-in state and settings data that is specific to the user's mailbox.</span></span> <span data-ttu-id="2b2f3-207">Эти данные доступны только этой надстройке Outlook от имени пользователя, выполняющего надстройку.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-207">This data is accessible only by that Outlook add-in on behalf of the user running the add-in.</span></span> <span data-ttu-id="2b2f3-208">Данные хранятся в почтовом ящике сервера Exchange пользователя и доступны, когда этот пользователь входит в свою учетную запись и запускает надстройку Outlook.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-208">The data is stored on the user's Exchange Server mailbox, and is accessible when that user logs into their account and runs the Outlook add-in.</span></span>


### <a name="loading-roaming-settings"></a><span data-ttu-id="2b2f3-209">Загрузка параметров перемещения</span><span class="sxs-lookup"><span data-stu-id="2b2f3-209">Loading roaming settings</span></span>


<span data-ttu-id="2b2f3-p127">Надстройка Outlook обычно загружает параметры перемещения в обработчик событий [Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js). В следующем примере кода JavaScript показано, как выполняется загрузка существующих параметров перемещения.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p127">An Outlook add-in typically loads roaming settings in the [Office.initialize](https://docs.microsoft.com/javascript/api/office?view=office-js) event handler. The following JavaScript code example shows how to load existing roaming settings.</span></span>


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


### <a name="creating-or-assigning-a-roaming-setting"></a><span data-ttu-id="2b2f3-212">Создание или назначение параметра перемещения</span><span class="sxs-lookup"><span data-stu-id="2b2f3-212">Creating or assigning a roaming setting</span></span>


<span data-ttu-id="2b2f3-p128">Развивая предыдущий пример, следующая функция  `setAppSetting`, показывает, как использовать метод [RoamingSettings.set](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#set-name--value-) для определения или обновления заданного параметра `cookie` с указанием сегодняшнего числа. Затем он позволяет заново сохранить все параметры перемещения на сервере Exchange при помощи метода [RoamingSettings.saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#saveasync-callback-).</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p128">Continuing with the preceding example, the following  `setAppSetting` function shows how to use the [RoamingSettings.set](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#set-name--value-) method to set or update a setting named `cookie` with today's date. Then, it saves all the roaming settings back to the Exchange Server with the [RoamingSettings.saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#saveasync-callback-) method.</span></span>


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

<span data-ttu-id="2b2f3-215">Метод **saveAsync** сохраняет асинхронно параметры роуминга и использует  необязательную функцию обратного вызова.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-215">The **saveAsync** method saves roaming settings asynchronously and takes an optional callback function.</span></span> <span data-ttu-id="2b2f3-216">Этот пример кода передает функцию обратного вызова с именем `saveMyAppSettingsCallback` метода **saveAsync**.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-216">This code sample passes a callback function named `saveMyAppSettingsCallback` to the **saveAsync** method.</span></span> <span data-ttu-id="2b2f3-217">Когда параметр асинхронного вызова_asyncResult_ `saveMyAppSettingsCallback` возвращается, функция предоставляет доступ к объекту [AsyncResult](https://docs.microsoft.com/javascript/api/outlook?view=office-js), который можно использовать для определения успешного или неудачного выполнения операции свойства **AsyncResult.status**</span><span class="sxs-lookup"><span data-stu-id="2b2f3-217">The  saveAsync method saves roaming settings asynchronously and takes an optional callback function. This code sample passes a callback function named  to the saveAsync method. When the asynchronous call returns, the _asyncResult_ parameter of the `saveMyAppSettingsCallback` function provides access to an [AsyncResult](https://docs.microsoft.com/javascript/api/outlook?view=office-js) object that you can use to determine the success or failure of the operation with the **AsyncResult.status** property.</span></span>


### <a name="removing-a-roaming-setting"></a><span data-ttu-id="2b2f3-218">Удаление параметра перемещения</span><span class="sxs-lookup"><span data-stu-id="2b2f3-218">Removing a roaming setting</span></span>


<span data-ttu-id="2b2f3-219">Предыдущие примеры дополняет следующая функция  `removeAppSetting`, демонстрирующая применение метода [RoamingSettings.remove](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#remove-name-) для удаления параметра `cookie` и повторного сохранения всех параметров перемещения на сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-219">Also extending the preceding examples, the following  `removeAppSetting` function, shows how to use the [RoamingSettings.remove](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js#remove-name-) method to remove the `cookie` setting and save all the roaming settings back to the Exchange Server.</span></span>


```js
// Remove an application setting.
function removeAppSetting()
{
    _settings.remove("cookie");
    _settings.saveAsync(saveMyAppSettingsCallback);
}
```


## <a name="how-to-save-settings-per-item-for-outlook-add-ins-as-custom-properties"></a><span data-ttu-id="2b2f3-220">Сохранение параметров для каждого элемента надстройки Outlook в качестве пользовательских свойств</span><span class="sxs-lookup"><span data-stu-id="2b2f3-220">How to save settings per item for Outlook add-ins as custom properties</span></span>


<span data-ttu-id="2b2f3-p130">Пользовательские свойства позволяют надстройке Outlook сохранять сведения об элементе, который она использует. Например, если в надстройке Outlook создается встреча на основе приглашения на собрание в сообщении, с помощью пользовательских свойств можно сохранить сведения о факте создания собрания. Это гарантирует, что надстройка не предложит создать встречу еще раз при повторном открытии сообщения.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p130">Custom properties let your Outlook add-in store information about an item it is working with. For example, if your Outlook add-in creates an appointment from a meeting suggestion in a message, you can use custom properties to store the fact that the meeting was created. This makes sure that if the message is opened again, your Outlook add-in doesn't offer to create the appointment again.</span></span>

<span data-ttu-id="2b2f3-p131">Перед использованием пользовательских свойств для определенного сообщения, встречи или элемента приглашения на собрание, необходимо загрузить свойства в память путем вызова метода [loadCustomPropertiesAsync](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) объекта **Item**. Если какие-либо пользовательские свойства уже заданы для текущего элемента, на этом этапе они загружаются с сервера Exchange. После загрузки свойств можно использовать методы [set](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#set-name--value-) и [get](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) объекта **CustomProperties** для добавления, обновления и получения свойств в памяти. Чтобы сохранить любые изменения, внесенные в пользовательские свойства элемента, необходимо использовать метод [saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#saveasync-callback--asynccontext-) для сохранения изменений в элементе на сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p131">Before you can use custom properties for a particular message, appointment, or meeting request item, you must load the properties into memory by calling the [loadCustomPropertiesAsync](https://docs.microsoft.com/javascript/api/outlook/office.mailbox?view=office-js) method of the **Item** object. If any custom properties are already set for the current item, they are loaded from the Exchange server at this point. After you have loaded the properties, you can use the [set](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#set-name--value-) and [get](https://docs.microsoft.com/javascript/api/outlook/office.roamingsettings?view=office-js) methods of the **CustomProperties** object to add, update, and retrieve properties in memory. To save any changes that you make to the item's custom properties, you must use the [saveAsync](https://docs.microsoft.com/javascript/api/outlook/office.customproperties?view=office-js#saveasync-callback--asynccontext-) method to persist the changes to the item on the Exchange server.</span></span>


### <a name="custom-properties-example"></a><span data-ttu-id="2b2f3-228">Пример пользовательских свойств</span><span class="sxs-lookup"><span data-stu-id="2b2f3-228">Custom properties example</span></span>

<span data-ttu-id="2b2f3-p132">В следующем примере демонстрируется упрощенный набор функций для надстройки Outlook, применяющей пользовательские свойства. Этот пример можно использовать в качестве отправной точки для работы с такой надстройкой Outlook.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-p132">The following example shows a simplified set of functions for an Outlook add-in that uses custom properties. You can use this example as a starting point for your Outlook add-in that uses custom properties.</span></span> 

<span data-ttu-id="2b2f3-231">Надстройка Outlook, использующая эти функции, получает любые пользовательские свойства, вызывая метод**get** для переменной `_customProps`, как показано в приведенном ниже примере.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-231">An Outlook add-in that uses these functions retrieves any custom properties by calling the  **get** method on the `_customProps` variable, as shown in the following example.</span></span>




```js
var property = _customProps.get("propertyName");
```

<span data-ttu-id="2b2f3-232">Этот пример включает следующие функции:</span><span class="sxs-lookup"><span data-stu-id="2b2f3-232">This example includes the following functions:</span></span>



|<span data-ttu-id="2b2f3-233">**Имя функции**</span><span class="sxs-lookup"><span data-stu-id="2b2f3-233">**Function name**</span></span>|<span data-ttu-id="2b2f3-234">**Описание**</span><span class="sxs-lookup"><span data-stu-id="2b2f3-234">**Description**</span></span>|
|:-----|:-----|
| `Office.initialize`|<span data-ttu-id="2b2f3-235">Инициализирует надстройку и загружает пользовательские свойства текущего элемента с сервера Exchange.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-235">Initializes the add-in and loads the custom properties for the current item from the Exchange server.</span></span>|
| `customPropsCallback`|<span data-ttu-id="2b2f3-236">Получает пользовательские свойства, возвращенные сервером Exchange, и сохраняет их для дальнейшего использования.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-236">Gets the custom properties that are returned from the Exchange server and saves it for later use.</span></span>|
| `updateProperty`|<span data-ttu-id="2b2f3-237">Задает или обновляет определенное свойство, а затем сохраняет изменение на сервер Exchange.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-237">Sets or updates a specific property, and then saves the change to the Exchange server.</span></span>|
| `removeProperty`|<span data-ttu-id="2b2f3-238">Удаляет определенное свойство и сохраняет факт удаления на сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-238">Removes a specific property, and then persists the removal to the Exchange server.</span></span>|
| `saveCallback`|<span data-ttu-id="2b2f3-239">Обратный вызов метода **saveAsync** в функциях `updateProperty` и `removeProperty`.</span><span class="sxs-lookup"><span data-stu-id="2b2f3-239">Callback for calls to the  **saveAsync** method in the `updateProperty` and `removeProperty` functions.</span></span>|



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


## <a name="see-also"></a><span data-ttu-id="2b2f3-240">См. также</span><span class="sxs-lookup"><span data-stu-id="2b2f3-240">See also</span></span>

- [<span data-ttu-id="2b2f3-241">Общие сведения об интерфейсе API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="2b2f3-241">Understanding the JavaScript API for Office</span></span>](understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="2b2f3-242">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="2b2f3-242">Outlook add-ins</span></span>](https://docs.microsoft.com/outlook/add-ins/)
- [<span data-ttu-id="2b2f3-243">Excel-Add-in-JavaScript-PersistCustomSettings</span><span class="sxs-lookup"><span data-stu-id="2b2f3-243">Excel-Add-in-JavaScript-PersistCustomSettings</span></span>](https://github.com/OfficeDev/Excel-Add-in-JavaScript-PersistCustomSettings)
    
