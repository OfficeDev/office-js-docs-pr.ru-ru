---
title: Office.context — набор обязательных элементов 1.1
description: ''
ms.date: 10/11/2018
ms.openlocfilehash: 392e54f1004bb395672c026ef749113f94ec7479
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432727"
---
# <a name="context"></a><span data-ttu-id="2453d-102">context</span><span class="sxs-lookup"><span data-stu-id="2453d-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="2453d-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="2453d-103">[Office](Office.md).context</span></span>

<span data-ttu-id="2453d-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="2453d-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Shared API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="2453d-106">Требования</span><span class="sxs-lookup"><span data-stu-id="2453d-106">Requirements</span></span>

|<span data-ttu-id="2453d-107">Requirement</span><span class="sxs-lookup"><span data-stu-id="2453d-107">Requirement</span></span>| <span data-ttu-id="2453d-108">Значение</span><span class="sxs-lookup"><span data-stu-id="2453d-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2453d-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2453d-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2453d-110">1.0</span><span class="sxs-lookup"><span data-stu-id="2453d-110">1.0</span></span>|
|[<span data-ttu-id="2453d-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2453d-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2453d-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2453d-112">Compose or read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="2453d-113">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="2453d-113">Namespaces</span></span>

<span data-ttu-id="2453d-114">[mailbox](office.context.mailbox.md). Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="2453d-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook Add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="2453d-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="2453d-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="2453d-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="2453d-116">displayLanguage :String</span></span>

<span data-ttu-id="2453d-117">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="2453d-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="2453d-118">Значение `displayLanguage` отображает текущий параметр **Язык интерфейса**, заданный в разделе **Файл > Параметры > Язык** ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="2453d-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="2453d-119">Тип:</span><span class="sxs-lookup"><span data-stu-id="2453d-119">Type:</span></span>

*   <span data-ttu-id="2453d-120">String</span><span class="sxs-lookup"><span data-stu-id="2453d-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2453d-121">Требования</span><span class="sxs-lookup"><span data-stu-id="2453d-121">Requirements</span></span>

|<span data-ttu-id="2453d-122">Requirement</span><span class="sxs-lookup"><span data-stu-id="2453d-122">Requirement</span></span>| <span data-ttu-id="2453d-123">Значение</span><span class="sxs-lookup"><span data-stu-id="2453d-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="2453d-124">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2453d-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2453d-125">1.0</span><span class="sxs-lookup"><span data-stu-id="2453d-125">1.0</span></span>|
|[<span data-ttu-id="2453d-126">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2453d-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2453d-127">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2453d-127">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2453d-128">Пример</span><span class="sxs-lookup"><span data-stu-id="2453d-128">Example</span></span>

```js
function sayHelloWithDisplayLanguage() {
  var myDisplayLanguage = Office.context.displayLanguage;
  switch (myDisplayLanguage) {
    case 'en-US':
      write('Hello!');
      break;
    case 'en-NZ':
      write('G\'day mate!');
      break;
  }
}
// Function that writes to a div with id='message' on the page.
function write(message){
  document.getElementById('message').innerText += message;
}
```

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook11officeroamingsettings"></a><span data-ttu-id="2453d-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="2453d-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_1/office.RoamingSettings)</span></span>

<span data-ttu-id="2453d-130">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="2453d-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="2453d-131">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="2453d-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="2453d-132">Тип:</span><span class="sxs-lookup"><span data-stu-id="2453d-132">Type:</span></span>

*   [<span data-ttu-id="2453d-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="2453d-133">RoamingSettings</span></span>](/javascript/api/outlook_1_1/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="2453d-134">Требования</span><span class="sxs-lookup"><span data-stu-id="2453d-134">Requirements</span></span>

|<span data-ttu-id="2453d-135">Requirement</span><span class="sxs-lookup"><span data-stu-id="2453d-135">Requirement</span></span>| <span data-ttu-id="2453d-136">Значение</span><span class="sxs-lookup"><span data-stu-id="2453d-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="2453d-137">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2453d-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2453d-138">1.0</span><span class="sxs-lookup"><span data-stu-id="2453d-138">1.0</span></span>|
|[<span data-ttu-id="2453d-139">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2453d-139">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2453d-140">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="2453d-140">Restricted</span></span>|
|[<span data-ttu-id="2453d-141">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2453d-141">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2453d-142">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2453d-142">Compose or read</span></span>|