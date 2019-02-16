---
title: Office.context — набор обязательных элементов 1.2
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 8e8d83500e1cb9de730f6a7bac75c3bcc0012b4e
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30068135"
---
# <a name="context"></a><span data-ttu-id="4eacd-102">context</span><span class="sxs-lookup"><span data-stu-id="4eacd-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="4eacd-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="4eacd-103">[Office](Office.md).context</span></span>

<span data-ttu-id="4eacd-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="4eacd-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="4eacd-106">Требования</span><span class="sxs-lookup"><span data-stu-id="4eacd-106">Requirements</span></span>

|<span data-ttu-id="4eacd-107">Требование</span><span class="sxs-lookup"><span data-stu-id="4eacd-107">Requirement</span></span>| <span data-ttu-id="4eacd-108">Значение</span><span class="sxs-lookup"><span data-stu-id="4eacd-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="4eacd-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4eacd-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4eacd-110">1.0</span><span class="sxs-lookup"><span data-stu-id="4eacd-110">1.0</span></span>|
|[<span data-ttu-id="4eacd-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4eacd-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4eacd-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4eacd-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="4eacd-113">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="4eacd-113">Namespaces</span></span>

<span data-ttu-id="4eacd-114">[mailbox](office.context.mailbox.md). Предоставляет для Microsoft Outlook и Microsoft Outlook в Интернете доступ к объектной модели надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="4eacd-114">[mailbox](office.context.mailbox.md) - Provides access to the Outlook add-in object model for Microsoft Outlook and Microsoft Outlook on the web.</span></span>

### <a name="members"></a><span data-ttu-id="4eacd-115">Элементы</span><span class="sxs-lookup"><span data-stu-id="4eacd-115">Members</span></span>

####  <a name="displaylanguage-string"></a><span data-ttu-id="4eacd-116">displayLanguage :String</span><span class="sxs-lookup"><span data-stu-id="4eacd-116">displayLanguage :String</span></span>

<span data-ttu-id="4eacd-117">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="4eacd-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="4eacd-118">Значение `displayLanguage` отображает текущий параметр **Язык интерфейса**, заданный в разделе **Файл > Параметры > Язык** ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="4eacd-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="4eacd-119">Тип</span><span class="sxs-lookup"><span data-stu-id="4eacd-119">Type</span></span>

*   <span data-ttu-id="4eacd-120">String</span><span class="sxs-lookup"><span data-stu-id="4eacd-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4eacd-121">Требования</span><span class="sxs-lookup"><span data-stu-id="4eacd-121">Requirements</span></span>

|<span data-ttu-id="4eacd-122">Требование</span><span class="sxs-lookup"><span data-stu-id="4eacd-122">Requirement</span></span>| <span data-ttu-id="4eacd-123">Значение</span><span class="sxs-lookup"><span data-stu-id="4eacd-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="4eacd-124">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4eacd-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4eacd-125">1.0</span><span class="sxs-lookup"><span data-stu-id="4eacd-125">1.0</span></span>|
|[<span data-ttu-id="4eacd-126">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4eacd-126">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4eacd-127">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4eacd-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="4eacd-128">Пример</span><span class="sxs-lookup"><span data-stu-id="4eacd-128">Example</span></span>

```javascript
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

####  <a name="roamingsettings-roamingsettingsjavascriptapioutlook12officeroamingsettings"></a><span data-ttu-id="4eacd-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span><span class="sxs-lookup"><span data-stu-id="4eacd-129">roamingSettings :[RoamingSettings](/javascript/api/outlook_1_2/office.RoamingSettings)</span></span>

<span data-ttu-id="4eacd-130">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="4eacd-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="4eacd-131">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="4eacd-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="4eacd-132">Тип</span><span class="sxs-lookup"><span data-stu-id="4eacd-132">Type</span></span>

*   [<span data-ttu-id="4eacd-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="4eacd-133">RoamingSettings</span></span>](/javascript/api/outlook_1_2/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="4eacd-134">Требования</span><span class="sxs-lookup"><span data-stu-id="4eacd-134">Requirements</span></span>

|<span data-ttu-id="4eacd-135">Требование</span><span class="sxs-lookup"><span data-stu-id="4eacd-135">Requirement</span></span>| <span data-ttu-id="4eacd-136">Значение</span><span class="sxs-lookup"><span data-stu-id="4eacd-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="4eacd-137">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4eacd-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4eacd-138">1.0</span><span class="sxs-lookup"><span data-stu-id="4eacd-138">1.0</span></span>|
|[<span data-ttu-id="4eacd-139">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4eacd-139">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4eacd-140">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="4eacd-140">Restricted</span></span>|
|[<span data-ttu-id="4eacd-141">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4eacd-141">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4eacd-142">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4eacd-142">Compose or Read</span></span>|
