---
title: Office. Context — набор обязательных элементов 1,3
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: c71be3f16ca637e9c1cb2790cda2a34956f6c67a
ms.sourcegitcommit: 3f5d7f4794e3d3c8bc3a79fa05c54157613b9376
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/02/2019
ms.locfileid: "36064727"
---
# <a name="context"></a><span data-ttu-id="bdd5c-102">context</span><span class="sxs-lookup"><span data-stu-id="bdd5c-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="bdd5c-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="bdd5c-103">[Office](Office.md).context</span></span>

<span data-ttu-id="bdd5c-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="bdd5c-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="bdd5c-106">Требования</span><span class="sxs-lookup"><span data-stu-id="bdd5c-106">Requirements</span></span>

|<span data-ttu-id="bdd5c-107">Требование</span><span class="sxs-lookup"><span data-stu-id="bdd5c-107">Requirement</span></span>| <span data-ttu-id="bdd5c-108">Значение</span><span class="sxs-lookup"><span data-stu-id="bdd5c-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="bdd5c-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bdd5c-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bdd5c-110">1.0</span><span class="sxs-lookup"><span data-stu-id="bdd5c-110">1.0</span></span>|
|[<span data-ttu-id="bdd5c-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bdd5c-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bdd5c-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bdd5c-112">Compose or Read</span></span>|

### <a name="namespaces"></a><span data-ttu-id="bdd5c-113">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="bdd5c-113">Namespaces</span></span>

<span data-ttu-id="bdd5c-114">[почтовый ящик](office.context.mailbox.md): предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="bdd5c-114">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="bdd5c-115">Members</span><span class="sxs-lookup"><span data-stu-id="bdd5c-115">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="bdd5c-116">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="bdd5c-116">displayLanguage: String</span></span>

<span data-ttu-id="bdd5c-117">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="bdd5c-117">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="bdd5c-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="bdd5c-118">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="bdd5c-119">Тип</span><span class="sxs-lookup"><span data-stu-id="bdd5c-119">Type</span></span>

*   <span data-ttu-id="bdd5c-120">String</span><span class="sxs-lookup"><span data-stu-id="bdd5c-120">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bdd5c-121">Требования</span><span class="sxs-lookup"><span data-stu-id="bdd5c-121">Requirements</span></span>

|<span data-ttu-id="bdd5c-122">Требование</span><span class="sxs-lookup"><span data-stu-id="bdd5c-122">Requirement</span></span>| <span data-ttu-id="bdd5c-123">Значение</span><span class="sxs-lookup"><span data-stu-id="bdd5c-123">Value</span></span>|
|---|---|
|[<span data-ttu-id="bdd5c-124">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bdd5c-124">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bdd5c-125">1.0</span><span class="sxs-lookup"><span data-stu-id="bdd5c-125">1.0</span></span>|
|[<span data-ttu-id="bdd5c-126">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bdd5c-126">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bdd5c-127">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bdd5c-127">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bdd5c-128">Пример</span><span class="sxs-lookup"><span data-stu-id="bdd5c-128">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-13"></a><span data-ttu-id="bdd5c-129">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.3)</span><span class="sxs-lookup"><span data-stu-id="bdd5c-129">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.3)</span></span>

<span data-ttu-id="bdd5c-130">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="bdd5c-130">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="bdd5c-131">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="bdd5c-131">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="bdd5c-132">Тип</span><span class="sxs-lookup"><span data-stu-id="bdd5c-132">Type</span></span>

*   [<span data-ttu-id="bdd5c-133">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="bdd5c-133">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.3)

##### <a name="requirements"></a><span data-ttu-id="bdd5c-134">Требования</span><span class="sxs-lookup"><span data-stu-id="bdd5c-134">Requirements</span></span>

|<span data-ttu-id="bdd5c-135">Требование</span><span class="sxs-lookup"><span data-stu-id="bdd5c-135">Requirement</span></span>| <span data-ttu-id="bdd5c-136">Значение</span><span class="sxs-lookup"><span data-stu-id="bdd5c-136">Value</span></span>|
|---|---|
|[<span data-ttu-id="bdd5c-137">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bdd5c-137">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bdd5c-138">1.0</span><span class="sxs-lookup"><span data-stu-id="bdd5c-138">1.0</span></span>|
|[<span data-ttu-id="bdd5c-139">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bdd5c-139">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bdd5c-140">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="bdd5c-140">Restricted</span></span>|
|[<span data-ttu-id="bdd5c-141">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bdd5c-141">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bdd5c-142">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bdd5c-142">Compose or Read</span></span>|
