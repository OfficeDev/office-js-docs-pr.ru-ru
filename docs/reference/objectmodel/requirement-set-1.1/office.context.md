---
title: Office. Context — набор обязательных элементов 1,1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: aa8a30f60f1ba8848d7d7d342635286933a14a89
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696388"
---
# <a name="context"></a><span data-ttu-id="d0941-102">context</span><span class="sxs-lookup"><span data-stu-id="d0941-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="d0941-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="d0941-103">[Office](Office.md).context</span></span>

<span data-ttu-id="d0941-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="d0941-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="d0941-106">Требования</span><span class="sxs-lookup"><span data-stu-id="d0941-106">Requirements</span></span>

|<span data-ttu-id="d0941-107">Требование</span><span class="sxs-lookup"><span data-stu-id="d0941-107">Requirement</span></span>| <span data-ttu-id="d0941-108">Значение</span><span class="sxs-lookup"><span data-stu-id="d0941-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0941-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d0941-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0941-110">1.0</span><span class="sxs-lookup"><span data-stu-id="d0941-110">1.0</span></span>|
|[<span data-ttu-id="d0941-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0941-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d0941-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d0941-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d0941-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="d0941-113">Members and methods</span></span>

| <span data-ttu-id="d0941-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="d0941-114">Member</span></span> | <span data-ttu-id="d0941-115">Тип</span><span class="sxs-lookup"><span data-stu-id="d0941-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d0941-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="d0941-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="d0941-117">Member</span><span class="sxs-lookup"><span data-stu-id="d0941-117">Member</span></span> |
| [<span data-ttu-id="d0941-118">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="d0941-118">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="d0941-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="d0941-119">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="d0941-120">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="d0941-120">Namespaces</span></span>

<span data-ttu-id="d0941-121">[почтовый ящик](office.context.mailbox.md): предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="d0941-121">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="d0941-122">Members</span><span class="sxs-lookup"><span data-stu-id="d0941-122">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="d0941-123">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="d0941-123">displayLanguage: String</span></span>

<span data-ttu-id="d0941-124">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="d0941-124">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="d0941-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="d0941-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="d0941-126">Тип</span><span class="sxs-lookup"><span data-stu-id="d0941-126">Type</span></span>

*   <span data-ttu-id="d0941-127">String</span><span class="sxs-lookup"><span data-stu-id="d0941-127">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d0941-128">Требования</span><span class="sxs-lookup"><span data-stu-id="d0941-128">Requirements</span></span>

|<span data-ttu-id="d0941-129">Требование</span><span class="sxs-lookup"><span data-stu-id="d0941-129">Requirement</span></span>| <span data-ttu-id="d0941-130">Значение</span><span class="sxs-lookup"><span data-stu-id="d0941-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0941-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d0941-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0941-132">1.0</span><span class="sxs-lookup"><span data-stu-id="d0941-132">1.0</span></span>|
|[<span data-ttu-id="d0941-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0941-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d0941-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d0941-134">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d0941-135">Пример</span><span class="sxs-lookup"><span data-stu-id="d0941-135">Example</span></span>

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

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-11"></a><span data-ttu-id="d0941-136">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.1)</span><span class="sxs-lookup"><span data-stu-id="d0941-136">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.1)</span></span>

<span data-ttu-id="d0941-137">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="d0941-137">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="d0941-138">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="d0941-138">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="d0941-139">Тип</span><span class="sxs-lookup"><span data-stu-id="d0941-139">Type</span></span>

*   [<span data-ttu-id="d0941-140">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d0941-140">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.1)

##### <a name="requirements"></a><span data-ttu-id="d0941-141">Требования</span><span class="sxs-lookup"><span data-stu-id="d0941-141">Requirements</span></span>

|<span data-ttu-id="d0941-142">Требование</span><span class="sxs-lookup"><span data-stu-id="d0941-142">Requirement</span></span>| <span data-ttu-id="d0941-143">Значение</span><span class="sxs-lookup"><span data-stu-id="d0941-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="d0941-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d0941-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d0941-145">1.0</span><span class="sxs-lookup"><span data-stu-id="d0941-145">1.0</span></span>|
|[<span data-ttu-id="d0941-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d0941-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d0941-147">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="d0941-147">Restricted</span></span>|
|[<span data-ttu-id="d0941-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d0941-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d0941-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d0941-149">Compose or Read</span></span>|
