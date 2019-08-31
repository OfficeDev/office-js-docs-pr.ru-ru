---
title: Office. Context — набор обязательных элементов 1,6
description: ''
ms.date: 06/25/2019
localization_priority: Normal
ms.openlocfilehash: 10453801f023ee928e9d5f4fcff3fc22f8cd0319
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695996"
---
# <a name="context"></a><span data-ttu-id="2d08b-102">context</span><span class="sxs-lookup"><span data-stu-id="2d08b-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="2d08b-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="2d08b-103">[Office](Office.md).context</span></span>

<span data-ttu-id="2d08b-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="2d08b-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d08b-106">Требования</span><span class="sxs-lookup"><span data-stu-id="2d08b-106">Requirements</span></span>

|<span data-ttu-id="2d08b-107">Требование</span><span class="sxs-lookup"><span data-stu-id="2d08b-107">Requirement</span></span>| <span data-ttu-id="2d08b-108">Значение</span><span class="sxs-lookup"><span data-stu-id="2d08b-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d08b-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2d08b-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d08b-110">1.0</span><span class="sxs-lookup"><span data-stu-id="2d08b-110">1.0</span></span>|
|[<span data-ttu-id="2d08b-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2d08b-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d08b-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2d08b-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2d08b-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="2d08b-113">Members and methods</span></span>

| <span data-ttu-id="2d08b-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="2d08b-114">Member</span></span> | <span data-ttu-id="2d08b-115">Тип</span><span class="sxs-lookup"><span data-stu-id="2d08b-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2d08b-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="2d08b-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="2d08b-117">Member</span><span class="sxs-lookup"><span data-stu-id="2d08b-117">Member</span></span> |
| [<span data-ttu-id="2d08b-118">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="2d08b-118">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="2d08b-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="2d08b-119">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="2d08b-120">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="2d08b-120">Namespaces</span></span>

<span data-ttu-id="2d08b-121">[почтовый ящик](office.context.mailbox.md): предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="2d08b-121">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="2d08b-122">Members</span><span class="sxs-lookup"><span data-stu-id="2d08b-122">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="2d08b-123">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="2d08b-123">displayLanguage: String</span></span>

<span data-ttu-id="2d08b-124">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="2d08b-124">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="2d08b-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="2d08b-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="2d08b-126">Тип</span><span class="sxs-lookup"><span data-stu-id="2d08b-126">Type</span></span>

*   <span data-ttu-id="2d08b-127">String</span><span class="sxs-lookup"><span data-stu-id="2d08b-127">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2d08b-128">Требования</span><span class="sxs-lookup"><span data-stu-id="2d08b-128">Requirements</span></span>

|<span data-ttu-id="2d08b-129">Требование</span><span class="sxs-lookup"><span data-stu-id="2d08b-129">Requirement</span></span>| <span data-ttu-id="2d08b-130">Значение</span><span class="sxs-lookup"><span data-stu-id="2d08b-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d08b-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2d08b-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d08b-132">1.0</span><span class="sxs-lookup"><span data-stu-id="2d08b-132">1.0</span></span>|
|[<span data-ttu-id="2d08b-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2d08b-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d08b-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2d08b-134">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="2d08b-135">Пример</span><span class="sxs-lookup"><span data-stu-id="2d08b-135">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-16"></a><span data-ttu-id="2d08b-136">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.6)</span><span class="sxs-lookup"><span data-stu-id="2d08b-136">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.6)</span></span>

<span data-ttu-id="2d08b-137">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="2d08b-137">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="2d08b-138">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="2d08b-138">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="2d08b-139">Тип</span><span class="sxs-lookup"><span data-stu-id="2d08b-139">Type</span></span>

*   [<span data-ttu-id="2d08b-140">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="2d08b-140">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.6)

##### <a name="requirements"></a><span data-ttu-id="2d08b-141">Требования</span><span class="sxs-lookup"><span data-stu-id="2d08b-141">Requirements</span></span>

|<span data-ttu-id="2d08b-142">Требование</span><span class="sxs-lookup"><span data-stu-id="2d08b-142">Requirement</span></span>| <span data-ttu-id="2d08b-143">Значение</span><span class="sxs-lookup"><span data-stu-id="2d08b-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="2d08b-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2d08b-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2d08b-145">1.0</span><span class="sxs-lookup"><span data-stu-id="2d08b-145">1.0</span></span>|
|[<span data-ttu-id="2d08b-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2d08b-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2d08b-147">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="2d08b-147">Restricted</span></span>|
|[<span data-ttu-id="2d08b-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2d08b-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="2d08b-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2d08b-149">Compose or Read</span></span>|
