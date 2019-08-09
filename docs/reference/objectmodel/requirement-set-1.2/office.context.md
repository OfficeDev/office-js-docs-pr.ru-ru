---
title: Office. Context — набор обязательных элементов 1,2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: a7621c3e9e29d229f66ee950119770cbc31d9fe3
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268497"
---
# <a name="context"></a><span data-ttu-id="df149-102">context</span><span class="sxs-lookup"><span data-stu-id="df149-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="df149-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="df149-103">[Office](Office.md).context</span></span>

<span data-ttu-id="df149-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="df149-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>


##### <a name="requirements"></a><span data-ttu-id="df149-106">Требования</span><span class="sxs-lookup"><span data-stu-id="df149-106">Requirements</span></span>

|<span data-ttu-id="df149-107">Требование</span><span class="sxs-lookup"><span data-stu-id="df149-107">Requirement</span></span>| <span data-ttu-id="df149-108">Значение</span><span class="sxs-lookup"><span data-stu-id="df149-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="df149-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df149-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df149-110">1.0</span><span class="sxs-lookup"><span data-stu-id="df149-110">1.0</span></span>|
|[<span data-ttu-id="df149-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df149-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="df149-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df149-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="df149-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="df149-113">Members and methods</span></span>

| <span data-ttu-id="df149-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="df149-114">Member</span></span> | <span data-ttu-id="df149-115">Тип</span><span class="sxs-lookup"><span data-stu-id="df149-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="df149-116">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="df149-116">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="df149-117">Member</span><span class="sxs-lookup"><span data-stu-id="df149-117">Member</span></span> |
| [<span data-ttu-id="df149-118">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="df149-118">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="df149-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="df149-119">Member</span></span> |

### <a name="namespaces"></a><span data-ttu-id="df149-120">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="df149-120">Namespaces</span></span>

<span data-ttu-id="df149-121">[Mailbox](office.context.mailbox.md) — предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="df149-121">[mailbox](office.context.mailbox.md) - Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

### <a name="members"></a><span data-ttu-id="df149-122">Members</span><span class="sxs-lookup"><span data-stu-id="df149-122">Members</span></span>

#### <a name="displaylanguage-string"></a><span data-ttu-id="df149-123">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="df149-123">displayLanguage: String</span></span>

<span data-ttu-id="df149-124">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="df149-124">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="df149-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="df149-125">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="df149-126">Тип</span><span class="sxs-lookup"><span data-stu-id="df149-126">Type</span></span>

*   <span data-ttu-id="df149-127">String</span><span class="sxs-lookup"><span data-stu-id="df149-127">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="df149-128">Требования</span><span class="sxs-lookup"><span data-stu-id="df149-128">Requirements</span></span>

|<span data-ttu-id="df149-129">Требование</span><span class="sxs-lookup"><span data-stu-id="df149-129">Requirement</span></span>| <span data-ttu-id="df149-130">Значение</span><span class="sxs-lookup"><span data-stu-id="df149-130">Value</span></span>|
|---|---|
|[<span data-ttu-id="df149-131">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df149-131">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df149-132">1.0</span><span class="sxs-lookup"><span data-stu-id="df149-132">1.0</span></span>|
|[<span data-ttu-id="df149-133">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df149-133">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="df149-134">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df149-134">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="df149-135">Пример</span><span class="sxs-lookup"><span data-stu-id="df149-135">Example</span></span>

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

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettingsviewoutlook-js-12"></a><span data-ttu-id="df149-136">roamingSettings: [roamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.2)</span><span class="sxs-lookup"><span data-stu-id="df149-136">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.2)</span></span>

<span data-ttu-id="df149-137">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="df149-137">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="df149-138">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="df149-138">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="df149-139">Тип</span><span class="sxs-lookup"><span data-stu-id="df149-139">Type</span></span>

*   [<span data-ttu-id="df149-140">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="df149-140">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings?view=outlook-js-1.2)

##### <a name="requirements"></a><span data-ttu-id="df149-141">Требования</span><span class="sxs-lookup"><span data-stu-id="df149-141">Requirements</span></span>

|<span data-ttu-id="df149-142">Требование</span><span class="sxs-lookup"><span data-stu-id="df149-142">Requirement</span></span>| <span data-ttu-id="df149-143">Значение</span><span class="sxs-lookup"><span data-stu-id="df149-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="df149-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="df149-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="df149-145">1.0</span><span class="sxs-lookup"><span data-stu-id="df149-145">1.0</span></span>|
|[<span data-ttu-id="df149-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="df149-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="df149-147">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="df149-147">Restricted</span></span>|
|[<span data-ttu-id="df149-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="df149-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="df149-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="df149-149">Compose or Read</span></span>|
