---
title: Office.context — набор требований 1.3
description: Office. Участники объектов context, доступные для Outlook надстройки с помощью API почтовых ящиков, устанавливают 1.3.
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: c5c7f6eaa46bfa5067572878b03e47511310f894
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590395"
---
# <a name="context-mailbox-requirement-set-13"></a><span data-ttu-id="e489f-103">контекст (требования к почтовым ящикам 1.3)</span><span class="sxs-lookup"><span data-stu-id="e489f-103">context (Mailbox requirement set 1.3)</span></span>

### <a name="officecontext"></a><span data-ttu-id="e489f-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="e489f-104">[Office](office.md).context</span></span>

<span data-ttu-id="e489f-105">Office.context предоставляет общие интерфейсы, используемые надстройки во всех Office приложениях.</span><span class="sxs-lookup"><span data-stu-id="e489f-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="e489f-106">Этот список документов только те интерфейсы, которые используются Outlook надстройки. Полный список пространства имен Office.context см. в [ссылке Office.context в общем API.](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="e489f-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e489f-107">Требования</span><span class="sxs-lookup"><span data-stu-id="e489f-107">Requirements</span></span>

|<span data-ttu-id="e489f-108">Требование</span><span class="sxs-lookup"><span data-stu-id="e489f-108">Requirement</span></span>| <span data-ttu-id="e489f-109">Значение</span><span class="sxs-lookup"><span data-stu-id="e489f-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="e489f-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e489f-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e489f-111">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-111">1.1</span></span>|
|[<span data-ttu-id="e489f-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e489f-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e489f-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="e489f-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="e489f-114">Properties</span></span>

| <span data-ttu-id="e489f-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="e489f-115">Property</span></span> | <span data-ttu-id="e489f-116">Режимы</span><span class="sxs-lookup"><span data-stu-id="e489f-116">Modes</span></span> | <span data-ttu-id="e489f-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="e489f-117">Return type</span></span> | <span data-ttu-id="e489f-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="e489f-118">Minimum</span></span><br><span data-ttu-id="e489f-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="e489f-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e489f-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="e489f-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="e489f-121">Создание</span><span class="sxs-lookup"><span data-stu-id="e489f-121">Compose</span></span><br><span data-ttu-id="e489f-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-122">Read</span></span> | <span data-ttu-id="e489f-123">Строка</span><span class="sxs-lookup"><span data-stu-id="e489f-123">String</span></span> | [<span data-ttu-id="e489f-124">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e489f-125">диагностика</span><span class="sxs-lookup"><span data-stu-id="e489f-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="e489f-126">Создание</span><span class="sxs-lookup"><span data-stu-id="e489f-126">Compose</span></span><br><span data-ttu-id="e489f-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-127">Read</span></span> | [<span data-ttu-id="e489f-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="e489f-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="e489f-129">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e489f-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="e489f-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="e489f-131">Создание</span><span class="sxs-lookup"><span data-stu-id="e489f-131">Compose</span></span><br><span data-ttu-id="e489f-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-132">Read</span></span> | <span data-ttu-id="e489f-133">Строка</span><span class="sxs-lookup"><span data-stu-id="e489f-133">String</span></span> | [<span data-ttu-id="e489f-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e489f-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="e489f-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="e489f-136">Создание</span><span class="sxs-lookup"><span data-stu-id="e489f-136">Compose</span></span><br><span data-ttu-id="e489f-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-137">Read</span></span> | [<span data-ttu-id="e489f-138">Mailbox</span><span class="sxs-lookup"><span data-stu-id="e489f-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="e489f-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e489f-140">требования</span><span class="sxs-lookup"><span data-stu-id="e489f-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="e489f-141">Создание</span><span class="sxs-lookup"><span data-stu-id="e489f-141">Compose</span></span><br><span data-ttu-id="e489f-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-142">Read</span></span> | [<span data-ttu-id="e489f-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="e489f-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="e489f-144">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e489f-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="e489f-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="e489f-146">Создание</span><span class="sxs-lookup"><span data-stu-id="e489f-146">Compose</span></span><br><span data-ttu-id="e489f-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-147">Read</span></span> | [<span data-ttu-id="e489f-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e489f-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="e489f-149">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e489f-150">ui</span><span class="sxs-lookup"><span data-stu-id="e489f-150">ui</span></span>](#ui-ui) | <span data-ttu-id="e489f-151">Создание</span><span class="sxs-lookup"><span data-stu-id="e489f-151">Compose</span></span><br><span data-ttu-id="e489f-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-152">Read</span></span> | [<span data-ttu-id="e489f-153">UI</span><span class="sxs-lookup"><span data-stu-id="e489f-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3&preserve-view=true) | [<span data-ttu-id="e489f-154">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="e489f-155">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="e489f-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="e489f-156">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="e489f-156">contentLanguage: String</span></span>

<span data-ttu-id="e489f-157">Получает локализ (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="e489f-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="e489f-158">Это значение отражает текущий параметр Язык редактирования, указанный в файле > `contentLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="e489f-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="e489f-159">Тип</span><span class="sxs-lookup"><span data-stu-id="e489f-159">Type</span></span>

*   <span data-ttu-id="e489f-160">String</span><span class="sxs-lookup"><span data-stu-id="e489f-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e489f-161">Требования</span><span class="sxs-lookup"><span data-stu-id="e489f-161">Requirements</span></span>

|<span data-ttu-id="e489f-162">Требование</span><span class="sxs-lookup"><span data-stu-id="e489f-162">Requirement</span></span>| <span data-ttu-id="e489f-163">Значение</span><span class="sxs-lookup"><span data-stu-id="e489f-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="e489f-164">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e489f-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e489f-165">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-165">1.1</span></span>|
|[<span data-ttu-id="e489f-166">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e489f-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e489f-167">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e489f-168">Пример</span><span class="sxs-lookup"><span data-stu-id="e489f-168">Example</span></span>

```js
function sayHelloWithContentLanguage() {
  var myContentLanguage = Office.context.contentLanguage;
  switch (myContentLanguage) {
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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="e489f-169">диагностика: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="e489f-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="e489f-170">Получает сведения об среде, в которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="e489f-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="e489f-171">Тип</span><span class="sxs-lookup"><span data-stu-id="e489f-171">Type</span></span>

*   [<span data-ttu-id="e489f-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="e489f-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="e489f-173">Требования</span><span class="sxs-lookup"><span data-stu-id="e489f-173">Requirements</span></span>

|<span data-ttu-id="e489f-174">Требование</span><span class="sxs-lookup"><span data-stu-id="e489f-174">Requirement</span></span>| <span data-ttu-id="e489f-175">Значение</span><span class="sxs-lookup"><span data-stu-id="e489f-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="e489f-176">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e489f-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e489f-177">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-177">1.1</span></span>|
|[<span data-ttu-id="e489f-178">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e489f-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e489f-179">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e489f-180">Пример</span><span class="sxs-lookup"><span data-stu-id="e489f-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="e489f-181">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="e489f-181">displayLanguage: String</span></span>

<span data-ttu-id="e489f-182">Получает локализ (язык) в формате языковых тегов RFC 1766, заданный пользователем для пользовательского интерфейса Office клиентского приложения.</span><span class="sxs-lookup"><span data-stu-id="e489f-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="e489f-183">Это значение отражает текущий параметр Язык отображения, указанный в файле > `displayLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="e489f-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="e489f-184">Тип</span><span class="sxs-lookup"><span data-stu-id="e489f-184">Type</span></span>

*   <span data-ttu-id="e489f-185">String</span><span class="sxs-lookup"><span data-stu-id="e489f-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e489f-186">Требования</span><span class="sxs-lookup"><span data-stu-id="e489f-186">Requirements</span></span>

|<span data-ttu-id="e489f-187">Требование</span><span class="sxs-lookup"><span data-stu-id="e489f-187">Requirement</span></span>| <span data-ttu-id="e489f-188">Значение</span><span class="sxs-lookup"><span data-stu-id="e489f-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="e489f-189">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e489f-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e489f-190">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-190">1.1</span></span>|
|[<span data-ttu-id="e489f-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e489f-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e489f-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e489f-193">Пример</span><span class="sxs-lookup"><span data-stu-id="e489f-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="e489f-194">требования: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="e489f-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="e489f-195">Предоставляет метод определения, какие наборы требований поддерживаются в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="e489f-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="e489f-196">Тип</span><span class="sxs-lookup"><span data-stu-id="e489f-196">Type</span></span>

*   [<span data-ttu-id="e489f-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="e489f-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="e489f-198">Требования</span><span class="sxs-lookup"><span data-stu-id="e489f-198">Requirements</span></span>

|<span data-ttu-id="e489f-199">Требование</span><span class="sxs-lookup"><span data-stu-id="e489f-199">Requirement</span></span>| <span data-ttu-id="e489f-200">Значение</span><span class="sxs-lookup"><span data-stu-id="e489f-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="e489f-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e489f-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e489f-202">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-202">1.1</span></span>|
|[<span data-ttu-id="e489f-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e489f-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e489f-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e489f-205">Пример</span><span class="sxs-lookup"><span data-stu-id="e489f-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="e489f-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="e489f-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="e489f-207">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="e489f-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="e489f-208">Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранимой в почтовом ящике пользователя, чтобы она была доступна этой надстройке, когда она запущена из любого клиента Outlook, используемого для доступа к этому `RoamingSettings` почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="e489f-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="e489f-209">Тип</span><span class="sxs-lookup"><span data-stu-id="e489f-209">Type</span></span>

*   [<span data-ttu-id="e489f-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e489f-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="e489f-211">Требования</span><span class="sxs-lookup"><span data-stu-id="e489f-211">Requirements</span></span>

|<span data-ttu-id="e489f-212">Требование</span><span class="sxs-lookup"><span data-stu-id="e489f-212">Requirement</span></span>| <span data-ttu-id="e489f-213">Значение</span><span class="sxs-lookup"><span data-stu-id="e489f-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="e489f-214">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e489f-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e489f-215">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-215">1.1</span></span>|
|[<span data-ttu-id="e489f-216">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e489f-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="e489f-217">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e489f-217">Restricted</span></span>|
|[<span data-ttu-id="e489f-218">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e489f-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e489f-219">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="e489f-220">ui: [пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="e489f-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="e489f-221">Предоставляет объекты и методы, которые можно использовать для создания и управления компонентами пользовательского интерфейса, такими как диалоговое окно, в Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="e489f-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="e489f-222">Тип</span><span class="sxs-lookup"><span data-stu-id="e489f-222">Type</span></span>

*   [<span data-ttu-id="e489f-223">UI</span><span class="sxs-lookup"><span data-stu-id="e489f-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="e489f-224">Требования</span><span class="sxs-lookup"><span data-stu-id="e489f-224">Requirements</span></span>

|<span data-ttu-id="e489f-225">Требование</span><span class="sxs-lookup"><span data-stu-id="e489f-225">Requirement</span></span>| <span data-ttu-id="e489f-226">Значение</span><span class="sxs-lookup"><span data-stu-id="e489f-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="e489f-227">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e489f-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e489f-228">1.1</span><span class="sxs-lookup"><span data-stu-id="e489f-228">1.1</span></span>|
|[<span data-ttu-id="e489f-229">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e489f-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e489f-230">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e489f-230">Compose or Read</span></span>|
