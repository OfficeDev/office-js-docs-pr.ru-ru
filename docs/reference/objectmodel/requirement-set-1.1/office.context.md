---
title: Office.context — набор требований 1.1
description: Office. Участники объектов context, доступные для Outlook надстройки с помощью API почтовых ящиков, устанавливают 1.1.
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 41273bfc5362a9d5572e38b8e80b81041f5aa312
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590878"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="315e1-103">контекст (набор требований к почтовым ящикам 1.1)</span><span class="sxs-lookup"><span data-stu-id="315e1-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="315e1-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="315e1-104">[Office](office.md).context</span></span>

<span data-ttu-id="315e1-105">Office.context предоставляет общие интерфейсы, используемые надстройки во всех Office приложениях.</span><span class="sxs-lookup"><span data-stu-id="315e1-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="315e1-106">Этот список документов только те интерфейсы, которые используются Outlook надстройки. Полный список пространства имен Office.context см. в [ссылке Office.context в общем API.](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="315e1-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="315e1-107">Требования</span><span class="sxs-lookup"><span data-stu-id="315e1-107">Requirements</span></span>

|<span data-ttu-id="315e1-108">Требование</span><span class="sxs-lookup"><span data-stu-id="315e1-108">Requirement</span></span>| <span data-ttu-id="315e1-109">Значение</span><span class="sxs-lookup"><span data-stu-id="315e1-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="315e1-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="315e1-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="315e1-111">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-111">1.1</span></span>|
|[<span data-ttu-id="315e1-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="315e1-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="315e1-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="315e1-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="315e1-114">Properties</span></span>

| <span data-ttu-id="315e1-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="315e1-115">Property</span></span> | <span data-ttu-id="315e1-116">Режимы</span><span class="sxs-lookup"><span data-stu-id="315e1-116">Modes</span></span> | <span data-ttu-id="315e1-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="315e1-117">Return type</span></span> | <span data-ttu-id="315e1-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="315e1-118">Minimum</span></span><br><span data-ttu-id="315e1-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="315e1-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="315e1-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="315e1-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="315e1-121">Создание</span><span class="sxs-lookup"><span data-stu-id="315e1-121">Compose</span></span><br><span data-ttu-id="315e1-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-122">Read</span></span> | <span data-ttu-id="315e1-123">Строка</span><span class="sxs-lookup"><span data-stu-id="315e1-123">String</span></span> | [<span data-ttu-id="315e1-124">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="315e1-125">диагностика</span><span class="sxs-lookup"><span data-stu-id="315e1-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="315e1-126">Создание</span><span class="sxs-lookup"><span data-stu-id="315e1-126">Compose</span></span><br><span data-ttu-id="315e1-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-127">Read</span></span> | [<span data-ttu-id="315e1-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="315e1-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="315e1-129">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="315e1-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="315e1-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="315e1-131">Создание</span><span class="sxs-lookup"><span data-stu-id="315e1-131">Compose</span></span><br><span data-ttu-id="315e1-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-132">Read</span></span> | <span data-ttu-id="315e1-133">Строка</span><span class="sxs-lookup"><span data-stu-id="315e1-133">String</span></span> | [<span data-ttu-id="315e1-134">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="315e1-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="315e1-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="315e1-136">Создание</span><span class="sxs-lookup"><span data-stu-id="315e1-136">Compose</span></span><br><span data-ttu-id="315e1-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-137">Read</span></span> | [<span data-ttu-id="315e1-138">Mailbox</span><span class="sxs-lookup"><span data-stu-id="315e1-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="315e1-139">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="315e1-140">требования</span><span class="sxs-lookup"><span data-stu-id="315e1-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="315e1-141">Создание</span><span class="sxs-lookup"><span data-stu-id="315e1-141">Compose</span></span><br><span data-ttu-id="315e1-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-142">Read</span></span> | [<span data-ttu-id="315e1-143">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="315e1-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="315e1-144">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="315e1-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="315e1-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="315e1-146">Создание</span><span class="sxs-lookup"><span data-stu-id="315e1-146">Compose</span></span><br><span data-ttu-id="315e1-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-147">Read</span></span> | [<span data-ttu-id="315e1-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="315e1-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="315e1-149">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="315e1-150">ui</span><span class="sxs-lookup"><span data-stu-id="315e1-150">ui</span></span>](#ui-ui) | <span data-ttu-id="315e1-151">Создание</span><span class="sxs-lookup"><span data-stu-id="315e1-151">Compose</span></span><br><span data-ttu-id="315e1-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-152">Read</span></span> | [<span data-ttu-id="315e1-153">UI</span><span class="sxs-lookup"><span data-stu-id="315e1-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1&preserve-view=true) | [<span data-ttu-id="315e1-154">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="315e1-155">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="315e1-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="315e1-156">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="315e1-156">contentLanguage: String</span></span>

<span data-ttu-id="315e1-157">Получает локализ (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="315e1-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="315e1-158">Это значение отражает текущий параметр Язык редактирования, указанный в файле > `contentLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="315e1-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="315e1-159">Тип</span><span class="sxs-lookup"><span data-stu-id="315e1-159">Type</span></span>

*   <span data-ttu-id="315e1-160">String</span><span class="sxs-lookup"><span data-stu-id="315e1-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="315e1-161">Требования</span><span class="sxs-lookup"><span data-stu-id="315e1-161">Requirements</span></span>

|<span data-ttu-id="315e1-162">Требование</span><span class="sxs-lookup"><span data-stu-id="315e1-162">Requirement</span></span>| <span data-ttu-id="315e1-163">Значение</span><span class="sxs-lookup"><span data-stu-id="315e1-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="315e1-164">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="315e1-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="315e1-165">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-165">1.1</span></span>|
|[<span data-ttu-id="315e1-166">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="315e1-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="315e1-167">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="315e1-168">Пример</span><span class="sxs-lookup"><span data-stu-id="315e1-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="315e1-169">диагностика: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="315e1-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="315e1-170">Получает сведения об среде, в которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="315e1-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="315e1-171">Тип</span><span class="sxs-lookup"><span data-stu-id="315e1-171">Type</span></span>

*   [<span data-ttu-id="315e1-172">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="315e1-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="315e1-173">Требования</span><span class="sxs-lookup"><span data-stu-id="315e1-173">Requirements</span></span>

|<span data-ttu-id="315e1-174">Требование</span><span class="sxs-lookup"><span data-stu-id="315e1-174">Requirement</span></span>| <span data-ttu-id="315e1-175">Значение</span><span class="sxs-lookup"><span data-stu-id="315e1-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="315e1-176">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="315e1-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="315e1-177">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-177">1.1</span></span>|
|[<span data-ttu-id="315e1-178">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="315e1-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="315e1-179">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="315e1-180">Пример</span><span class="sxs-lookup"><span data-stu-id="315e1-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="315e1-181">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="315e1-181">displayLanguage: String</span></span>

<span data-ttu-id="315e1-182">Получает локализ (язык) в формате языковых тегов RFC 1766, заданный пользователем для пользовательского интерфейса Office клиентского приложения.</span><span class="sxs-lookup"><span data-stu-id="315e1-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="315e1-183">Это значение отражает текущий параметр Язык отображения, указанный в файле > `displayLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="315e1-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="315e1-184">Тип</span><span class="sxs-lookup"><span data-stu-id="315e1-184">Type</span></span>

*   <span data-ttu-id="315e1-185">String</span><span class="sxs-lookup"><span data-stu-id="315e1-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="315e1-186">Требования</span><span class="sxs-lookup"><span data-stu-id="315e1-186">Requirements</span></span>

|<span data-ttu-id="315e1-187">Требование</span><span class="sxs-lookup"><span data-stu-id="315e1-187">Requirement</span></span>| <span data-ttu-id="315e1-188">Значение</span><span class="sxs-lookup"><span data-stu-id="315e1-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="315e1-189">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="315e1-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="315e1-190">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-190">1.1</span></span>|
|[<span data-ttu-id="315e1-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="315e1-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="315e1-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="315e1-193">Пример</span><span class="sxs-lookup"><span data-stu-id="315e1-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="315e1-194">требования: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="315e1-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="315e1-195">Предоставляет метод определения, какие наборы требований поддерживаются в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="315e1-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="315e1-196">Тип</span><span class="sxs-lookup"><span data-stu-id="315e1-196">Type</span></span>

*   [<span data-ttu-id="315e1-197">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="315e1-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="315e1-198">Требования</span><span class="sxs-lookup"><span data-stu-id="315e1-198">Requirements</span></span>

|<span data-ttu-id="315e1-199">Требование</span><span class="sxs-lookup"><span data-stu-id="315e1-199">Requirement</span></span>| <span data-ttu-id="315e1-200">Значение</span><span class="sxs-lookup"><span data-stu-id="315e1-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="315e1-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="315e1-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="315e1-202">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-202">1.1</span></span>|
|[<span data-ttu-id="315e1-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="315e1-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="315e1-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="315e1-205">Пример</span><span class="sxs-lookup"><span data-stu-id="315e1-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="315e1-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="315e1-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="315e1-207">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="315e1-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="315e1-208">Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранимой в почтовом ящике пользователя, чтобы она была доступна этой надстройке, когда она запущена из любого клиента Outlook, используемого для доступа к этому `RoamingSettings` почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="315e1-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="315e1-209">Тип</span><span class="sxs-lookup"><span data-stu-id="315e1-209">Type</span></span>

*   [<span data-ttu-id="315e1-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="315e1-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="315e1-211">Требования</span><span class="sxs-lookup"><span data-stu-id="315e1-211">Requirements</span></span>

|<span data-ttu-id="315e1-212">Требование</span><span class="sxs-lookup"><span data-stu-id="315e1-212">Requirement</span></span>| <span data-ttu-id="315e1-213">Значение</span><span class="sxs-lookup"><span data-stu-id="315e1-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="315e1-214">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="315e1-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="315e1-215">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-215">1.1</span></span>|
|[<span data-ttu-id="315e1-216">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="315e1-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="315e1-217">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="315e1-217">Restricted</span></span>|
|[<span data-ttu-id="315e1-218">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="315e1-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="315e1-219">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="315e1-220">ui: [пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="315e1-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="315e1-221">Предоставляет объекты и методы, которые можно использовать для создания и управления компонентами пользовательского интерфейса, такими как диалоговое окно, в Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="315e1-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="315e1-222">Тип</span><span class="sxs-lookup"><span data-stu-id="315e1-222">Type</span></span>

*   [<span data-ttu-id="315e1-223">UI</span><span class="sxs-lookup"><span data-stu-id="315e1-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="315e1-224">Требования</span><span class="sxs-lookup"><span data-stu-id="315e1-224">Requirements</span></span>

|<span data-ttu-id="315e1-225">Требование</span><span class="sxs-lookup"><span data-stu-id="315e1-225">Requirement</span></span>| <span data-ttu-id="315e1-226">Значение</span><span class="sxs-lookup"><span data-stu-id="315e1-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="315e1-227">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="315e1-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="315e1-228">1.1</span><span class="sxs-lookup"><span data-stu-id="315e1-228">1.1</span></span>|
|[<span data-ttu-id="315e1-229">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="315e1-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="315e1-230">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="315e1-230">Compose or Read</span></span>|
