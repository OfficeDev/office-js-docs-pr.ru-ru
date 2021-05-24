---
title: Office.context — набор требований 1.6
description: Office. Участники объектов Context, доступные для Outlook надстройки с помощью API почтовых ящиков, устанавливают 1.6.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: d4c65cea9b581665e0dc7b38a8e0bf10d6b544f9
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591004"
---
# <a name="context-mailbox-requirement-set-16"></a><span data-ttu-id="40d7a-103">контекст (набор требований к почтовым ящикам 1.6)</span><span class="sxs-lookup"><span data-stu-id="40d7a-103">context (Mailbox requirement set 1.6)</span></span>

### <a name="officecontext"></a><span data-ttu-id="40d7a-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="40d7a-104">[Office](office.md).context</span></span>

<span data-ttu-id="40d7a-105">Office.context предоставляет общие интерфейсы, используемые надстройки во всех Office приложениях.</span><span class="sxs-lookup"><span data-stu-id="40d7a-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="40d7a-106">Этот список документов только те интерфейсы, которые используются Outlook надстройки. Полный список пространства имен Office.context см. в [ссылке Office.context в общем API.](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="40d7a-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="40d7a-107">Требования</span><span class="sxs-lookup"><span data-stu-id="40d7a-107">Requirements</span></span>

|<span data-ttu-id="40d7a-108">Требование</span><span class="sxs-lookup"><span data-stu-id="40d7a-108">Requirement</span></span>| <span data-ttu-id="40d7a-109">Значение</span><span class="sxs-lookup"><span data-stu-id="40d7a-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="40d7a-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="40d7a-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="40d7a-111">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-111">1.1</span></span>|
|[<span data-ttu-id="40d7a-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40d7a-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="40d7a-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="40d7a-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="40d7a-114">Properties</span></span>

| <span data-ttu-id="40d7a-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="40d7a-115">Property</span></span> | <span data-ttu-id="40d7a-116">Режимы</span><span class="sxs-lookup"><span data-stu-id="40d7a-116">Modes</span></span> | <span data-ttu-id="40d7a-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="40d7a-117">Return type</span></span> | <span data-ttu-id="40d7a-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="40d7a-118">Minimum</span></span><br><span data-ttu-id="40d7a-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="40d7a-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="40d7a-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="40d7a-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="40d7a-121">Создание</span><span class="sxs-lookup"><span data-stu-id="40d7a-121">Compose</span></span><br><span data-ttu-id="40d7a-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-122">Read</span></span> | <span data-ttu-id="40d7a-123">Строка</span><span class="sxs-lookup"><span data-stu-id="40d7a-123">String</span></span> | [<span data-ttu-id="40d7a-124">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="40d7a-125">диагностика</span><span class="sxs-lookup"><span data-stu-id="40d7a-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="40d7a-126">Создание</span><span class="sxs-lookup"><span data-stu-id="40d7a-126">Compose</span></span><br><span data-ttu-id="40d7a-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-127">Read</span></span> | [<span data-ttu-id="40d7a-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="40d7a-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="40d7a-129">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="40d7a-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="40d7a-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="40d7a-131">Создание</span><span class="sxs-lookup"><span data-stu-id="40d7a-131">Compose</span></span><br><span data-ttu-id="40d7a-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-132">Read</span></span> | <span data-ttu-id="40d7a-133">Строка</span><span class="sxs-lookup"><span data-stu-id="40d7a-133">String</span></span> | [<span data-ttu-id="40d7a-134">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="40d7a-135">хост</span><span class="sxs-lookup"><span data-stu-id="40d7a-135">host</span></span>](#host-hosttype) | <span data-ttu-id="40d7a-136">Создание</span><span class="sxs-lookup"><span data-stu-id="40d7a-136">Compose</span></span><br><span data-ttu-id="40d7a-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-137">Read</span></span> | [<span data-ttu-id="40d7a-138">HostType</span><span class="sxs-lookup"><span data-stu-id="40d7a-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="40d7a-139">1.5</span><span class="sxs-lookup"><span data-stu-id="40d7a-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="40d7a-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="40d7a-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="40d7a-141">Создание</span><span class="sxs-lookup"><span data-stu-id="40d7a-141">Compose</span></span><br><span data-ttu-id="40d7a-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-142">Read</span></span> | [<span data-ttu-id="40d7a-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="40d7a-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="40d7a-144">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="40d7a-145">платформа</span><span class="sxs-lookup"><span data-stu-id="40d7a-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="40d7a-146">Создание</span><span class="sxs-lookup"><span data-stu-id="40d7a-146">Compose</span></span><br><span data-ttu-id="40d7a-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-147">Read</span></span> | [<span data-ttu-id="40d7a-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="40d7a-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="40d7a-149">1.5</span><span class="sxs-lookup"><span data-stu-id="40d7a-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="40d7a-150">требования</span><span class="sxs-lookup"><span data-stu-id="40d7a-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="40d7a-151">Создание</span><span class="sxs-lookup"><span data-stu-id="40d7a-151">Compose</span></span><br><span data-ttu-id="40d7a-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-152">Read</span></span> | [<span data-ttu-id="40d7a-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="40d7a-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="40d7a-154">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="40d7a-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="40d7a-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="40d7a-156">Создание</span><span class="sxs-lookup"><span data-stu-id="40d7a-156">Compose</span></span><br><span data-ttu-id="40d7a-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-157">Read</span></span> | [<span data-ttu-id="40d7a-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="40d7a-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="40d7a-159">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="40d7a-160">ui</span><span class="sxs-lookup"><span data-stu-id="40d7a-160">ui</span></span>](#ui-ui) | <span data-ttu-id="40d7a-161">Создание</span><span class="sxs-lookup"><span data-stu-id="40d7a-161">Compose</span></span><br><span data-ttu-id="40d7a-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-162">Read</span></span> | [<span data-ttu-id="40d7a-163">UI</span><span class="sxs-lookup"><span data-stu-id="40d7a-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="40d7a-164">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="40d7a-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="40d7a-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="40d7a-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="40d7a-166">contentLanguage: String</span></span>

<span data-ttu-id="40d7a-167">Получает локализ (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="40d7a-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="40d7a-168">Это значение отражает текущий параметр Язык редактирования, указанный в файле > `contentLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="40d7a-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="40d7a-169">Тип</span><span class="sxs-lookup"><span data-stu-id="40d7a-169">Type</span></span>

*   <span data-ttu-id="40d7a-170">String</span><span class="sxs-lookup"><span data-stu-id="40d7a-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="40d7a-171">Требования</span><span class="sxs-lookup"><span data-stu-id="40d7a-171">Requirements</span></span>

|<span data-ttu-id="40d7a-172">Требование</span><span class="sxs-lookup"><span data-stu-id="40d7a-172">Requirement</span></span>| <span data-ttu-id="40d7a-173">Значение</span><span class="sxs-lookup"><span data-stu-id="40d7a-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="40d7a-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="40d7a-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="40d7a-175">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-175">1.1</span></span>|
|[<span data-ttu-id="40d7a-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40d7a-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="40d7a-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40d7a-178">Пример</span><span class="sxs-lookup"><span data-stu-id="40d7a-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="40d7a-179">диагностика: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="40d7a-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="40d7a-180">Получает сведения об среде, в которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="40d7a-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="40d7a-181">Тип</span><span class="sxs-lookup"><span data-stu-id="40d7a-181">Type</span></span>

*   [<span data-ttu-id="40d7a-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="40d7a-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="40d7a-183">Требования</span><span class="sxs-lookup"><span data-stu-id="40d7a-183">Requirements</span></span>

|<span data-ttu-id="40d7a-184">Требование</span><span class="sxs-lookup"><span data-stu-id="40d7a-184">Requirement</span></span>| <span data-ttu-id="40d7a-185">Значение</span><span class="sxs-lookup"><span data-stu-id="40d7a-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="40d7a-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="40d7a-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="40d7a-187">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-187">1.1</span></span>|
|[<span data-ttu-id="40d7a-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40d7a-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="40d7a-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40d7a-190">Пример</span><span class="sxs-lookup"><span data-stu-id="40d7a-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="40d7a-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="40d7a-191">displayLanguage: String</span></span>

<span data-ttu-id="40d7a-192">Получает локализ (язык) в формате языковых тегов RFC 1766, заданный пользователем для пользовательского интерфейса Office клиентского приложения.</span><span class="sxs-lookup"><span data-stu-id="40d7a-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="40d7a-193">Это значение отражает текущий параметр Язык отображения, указанный в файле > `displayLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="40d7a-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="40d7a-194">Тип</span><span class="sxs-lookup"><span data-stu-id="40d7a-194">Type</span></span>

*   <span data-ttu-id="40d7a-195">String</span><span class="sxs-lookup"><span data-stu-id="40d7a-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="40d7a-196">Требования</span><span class="sxs-lookup"><span data-stu-id="40d7a-196">Requirements</span></span>

|<span data-ttu-id="40d7a-197">Требование</span><span class="sxs-lookup"><span data-stu-id="40d7a-197">Requirement</span></span>| <span data-ttu-id="40d7a-198">Значение</span><span class="sxs-lookup"><span data-stu-id="40d7a-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="40d7a-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="40d7a-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="40d7a-200">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-200">1.1</span></span>|
|[<span data-ttu-id="40d7a-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40d7a-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="40d7a-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40d7a-203">Пример</span><span class="sxs-lookup"><span data-stu-id="40d7a-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="40d7a-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="40d7a-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="40d7a-205">Получает Office приложение, в которое размещена надстройка.</span><span class="sxs-lookup"><span data-stu-id="40d7a-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="40d7a-206">Кроме того, для получения хоста можно использовать [свойство Office.context.diagnostics.](#diagnostics-contextinformation)</span><span class="sxs-lookup"><span data-stu-id="40d7a-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="40d7a-207">Тип</span><span class="sxs-lookup"><span data-stu-id="40d7a-207">Type</span></span>

*   [<span data-ttu-id="40d7a-208">HostType</span><span class="sxs-lookup"><span data-stu-id="40d7a-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="40d7a-209">Требования</span><span class="sxs-lookup"><span data-stu-id="40d7a-209">Requirements</span></span>

|<span data-ttu-id="40d7a-210">Требование</span><span class="sxs-lookup"><span data-stu-id="40d7a-210">Requirement</span></span>| <span data-ttu-id="40d7a-211">Значение</span><span class="sxs-lookup"><span data-stu-id="40d7a-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="40d7a-212">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="40d7a-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="40d7a-213">1.5</span><span class="sxs-lookup"><span data-stu-id="40d7a-213">1.5</span></span>|
|[<span data-ttu-id="40d7a-214">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40d7a-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="40d7a-215">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40d7a-216">Пример</span><span class="sxs-lookup"><span data-stu-id="40d7a-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="40d7a-217">платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="40d7a-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="40d7a-218">Предоставляет платформу, на которой запущена надстройка.</span><span class="sxs-lookup"><span data-stu-id="40d7a-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="40d7a-219">Кроме того, для получения [платформы можно использовать свойство Office.context.diagnostics.](#diagnostics-contextinformation)</span><span class="sxs-lookup"><span data-stu-id="40d7a-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="40d7a-220">Тип</span><span class="sxs-lookup"><span data-stu-id="40d7a-220">Type</span></span>

*   [<span data-ttu-id="40d7a-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="40d7a-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="40d7a-222">Требования</span><span class="sxs-lookup"><span data-stu-id="40d7a-222">Requirements</span></span>

|<span data-ttu-id="40d7a-223">Требование</span><span class="sxs-lookup"><span data-stu-id="40d7a-223">Requirement</span></span>| <span data-ttu-id="40d7a-224">Значение</span><span class="sxs-lookup"><span data-stu-id="40d7a-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="40d7a-225">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="40d7a-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="40d7a-226">1.5</span><span class="sxs-lookup"><span data-stu-id="40d7a-226">1.5</span></span>|
|[<span data-ttu-id="40d7a-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40d7a-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="40d7a-228">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40d7a-229">Пример</span><span class="sxs-lookup"><span data-stu-id="40d7a-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="40d7a-230">требования: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="40d7a-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="40d7a-231">Предоставляет метод определения, какие наборы требований поддерживаются в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="40d7a-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="40d7a-232">Тип</span><span class="sxs-lookup"><span data-stu-id="40d7a-232">Type</span></span>

*   [<span data-ttu-id="40d7a-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="40d7a-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="40d7a-234">Требования</span><span class="sxs-lookup"><span data-stu-id="40d7a-234">Requirements</span></span>

|<span data-ttu-id="40d7a-235">Требование</span><span class="sxs-lookup"><span data-stu-id="40d7a-235">Requirement</span></span>| <span data-ttu-id="40d7a-236">Значение</span><span class="sxs-lookup"><span data-stu-id="40d7a-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="40d7a-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="40d7a-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="40d7a-238">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-238">1.1</span></span>|
|[<span data-ttu-id="40d7a-239">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40d7a-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="40d7a-240">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="40d7a-241">Пример</span><span class="sxs-lookup"><span data-stu-id="40d7a-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="40d7a-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="40d7a-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="40d7a-243">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="40d7a-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="40d7a-244">Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранимой в почтовом ящике пользователя, чтобы она была доступна этой надстройке, когда она запущена из любого клиента Outlook, используемого для доступа к этому `RoamingSettings` почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="40d7a-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="40d7a-245">Тип</span><span class="sxs-lookup"><span data-stu-id="40d7a-245">Type</span></span>

*   [<span data-ttu-id="40d7a-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="40d7a-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="40d7a-247">Требования</span><span class="sxs-lookup"><span data-stu-id="40d7a-247">Requirements</span></span>

|<span data-ttu-id="40d7a-248">Требование</span><span class="sxs-lookup"><span data-stu-id="40d7a-248">Requirement</span></span>| <span data-ttu-id="40d7a-249">Значение</span><span class="sxs-lookup"><span data-stu-id="40d7a-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="40d7a-250">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="40d7a-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="40d7a-251">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-251">1.1</span></span>|
|[<span data-ttu-id="40d7a-252">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="40d7a-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="40d7a-253">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="40d7a-253">Restricted</span></span>|
|[<span data-ttu-id="40d7a-254">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40d7a-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="40d7a-255">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="40d7a-256">ui: [пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="40d7a-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="40d7a-257">Предоставляет объекты и методы, которые можно использовать для создания и управления компонентами пользовательского интерфейса, такими как диалоговое окно, в Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="40d7a-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="40d7a-258">Тип</span><span class="sxs-lookup"><span data-stu-id="40d7a-258">Type</span></span>

*   [<span data-ttu-id="40d7a-259">UI</span><span class="sxs-lookup"><span data-stu-id="40d7a-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="40d7a-260">Требования</span><span class="sxs-lookup"><span data-stu-id="40d7a-260">Requirements</span></span>

|<span data-ttu-id="40d7a-261">Требование</span><span class="sxs-lookup"><span data-stu-id="40d7a-261">Requirement</span></span>| <span data-ttu-id="40d7a-262">Значение</span><span class="sxs-lookup"><span data-stu-id="40d7a-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="40d7a-263">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="40d7a-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="40d7a-264">1.1</span><span class="sxs-lookup"><span data-stu-id="40d7a-264">1.1</span></span>|
|[<span data-ttu-id="40d7a-265">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="40d7a-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="40d7a-266">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="40d7a-266">Compose or Read</span></span>|
