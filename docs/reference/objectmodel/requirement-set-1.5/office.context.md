---
title: Office.context — набор требований 1.5
description: Office. Участники объектов context, доступные для Outlook надстройки с помощью API почтовых ящиков, устанавливают 1.5.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 8aedd711665d902cf3cc733901df9e3a3cc86886
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591011"
---
# <a name="context-mailbox-requirement-set-15"></a><span data-ttu-id="1b9a1-103">контекст (набор требований к почтовым ящикам 1.5)</span><span class="sxs-lookup"><span data-stu-id="1b9a1-103">context (Mailbox requirement set 1.5)</span></span>

### <a name="officecontext"></a><span data-ttu-id="1b9a1-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="1b9a1-104">[Office](office.md).context</span></span>

<span data-ttu-id="1b9a1-105">Office.context предоставляет общие интерфейсы, используемые надстройки во всех Office приложениях.</span><span class="sxs-lookup"><span data-stu-id="1b9a1-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="1b9a1-106">Этот список документов только те интерфейсы, которые используются Outlook надстройки. Полный список пространства имен Office.context см. в [ссылке Office.context в общем API.](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="1b9a1-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b9a1-107">Требования</span><span class="sxs-lookup"><span data-stu-id="1b9a1-107">Requirements</span></span>

|<span data-ttu-id="1b9a1-108">Требование</span><span class="sxs-lookup"><span data-stu-id="1b9a1-108">Requirement</span></span>| <span data-ttu-id="1b9a1-109">Значение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b9a1-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1b9a1-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1b9a1-111">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-111">1.1</span></span>|
|[<span data-ttu-id="1b9a1-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1b9a1-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1b9a1-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="1b9a1-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="1b9a1-114">Properties</span></span>

| <span data-ttu-id="1b9a1-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="1b9a1-115">Property</span></span> | <span data-ttu-id="1b9a1-116">Режимы</span><span class="sxs-lookup"><span data-stu-id="1b9a1-116">Modes</span></span> | <span data-ttu-id="1b9a1-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="1b9a1-117">Return type</span></span> | <span data-ttu-id="1b9a1-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="1b9a1-118">Minimum</span></span><br><span data-ttu-id="1b9a1-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="1b9a1-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1b9a1-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="1b9a1-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="1b9a1-121">Создание</span><span class="sxs-lookup"><span data-stu-id="1b9a1-121">Compose</span></span><br><span data-ttu-id="1b9a1-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-122">Read</span></span> | <span data-ttu-id="1b9a1-123">Строка</span><span class="sxs-lookup"><span data-stu-id="1b9a1-123">String</span></span> | [<span data-ttu-id="1b9a1-124">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1b9a1-125">диагностика</span><span class="sxs-lookup"><span data-stu-id="1b9a1-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="1b9a1-126">Создание</span><span class="sxs-lookup"><span data-stu-id="1b9a1-126">Compose</span></span><br><span data-ttu-id="1b9a1-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-127">Read</span></span> | [<span data-ttu-id="1b9a1-128">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1b9a1-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="1b9a1-129">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1b9a1-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="1b9a1-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="1b9a1-131">Создание</span><span class="sxs-lookup"><span data-stu-id="1b9a1-131">Compose</span></span><br><span data-ttu-id="1b9a1-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-132">Read</span></span> | <span data-ttu-id="1b9a1-133">Строка</span><span class="sxs-lookup"><span data-stu-id="1b9a1-133">String</span></span> | [<span data-ttu-id="1b9a1-134">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1b9a1-135">хост</span><span class="sxs-lookup"><span data-stu-id="1b9a1-135">host</span></span>](#host-hosttype) | <span data-ttu-id="1b9a1-136">Создание</span><span class="sxs-lookup"><span data-stu-id="1b9a1-136">Compose</span></span><br><span data-ttu-id="1b9a1-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-137">Read</span></span> | [<span data-ttu-id="1b9a1-138">HostType</span><span class="sxs-lookup"><span data-stu-id="1b9a1-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="1b9a1-139">1.5</span><span class="sxs-lookup"><span data-stu-id="1b9a1-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="1b9a1-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="1b9a1-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="1b9a1-141">Создание</span><span class="sxs-lookup"><span data-stu-id="1b9a1-141">Compose</span></span><br><span data-ttu-id="1b9a1-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-142">Read</span></span> | [<span data-ttu-id="1b9a1-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="1b9a1-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="1b9a1-144">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1b9a1-145">платформа</span><span class="sxs-lookup"><span data-stu-id="1b9a1-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="1b9a1-146">Создание</span><span class="sxs-lookup"><span data-stu-id="1b9a1-146">Compose</span></span><br><span data-ttu-id="1b9a1-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-147">Read</span></span> | [<span data-ttu-id="1b9a1-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1b9a1-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="1b9a1-149">1.5</span><span class="sxs-lookup"><span data-stu-id="1b9a1-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="1b9a1-150">требования</span><span class="sxs-lookup"><span data-stu-id="1b9a1-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="1b9a1-151">Создание</span><span class="sxs-lookup"><span data-stu-id="1b9a1-151">Compose</span></span><br><span data-ttu-id="1b9a1-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-152">Read</span></span> | [<span data-ttu-id="1b9a1-153">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1b9a1-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="1b9a1-154">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1b9a1-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="1b9a1-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="1b9a1-156">Создание</span><span class="sxs-lookup"><span data-stu-id="1b9a1-156">Compose</span></span><br><span data-ttu-id="1b9a1-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-157">Read</span></span> | [<span data-ttu-id="1b9a1-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1b9a1-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="1b9a1-159">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1b9a1-160">ui</span><span class="sxs-lookup"><span data-stu-id="1b9a1-160">ui</span></span>](#ui-ui) | <span data-ttu-id="1b9a1-161">Создание</span><span class="sxs-lookup"><span data-stu-id="1b9a1-161">Compose</span></span><br><span data-ttu-id="1b9a1-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-162">Read</span></span> | [<span data-ttu-id="1b9a1-163">UI</span><span class="sxs-lookup"><span data-stu-id="1b9a1-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="1b9a1-164">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="1b9a1-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="1b9a1-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="1b9a1-166">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="1b9a1-166">contentLanguage: String</span></span>

<span data-ttu-id="1b9a1-167">Получает локализ (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="1b9a1-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="1b9a1-168">Это значение отражает текущий параметр Язык редактирования, указанный в файле > `contentLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="1b9a1-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="1b9a1-169">Тип</span><span class="sxs-lookup"><span data-stu-id="1b9a1-169">Type</span></span>

*   <span data-ttu-id="1b9a1-170">String</span><span class="sxs-lookup"><span data-stu-id="1b9a1-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b9a1-171">Требования</span><span class="sxs-lookup"><span data-stu-id="1b9a1-171">Requirements</span></span>

|<span data-ttu-id="1b9a1-172">Требование</span><span class="sxs-lookup"><span data-stu-id="1b9a1-172">Requirement</span></span>| <span data-ttu-id="1b9a1-173">Значение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b9a1-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1b9a1-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1b9a1-175">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-175">1.1</span></span>|
|[<span data-ttu-id="1b9a1-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1b9a1-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1b9a1-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b9a1-178">Пример</span><span class="sxs-lookup"><span data-stu-id="1b9a1-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="1b9a1-179">диагностика: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="1b9a1-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="1b9a1-180">Получает сведения об среде, в которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="1b9a1-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1b9a1-181">Тип</span><span class="sxs-lookup"><span data-stu-id="1b9a1-181">Type</span></span>

*   [<span data-ttu-id="1b9a1-182">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="1b9a1-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="1b9a1-183">Требования</span><span class="sxs-lookup"><span data-stu-id="1b9a1-183">Requirements</span></span>

|<span data-ttu-id="1b9a1-184">Требование</span><span class="sxs-lookup"><span data-stu-id="1b9a1-184">Requirement</span></span>| <span data-ttu-id="1b9a1-185">Значение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b9a1-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1b9a1-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1b9a1-187">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-187">1.1</span></span>|
|[<span data-ttu-id="1b9a1-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1b9a1-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1b9a1-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b9a1-190">Пример</span><span class="sxs-lookup"><span data-stu-id="1b9a1-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="1b9a1-191">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="1b9a1-191">displayLanguage: String</span></span>

<span data-ttu-id="1b9a1-192">Получает локализ (язык) в формате языковых тегов RFC 1766, заданный пользователем для пользовательского интерфейса Office клиентского приложения.</span><span class="sxs-lookup"><span data-stu-id="1b9a1-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="1b9a1-193">Это значение отражает текущий параметр Язык отображения, указанный в файле > `displayLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="1b9a1-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="1b9a1-194">Тип</span><span class="sxs-lookup"><span data-stu-id="1b9a1-194">Type</span></span>

*   <span data-ttu-id="1b9a1-195">String</span><span class="sxs-lookup"><span data-stu-id="1b9a1-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1b9a1-196">Требования</span><span class="sxs-lookup"><span data-stu-id="1b9a1-196">Requirements</span></span>

|<span data-ttu-id="1b9a1-197">Требование</span><span class="sxs-lookup"><span data-stu-id="1b9a1-197">Requirement</span></span>| <span data-ttu-id="1b9a1-198">Значение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b9a1-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1b9a1-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1b9a1-200">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-200">1.1</span></span>|
|[<span data-ttu-id="1b9a1-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1b9a1-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1b9a1-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b9a1-203">Пример</span><span class="sxs-lookup"><span data-stu-id="1b9a1-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="1b9a1-204">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="1b9a1-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="1b9a1-205">Получает Office приложение, в которое размещена надстройка.</span><span class="sxs-lookup"><span data-stu-id="1b9a1-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="1b9a1-206">Кроме того, для получения хоста можно использовать [свойство Office.context.diagnostics.](#diagnostics-contextinformation)</span><span class="sxs-lookup"><span data-stu-id="1b9a1-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="1b9a1-207">Тип</span><span class="sxs-lookup"><span data-stu-id="1b9a1-207">Type</span></span>

*   [<span data-ttu-id="1b9a1-208">HostType</span><span class="sxs-lookup"><span data-stu-id="1b9a1-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="1b9a1-209">Требования</span><span class="sxs-lookup"><span data-stu-id="1b9a1-209">Requirements</span></span>

|<span data-ttu-id="1b9a1-210">Требование</span><span class="sxs-lookup"><span data-stu-id="1b9a1-210">Requirement</span></span>| <span data-ttu-id="1b9a1-211">Значение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b9a1-212">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1b9a1-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1b9a1-213">1.5</span><span class="sxs-lookup"><span data-stu-id="1b9a1-213">1.5</span></span>|
|[<span data-ttu-id="1b9a1-214">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1b9a1-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1b9a1-215">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b9a1-216">Пример</span><span class="sxs-lookup"><span data-stu-id="1b9a1-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="1b9a1-217">платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="1b9a1-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="1b9a1-218">Предоставляет платформу, на которой запущена надстройка.</span><span class="sxs-lookup"><span data-stu-id="1b9a1-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="1b9a1-219">Кроме того, для получения [платформы можно использовать свойство Office.context.diagnostics.](#diagnostics-contextinformation)</span><span class="sxs-lookup"><span data-stu-id="1b9a1-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="1b9a1-220">Тип</span><span class="sxs-lookup"><span data-stu-id="1b9a1-220">Type</span></span>

*   [<span data-ttu-id="1b9a1-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1b9a1-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="1b9a1-222">Требования</span><span class="sxs-lookup"><span data-stu-id="1b9a1-222">Requirements</span></span>

|<span data-ttu-id="1b9a1-223">Требование</span><span class="sxs-lookup"><span data-stu-id="1b9a1-223">Requirement</span></span>| <span data-ttu-id="1b9a1-224">Значение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b9a1-225">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="1b9a1-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1b9a1-226">1.5</span><span class="sxs-lookup"><span data-stu-id="1b9a1-226">1.5</span></span>|
|[<span data-ttu-id="1b9a1-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1b9a1-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1b9a1-228">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b9a1-229">Пример</span><span class="sxs-lookup"><span data-stu-id="1b9a1-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="1b9a1-230">требования: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="1b9a1-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="1b9a1-231">Предоставляет метод определения, какие наборы требований поддерживаются в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="1b9a1-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="1b9a1-232">Тип</span><span class="sxs-lookup"><span data-stu-id="1b9a1-232">Type</span></span>

*   [<span data-ttu-id="1b9a1-233">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="1b9a1-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="1b9a1-234">Требования</span><span class="sxs-lookup"><span data-stu-id="1b9a1-234">Requirements</span></span>

|<span data-ttu-id="1b9a1-235">Требование</span><span class="sxs-lookup"><span data-stu-id="1b9a1-235">Requirement</span></span>| <span data-ttu-id="1b9a1-236">Значение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b9a1-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1b9a1-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1b9a1-238">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-238">1.1</span></span>|
|[<span data-ttu-id="1b9a1-239">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1b9a1-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1b9a1-240">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1b9a1-241">Пример</span><span class="sxs-lookup"><span data-stu-id="1b9a1-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="1b9a1-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="1b9a1-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="1b9a1-243">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="1b9a1-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="1b9a1-244">Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранимой в почтовом ящике пользователя, чтобы она была доступна этой надстройке, когда она запущена из любого клиента Outlook, используемого для доступа к этому `RoamingSettings` почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="1b9a1-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="1b9a1-245">Тип</span><span class="sxs-lookup"><span data-stu-id="1b9a1-245">Type</span></span>

*   [<span data-ttu-id="1b9a1-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1b9a1-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="1b9a1-247">Требования</span><span class="sxs-lookup"><span data-stu-id="1b9a1-247">Requirements</span></span>

|<span data-ttu-id="1b9a1-248">Требование</span><span class="sxs-lookup"><span data-stu-id="1b9a1-248">Requirement</span></span>| <span data-ttu-id="1b9a1-249">Значение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b9a1-250">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1b9a1-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1b9a1-251">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-251">1.1</span></span>|
|[<span data-ttu-id="1b9a1-252">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1b9a1-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="1b9a1-253">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="1b9a1-253">Restricted</span></span>|
|[<span data-ttu-id="1b9a1-254">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1b9a1-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1b9a1-255">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="1b9a1-256">ui: [пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="1b9a1-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="1b9a1-257">Предоставляет объекты и методы, которые можно использовать для создания и управления компонентами пользовательского интерфейса, такими как диалоговое окно, в Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="1b9a1-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="1b9a1-258">Тип</span><span class="sxs-lookup"><span data-stu-id="1b9a1-258">Type</span></span>

*   [<span data-ttu-id="1b9a1-259">UI</span><span class="sxs-lookup"><span data-stu-id="1b9a1-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="1b9a1-260">Требования</span><span class="sxs-lookup"><span data-stu-id="1b9a1-260">Requirements</span></span>

|<span data-ttu-id="1b9a1-261">Требование</span><span class="sxs-lookup"><span data-stu-id="1b9a1-261">Requirement</span></span>| <span data-ttu-id="1b9a1-262">Значение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="1b9a1-263">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1b9a1-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1b9a1-264">1.1</span><span class="sxs-lookup"><span data-stu-id="1b9a1-264">1.1</span></span>|
|[<span data-ttu-id="1b9a1-265">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1b9a1-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1b9a1-266">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1b9a1-266">Compose or Read</span></span>|
