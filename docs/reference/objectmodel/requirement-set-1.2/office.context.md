---
title: Office. Context — набор обязательных элементов 1,2
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,2.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 15172ac9b42455a5b087f48c368d210596922589
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890727"
---
# <a name="context-mailbox-requirement-set-12"></a><span data-ttu-id="82878-103">контекст (набор требований для почтового ящика 1,2)</span><span class="sxs-lookup"><span data-stu-id="82878-103">context (Mailbox requirement set 1.2)</span></span>

### <a name="officecontext"></a><span data-ttu-id="82878-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="82878-104">[Office](office.md).context</span></span>

<span data-ttu-id="82878-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="82878-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="82878-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.2).</span><span class="sxs-lookup"><span data-stu-id="82878-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.2).</span></span>

##### <a name="requirements"></a><span data-ttu-id="82878-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="82878-107">Requirements</span></span>

|<span data-ttu-id="82878-108">Требование</span><span class="sxs-lookup"><span data-stu-id="82878-108">Requirement</span></span>| <span data-ttu-id="82878-109">Значение</span><span class="sxs-lookup"><span data-stu-id="82878-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="82878-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="82878-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="82878-111">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-111">1.1</span></span>|
|[<span data-ttu-id="82878-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="82878-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="82878-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="82878-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="82878-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="82878-114">Properties</span></span>

| <span data-ttu-id="82878-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="82878-115">Property</span></span> | <span data-ttu-id="82878-116">Способов</span><span class="sxs-lookup"><span data-stu-id="82878-116">Modes</span></span> | <span data-ttu-id="82878-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="82878-117">Return type</span></span> | <span data-ttu-id="82878-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="82878-118">Minimum</span></span><br><span data-ttu-id="82878-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="82878-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="82878-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="82878-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="82878-121">Создание</span><span class="sxs-lookup"><span data-stu-id="82878-121">Compose</span></span><br><span data-ttu-id="82878-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="82878-122">Read</span></span> | <span data-ttu-id="82878-123">String</span><span class="sxs-lookup"><span data-stu-id="82878-123">String</span></span> | [<span data-ttu-id="82878-124">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="82878-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="82878-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="82878-126">Создание</span><span class="sxs-lookup"><span data-stu-id="82878-126">Compose</span></span><br><span data-ttu-id="82878-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="82878-127">Read</span></span> | [<span data-ttu-id="82878-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="82878-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.2) | [<span data-ttu-id="82878-129">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="82878-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="82878-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="82878-131">Создание</span><span class="sxs-lookup"><span data-stu-id="82878-131">Compose</span></span><br><span data-ttu-id="82878-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="82878-132">Read</span></span> | <span data-ttu-id="82878-133">String</span><span class="sxs-lookup"><span data-stu-id="82878-133">String</span></span> | [<span data-ttu-id="82878-134">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="82878-135">размещать</span><span class="sxs-lookup"><span data-stu-id="82878-135">host</span></span>](#host-hosttype) | <span data-ttu-id="82878-136">Создание</span><span class="sxs-lookup"><span data-stu-id="82878-136">Compose</span></span><br><span data-ttu-id="82878-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="82878-137">Read</span></span> | [<span data-ttu-id="82878-138">HostType</span><span class="sxs-lookup"><span data-stu-id="82878-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.2) | [<span data-ttu-id="82878-139">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="82878-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="82878-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="82878-141">Создание</span><span class="sxs-lookup"><span data-stu-id="82878-141">Compose</span></span><br><span data-ttu-id="82878-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="82878-142">Read</span></span> | [<span data-ttu-id="82878-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="82878-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2) | [<span data-ttu-id="82878-144">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="82878-145">управляем</span><span class="sxs-lookup"><span data-stu-id="82878-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="82878-146">Создание</span><span class="sxs-lookup"><span data-stu-id="82878-146">Compose</span></span><br><span data-ttu-id="82878-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="82878-147">Read</span></span> | [<span data-ttu-id="82878-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="82878-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.2) | [<span data-ttu-id="82878-149">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="82878-150">потребность</span><span class="sxs-lookup"><span data-stu-id="82878-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="82878-151">Создание</span><span class="sxs-lookup"><span data-stu-id="82878-151">Compose</span></span><br><span data-ttu-id="82878-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="82878-152">Read</span></span> | [<span data-ttu-id="82878-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="82878-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.2) | [<span data-ttu-id="82878-154">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="82878-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="82878-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="82878-156">Создание</span><span class="sxs-lookup"><span data-stu-id="82878-156">Compose</span></span><br><span data-ttu-id="82878-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="82878-157">Read</span></span> | [<span data-ttu-id="82878-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="82878-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.2) | [<span data-ttu-id="82878-159">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="82878-160">ui</span><span class="sxs-lookup"><span data-stu-id="82878-160">ui</span></span>](#ui-ui) | <span data-ttu-id="82878-161">Создание</span><span class="sxs-lookup"><span data-stu-id="82878-161">Compose</span></span><br><span data-ttu-id="82878-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="82878-162">Read</span></span> | [<span data-ttu-id="82878-163">UI</span><span class="sxs-lookup"><span data-stu-id="82878-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.2) | [<span data-ttu-id="82878-164">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="82878-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="82878-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="82878-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="82878-166">contentLanguage: String</span></span>

<span data-ttu-id="82878-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="82878-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="82878-168">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="82878-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="82878-169">Тип</span><span class="sxs-lookup"><span data-stu-id="82878-169">Type</span></span>

*   <span data-ttu-id="82878-170">String</span><span class="sxs-lookup"><span data-stu-id="82878-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="82878-171">Требования</span><span class="sxs-lookup"><span data-stu-id="82878-171">Requirements</span></span>

|<span data-ttu-id="82878-172">Требование</span><span class="sxs-lookup"><span data-stu-id="82878-172">Requirement</span></span>| <span data-ttu-id="82878-173">Значение</span><span class="sxs-lookup"><span data-stu-id="82878-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="82878-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="82878-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="82878-175">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-175">1.1</span></span>|
|[<span data-ttu-id="82878-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="82878-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="82878-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="82878-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="82878-178">Пример</span><span class="sxs-lookup"><span data-stu-id="82878-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="82878-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="82878-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="82878-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="82878-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="82878-181">Тип</span><span class="sxs-lookup"><span data-stu-id="82878-181">Type</span></span>

*   [<span data-ttu-id="82878-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="82878-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="82878-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="82878-183">Requirements</span></span>

|<span data-ttu-id="82878-184">Требование</span><span class="sxs-lookup"><span data-stu-id="82878-184">Requirement</span></span>| <span data-ttu-id="82878-185">Значение</span><span class="sxs-lookup"><span data-stu-id="82878-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="82878-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="82878-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="82878-187">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-187">1.1</span></span>|
|[<span data-ttu-id="82878-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="82878-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="82878-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="82878-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="82878-190">Пример</span><span class="sxs-lookup"><span data-stu-id="82878-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="82878-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="82878-191">displayLanguage: String</span></span>

<span data-ttu-id="82878-192">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="82878-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="82878-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="82878-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="82878-194">Тип</span><span class="sxs-lookup"><span data-stu-id="82878-194">Type</span></span>

*   <span data-ttu-id="82878-195">String</span><span class="sxs-lookup"><span data-stu-id="82878-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="82878-196">Требования</span><span class="sxs-lookup"><span data-stu-id="82878-196">Requirements</span></span>

|<span data-ttu-id="82878-197">Требование</span><span class="sxs-lookup"><span data-stu-id="82878-197">Requirement</span></span>| <span data-ttu-id="82878-198">Значение</span><span class="sxs-lookup"><span data-stu-id="82878-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="82878-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="82878-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="82878-200">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-200">1.1</span></span>|
|[<span data-ttu-id="82878-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="82878-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="82878-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="82878-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="82878-203">Пример</span><span class="sxs-lookup"><span data-stu-id="82878-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="82878-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="82878-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="82878-205">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="82878-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="82878-206">Тип</span><span class="sxs-lookup"><span data-stu-id="82878-206">Type</span></span>

*   [<span data-ttu-id="82878-207">HostType</span><span class="sxs-lookup"><span data-stu-id="82878-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="82878-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="82878-208">Requirements</span></span>

|<span data-ttu-id="82878-209">Требование</span><span class="sxs-lookup"><span data-stu-id="82878-209">Requirement</span></span>| <span data-ttu-id="82878-210">Значение</span><span class="sxs-lookup"><span data-stu-id="82878-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="82878-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="82878-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="82878-212">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-212">1.1</span></span>|
|[<span data-ttu-id="82878-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="82878-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="82878-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="82878-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="82878-215">Пример</span><span class="sxs-lookup"><span data-stu-id="82878-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="82878-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="82878-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="82878-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="82878-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="82878-218">Тип</span><span class="sxs-lookup"><span data-stu-id="82878-218">Type</span></span>

*   [<span data-ttu-id="82878-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="82878-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="82878-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="82878-220">Requirements</span></span>

|<span data-ttu-id="82878-221">Требование</span><span class="sxs-lookup"><span data-stu-id="82878-221">Requirement</span></span>| <span data-ttu-id="82878-222">Значение</span><span class="sxs-lookup"><span data-stu-id="82878-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="82878-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="82878-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="82878-224">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-224">1.1</span></span>|
|[<span data-ttu-id="82878-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="82878-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="82878-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="82878-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="82878-227">Пример</span><span class="sxs-lookup"><span data-stu-id="82878-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="82878-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="82878-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="82878-229">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="82878-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="82878-230">Тип</span><span class="sxs-lookup"><span data-stu-id="82878-230">Type</span></span>

*   [<span data-ttu-id="82878-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="82878-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="82878-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="82878-232">Requirements</span></span>

|<span data-ttu-id="82878-233">Требование</span><span class="sxs-lookup"><span data-stu-id="82878-233">Requirement</span></span>| <span data-ttu-id="82878-234">Значение</span><span class="sxs-lookup"><span data-stu-id="82878-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="82878-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="82878-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="82878-236">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-236">1.1</span></span>|
|[<span data-ttu-id="82878-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="82878-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="82878-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="82878-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="82878-239">Пример</span><span class="sxs-lookup"><span data-stu-id="82878-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="82878-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="82878-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="82878-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="82878-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="82878-242">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="82878-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="82878-243">Тип</span><span class="sxs-lookup"><span data-stu-id="82878-243">Type</span></span>

*   [<span data-ttu-id="82878-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="82878-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="82878-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="82878-245">Requirements</span></span>

|<span data-ttu-id="82878-246">Требование</span><span class="sxs-lookup"><span data-stu-id="82878-246">Requirement</span></span>| <span data-ttu-id="82878-247">Значение</span><span class="sxs-lookup"><span data-stu-id="82878-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="82878-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="82878-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="82878-249">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-249">1.1</span></span>|
|[<span data-ttu-id="82878-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="82878-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="82878-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="82878-251">Restricted</span></span>|
|[<span data-ttu-id="82878-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="82878-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="82878-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="82878-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="82878-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="82878-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="82878-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="82878-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="82878-256">Тип</span><span class="sxs-lookup"><span data-stu-id="82878-256">Type</span></span>

*   [<span data-ttu-id="82878-257">UI</span><span class="sxs-lookup"><span data-stu-id="82878-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="82878-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="82878-258">Requirements</span></span>

|<span data-ttu-id="82878-259">Требование</span><span class="sxs-lookup"><span data-stu-id="82878-259">Requirement</span></span>| <span data-ttu-id="82878-260">Значение</span><span class="sxs-lookup"><span data-stu-id="82878-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="82878-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="82878-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="82878-262">1.1</span><span class="sxs-lookup"><span data-stu-id="82878-262">1.1</span></span>|
|[<span data-ttu-id="82878-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="82878-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="82878-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="82878-264">Compose or Read</span></span>|
