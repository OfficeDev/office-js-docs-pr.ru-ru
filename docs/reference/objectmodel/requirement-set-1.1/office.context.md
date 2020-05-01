---
title: Office. Context — набор обязательных элементов 1,1
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,1.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: dac630092d3b15cff0c081102e452d2c698c5533
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890832"
---
# <a name="context-mailbox-requirement-set-11"></a><span data-ttu-id="311af-103">контекст (набор требований для почтового ящика 1,1)</span><span class="sxs-lookup"><span data-stu-id="311af-103">context (Mailbox requirement set 1.1)</span></span>

### <a name="officecontext"></a><span data-ttu-id="311af-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="311af-104">[Office](office.md).context</span></span>

<span data-ttu-id="311af-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="311af-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="311af-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.1).</span><span class="sxs-lookup"><span data-stu-id="311af-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.1).</span></span>

##### <a name="requirements"></a><span data-ttu-id="311af-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="311af-107">Requirements</span></span>

|<span data-ttu-id="311af-108">Требование</span><span class="sxs-lookup"><span data-stu-id="311af-108">Requirement</span></span>| <span data-ttu-id="311af-109">Значение</span><span class="sxs-lookup"><span data-stu-id="311af-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="311af-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="311af-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="311af-111">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-111">1.1</span></span>|
|[<span data-ttu-id="311af-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="311af-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="311af-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="311af-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="311af-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="311af-114">Properties</span></span>

| <span data-ttu-id="311af-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="311af-115">Property</span></span> | <span data-ttu-id="311af-116">Способов</span><span class="sxs-lookup"><span data-stu-id="311af-116">Modes</span></span> | <span data-ttu-id="311af-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="311af-117">Return type</span></span> | <span data-ttu-id="311af-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="311af-118">Minimum</span></span><br><span data-ttu-id="311af-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="311af-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="311af-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="311af-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="311af-121">Создание</span><span class="sxs-lookup"><span data-stu-id="311af-121">Compose</span></span><br><span data-ttu-id="311af-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="311af-122">Read</span></span> | <span data-ttu-id="311af-123">Строка</span><span class="sxs-lookup"><span data-stu-id="311af-123">String</span></span> | [<span data-ttu-id="311af-124">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="311af-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="311af-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="311af-126">Создание</span><span class="sxs-lookup"><span data-stu-id="311af-126">Compose</span></span><br><span data-ttu-id="311af-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="311af-127">Read</span></span> | [<span data-ttu-id="311af-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="311af-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.1) | [<span data-ttu-id="311af-129">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="311af-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="311af-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="311af-131">Создание</span><span class="sxs-lookup"><span data-stu-id="311af-131">Compose</span></span><br><span data-ttu-id="311af-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="311af-132">Read</span></span> | <span data-ttu-id="311af-133">Строка</span><span class="sxs-lookup"><span data-stu-id="311af-133">String</span></span> | [<span data-ttu-id="311af-134">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="311af-135">размещать</span><span class="sxs-lookup"><span data-stu-id="311af-135">host</span></span>](#host-hosttype) | <span data-ttu-id="311af-136">Создание</span><span class="sxs-lookup"><span data-stu-id="311af-136">Compose</span></span><br><span data-ttu-id="311af-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="311af-137">Read</span></span> | [<span data-ttu-id="311af-138">HostType</span><span class="sxs-lookup"><span data-stu-id="311af-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.1) | [<span data-ttu-id="311af-139">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="311af-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="311af-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="311af-141">Создание</span><span class="sxs-lookup"><span data-stu-id="311af-141">Compose</span></span><br><span data-ttu-id="311af-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="311af-142">Read</span></span> | [<span data-ttu-id="311af-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="311af-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.1) | [<span data-ttu-id="311af-144">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="311af-145">управляем</span><span class="sxs-lookup"><span data-stu-id="311af-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="311af-146">Создание</span><span class="sxs-lookup"><span data-stu-id="311af-146">Compose</span></span><br><span data-ttu-id="311af-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="311af-147">Read</span></span> | [<span data-ttu-id="311af-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="311af-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.1) | [<span data-ttu-id="311af-149">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="311af-150">потребность</span><span class="sxs-lookup"><span data-stu-id="311af-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="311af-151">Создание</span><span class="sxs-lookup"><span data-stu-id="311af-151">Compose</span></span><br><span data-ttu-id="311af-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="311af-152">Read</span></span> | [<span data-ttu-id="311af-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="311af-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.1) | [<span data-ttu-id="311af-154">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="311af-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="311af-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="311af-156">Создание</span><span class="sxs-lookup"><span data-stu-id="311af-156">Compose</span></span><br><span data-ttu-id="311af-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="311af-157">Read</span></span> | [<span data-ttu-id="311af-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="311af-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.1) | [<span data-ttu-id="311af-159">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="311af-160">ui</span><span class="sxs-lookup"><span data-stu-id="311af-160">ui</span></span>](#ui-ui) | <span data-ttu-id="311af-161">Создание</span><span class="sxs-lookup"><span data-stu-id="311af-161">Compose</span></span><br><span data-ttu-id="311af-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="311af-162">Read</span></span> | [<span data-ttu-id="311af-163">UI</span><span class="sxs-lookup"><span data-stu-id="311af-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.1) | [<span data-ttu-id="311af-164">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="311af-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="311af-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="311af-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="311af-166">contentLanguage: String</span></span>

<span data-ttu-id="311af-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="311af-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="311af-168">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="311af-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="311af-169">Тип</span><span class="sxs-lookup"><span data-stu-id="311af-169">Type</span></span>

*   <span data-ttu-id="311af-170">String</span><span class="sxs-lookup"><span data-stu-id="311af-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="311af-171">Требования</span><span class="sxs-lookup"><span data-stu-id="311af-171">Requirements</span></span>

|<span data-ttu-id="311af-172">Требование</span><span class="sxs-lookup"><span data-stu-id="311af-172">Requirement</span></span>| <span data-ttu-id="311af-173">Значение</span><span class="sxs-lookup"><span data-stu-id="311af-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="311af-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="311af-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="311af-175">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-175">1.1</span></span>|
|[<span data-ttu-id="311af-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="311af-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="311af-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="311af-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="311af-178">Пример</span><span class="sxs-lookup"><span data-stu-id="311af-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="311af-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="311af-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="311af-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="311af-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="311af-181">Type</span><span class="sxs-lookup"><span data-stu-id="311af-181">Type</span></span>

*   [<span data-ttu-id="311af-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="311af-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="311af-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="311af-183">Requirements</span></span>

|<span data-ttu-id="311af-184">Требование</span><span class="sxs-lookup"><span data-stu-id="311af-184">Requirement</span></span>| <span data-ttu-id="311af-185">Значение</span><span class="sxs-lookup"><span data-stu-id="311af-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="311af-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="311af-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="311af-187">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-187">1.1</span></span>|
|[<span data-ttu-id="311af-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="311af-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="311af-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="311af-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="311af-190">Пример</span><span class="sxs-lookup"><span data-stu-id="311af-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="311af-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="311af-191">displayLanguage: String</span></span>

<span data-ttu-id="311af-192">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="311af-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="311af-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="311af-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="311af-194">Тип</span><span class="sxs-lookup"><span data-stu-id="311af-194">Type</span></span>

*   <span data-ttu-id="311af-195">String</span><span class="sxs-lookup"><span data-stu-id="311af-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="311af-196">Требования</span><span class="sxs-lookup"><span data-stu-id="311af-196">Requirements</span></span>

|<span data-ttu-id="311af-197">Требование</span><span class="sxs-lookup"><span data-stu-id="311af-197">Requirement</span></span>| <span data-ttu-id="311af-198">Значение</span><span class="sxs-lookup"><span data-stu-id="311af-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="311af-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="311af-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="311af-200">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-200">1.1</span></span>|
|[<span data-ttu-id="311af-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="311af-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="311af-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="311af-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="311af-203">Пример</span><span class="sxs-lookup"><span data-stu-id="311af-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="311af-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="311af-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="311af-205">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="311af-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="311af-206">Type</span><span class="sxs-lookup"><span data-stu-id="311af-206">Type</span></span>

*   [<span data-ttu-id="311af-207">HostType</span><span class="sxs-lookup"><span data-stu-id="311af-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="311af-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="311af-208">Requirements</span></span>

|<span data-ttu-id="311af-209">Требование</span><span class="sxs-lookup"><span data-stu-id="311af-209">Requirement</span></span>| <span data-ttu-id="311af-210">Значение</span><span class="sxs-lookup"><span data-stu-id="311af-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="311af-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="311af-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="311af-212">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-212">1.1</span></span>|
|[<span data-ttu-id="311af-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="311af-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="311af-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="311af-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="311af-215">Пример</span><span class="sxs-lookup"><span data-stu-id="311af-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="311af-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="311af-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="311af-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="311af-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="311af-218">Type</span><span class="sxs-lookup"><span data-stu-id="311af-218">Type</span></span>

*   [<span data-ttu-id="311af-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="311af-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="311af-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="311af-220">Requirements</span></span>

|<span data-ttu-id="311af-221">Требование</span><span class="sxs-lookup"><span data-stu-id="311af-221">Requirement</span></span>| <span data-ttu-id="311af-222">Значение</span><span class="sxs-lookup"><span data-stu-id="311af-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="311af-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="311af-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="311af-224">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-224">1.1</span></span>|
|[<span data-ttu-id="311af-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="311af-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="311af-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="311af-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="311af-227">Пример</span><span class="sxs-lookup"><span data-stu-id="311af-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="311af-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="311af-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="311af-229">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="311af-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="311af-230">Type</span><span class="sxs-lookup"><span data-stu-id="311af-230">Type</span></span>

*   [<span data-ttu-id="311af-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="311af-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="311af-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="311af-232">Requirements</span></span>

|<span data-ttu-id="311af-233">Требование</span><span class="sxs-lookup"><span data-stu-id="311af-233">Requirement</span></span>| <span data-ttu-id="311af-234">Значение</span><span class="sxs-lookup"><span data-stu-id="311af-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="311af-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="311af-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="311af-236">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-236">1.1</span></span>|
|[<span data-ttu-id="311af-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="311af-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="311af-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="311af-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="311af-239">Пример</span><span class="sxs-lookup"><span data-stu-id="311af-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="311af-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="311af-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="311af-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="311af-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="311af-242">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="311af-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="311af-243">Type</span><span class="sxs-lookup"><span data-stu-id="311af-243">Type</span></span>

*   [<span data-ttu-id="311af-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="311af-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="311af-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="311af-245">Requirements</span></span>

|<span data-ttu-id="311af-246">Требование</span><span class="sxs-lookup"><span data-stu-id="311af-246">Requirement</span></span>| <span data-ttu-id="311af-247">Значение</span><span class="sxs-lookup"><span data-stu-id="311af-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="311af-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="311af-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="311af-249">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-249">1.1</span></span>|
|[<span data-ttu-id="311af-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="311af-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="311af-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="311af-251">Restricted</span></span>|
|[<span data-ttu-id="311af-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="311af-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="311af-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="311af-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="311af-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="311af-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="311af-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="311af-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="311af-256">Type</span><span class="sxs-lookup"><span data-stu-id="311af-256">Type</span></span>

*   [<span data-ttu-id="311af-257">UI</span><span class="sxs-lookup"><span data-stu-id="311af-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="311af-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="311af-258">Requirements</span></span>

|<span data-ttu-id="311af-259">Требование</span><span class="sxs-lookup"><span data-stu-id="311af-259">Requirement</span></span>| <span data-ttu-id="311af-260">Значение</span><span class="sxs-lookup"><span data-stu-id="311af-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="311af-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="311af-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="311af-262">1.1</span><span class="sxs-lookup"><span data-stu-id="311af-262">1.1</span></span>|
|[<span data-ttu-id="311af-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="311af-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="311af-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="311af-264">Compose or Read</span></span>|
