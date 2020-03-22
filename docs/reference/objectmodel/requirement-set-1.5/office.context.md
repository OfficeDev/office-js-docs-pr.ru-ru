---
title: Office. Context — набор обязательных элементов 1,5
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 5b113c3ff98300cc70a98318714c794da139db3f
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891364"
---
# <a name="context-mailbox-requirement-set-15"></a><span data-ttu-id="a6d53-103">контекст (набор требований для почтового ящика 1,5)</span><span class="sxs-lookup"><span data-stu-id="a6d53-103">context (Mailbox requirement set 1.5)</span></span>

### <a name="officecontext"></a><span data-ttu-id="a6d53-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="a6d53-104">[Office](office.md).context</span></span>

<span data-ttu-id="a6d53-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="a6d53-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="a6d53-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.5).</span><span class="sxs-lookup"><span data-stu-id="a6d53-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6d53-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6d53-107">Requirements</span></span>

|<span data-ttu-id="a6d53-108">Требование</span><span class="sxs-lookup"><span data-stu-id="a6d53-108">Requirement</span></span>| <span data-ttu-id="a6d53-109">Значение</span><span class="sxs-lookup"><span data-stu-id="a6d53-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6d53-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a6d53-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a6d53-111">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-111">1.1</span></span>|
|[<span data-ttu-id="a6d53-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a6d53-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a6d53-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a6d53-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="a6d53-114">Properties</span></span>

| <span data-ttu-id="a6d53-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="a6d53-115">Property</span></span> | <span data-ttu-id="a6d53-116">Способов</span><span class="sxs-lookup"><span data-stu-id="a6d53-116">Modes</span></span> | <span data-ttu-id="a6d53-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="a6d53-117">Return type</span></span> | <span data-ttu-id="a6d53-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="a6d53-118">Minimum</span></span><br><span data-ttu-id="a6d53-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="a6d53-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a6d53-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="a6d53-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="a6d53-121">Создание</span><span class="sxs-lookup"><span data-stu-id="a6d53-121">Compose</span></span><br><span data-ttu-id="a6d53-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-122">Read</span></span> | <span data-ttu-id="a6d53-123">String</span><span class="sxs-lookup"><span data-stu-id="a6d53-123">String</span></span> | [<span data-ttu-id="a6d53-124">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a6d53-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="a6d53-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="a6d53-126">Создание</span><span class="sxs-lookup"><span data-stu-id="a6d53-126">Compose</span></span><br><span data-ttu-id="a6d53-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-127">Read</span></span> | [<span data-ttu-id="a6d53-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="a6d53-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5) | [<span data-ttu-id="a6d53-129">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a6d53-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="a6d53-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="a6d53-131">Создание</span><span class="sxs-lookup"><span data-stu-id="a6d53-131">Compose</span></span><br><span data-ttu-id="a6d53-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-132">Read</span></span> | <span data-ttu-id="a6d53-133">String</span><span class="sxs-lookup"><span data-stu-id="a6d53-133">String</span></span> | [<span data-ttu-id="a6d53-134">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a6d53-135">размещать</span><span class="sxs-lookup"><span data-stu-id="a6d53-135">host</span></span>](#host-hosttype) | <span data-ttu-id="a6d53-136">Создание</span><span class="sxs-lookup"><span data-stu-id="a6d53-136">Compose</span></span><br><span data-ttu-id="a6d53-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-137">Read</span></span> | [<span data-ttu-id="a6d53-138">HostType</span><span class="sxs-lookup"><span data-stu-id="a6d53-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5) | [<span data-ttu-id="a6d53-139">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a6d53-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="a6d53-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="a6d53-141">Создание</span><span class="sxs-lookup"><span data-stu-id="a6d53-141">Compose</span></span><br><span data-ttu-id="a6d53-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-142">Read</span></span> | [<span data-ttu-id="a6d53-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="a6d53-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5) | [<span data-ttu-id="a6d53-144">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a6d53-145">управляем</span><span class="sxs-lookup"><span data-stu-id="a6d53-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="a6d53-146">Создание</span><span class="sxs-lookup"><span data-stu-id="a6d53-146">Compose</span></span><br><span data-ttu-id="a6d53-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-147">Read</span></span> | [<span data-ttu-id="a6d53-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="a6d53-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5) | [<span data-ttu-id="a6d53-149">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a6d53-150">потребность</span><span class="sxs-lookup"><span data-stu-id="a6d53-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="a6d53-151">Создание</span><span class="sxs-lookup"><span data-stu-id="a6d53-151">Compose</span></span><br><span data-ttu-id="a6d53-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-152">Read</span></span> | [<span data-ttu-id="a6d53-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="a6d53-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5) | [<span data-ttu-id="a6d53-154">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a6d53-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="a6d53-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="a6d53-156">Создание</span><span class="sxs-lookup"><span data-stu-id="a6d53-156">Compose</span></span><br><span data-ttu-id="a6d53-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-157">Read</span></span> | [<span data-ttu-id="a6d53-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a6d53-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5) | [<span data-ttu-id="a6d53-159">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a6d53-160">ui</span><span class="sxs-lookup"><span data-stu-id="a6d53-160">ui</span></span>](#ui-ui) | <span data-ttu-id="a6d53-161">Создание</span><span class="sxs-lookup"><span data-stu-id="a6d53-161">Compose</span></span><br><span data-ttu-id="a6d53-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-162">Read</span></span> | [<span data-ttu-id="a6d53-163">UI</span><span class="sxs-lookup"><span data-stu-id="a6d53-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5) | [<span data-ttu-id="a6d53-164">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="a6d53-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="a6d53-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="a6d53-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="a6d53-166">contentLanguage: String</span></span>

<span data-ttu-id="a6d53-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="a6d53-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="a6d53-168">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="a6d53-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="a6d53-169">Тип</span><span class="sxs-lookup"><span data-stu-id="a6d53-169">Type</span></span>

*   <span data-ttu-id="a6d53-170">String</span><span class="sxs-lookup"><span data-stu-id="a6d53-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6d53-171">Требования</span><span class="sxs-lookup"><span data-stu-id="a6d53-171">Requirements</span></span>

|<span data-ttu-id="a6d53-172">Требование</span><span class="sxs-lookup"><span data-stu-id="a6d53-172">Requirement</span></span>| <span data-ttu-id="a6d53-173">Значение</span><span class="sxs-lookup"><span data-stu-id="a6d53-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6d53-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a6d53-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a6d53-175">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-175">1.1</span></span>|
|[<span data-ttu-id="a6d53-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a6d53-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a6d53-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6d53-178">Пример</span><span class="sxs-lookup"><span data-stu-id="a6d53-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="a6d53-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="a6d53-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="a6d53-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="a6d53-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="a6d53-181">Тип</span><span class="sxs-lookup"><span data-stu-id="a6d53-181">Type</span></span>

*   [<span data-ttu-id="a6d53-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="a6d53-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="a6d53-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6d53-183">Requirements</span></span>

|<span data-ttu-id="a6d53-184">Требование</span><span class="sxs-lookup"><span data-stu-id="a6d53-184">Requirement</span></span>| <span data-ttu-id="a6d53-185">Значение</span><span class="sxs-lookup"><span data-stu-id="a6d53-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6d53-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a6d53-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a6d53-187">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-187">1.1</span></span>|
|[<span data-ttu-id="a6d53-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a6d53-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a6d53-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6d53-190">Пример</span><span class="sxs-lookup"><span data-stu-id="a6d53-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="a6d53-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="a6d53-191">displayLanguage: String</span></span>

<span data-ttu-id="a6d53-192">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="a6d53-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="a6d53-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="a6d53-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="a6d53-194">Тип</span><span class="sxs-lookup"><span data-stu-id="a6d53-194">Type</span></span>

*   <span data-ttu-id="a6d53-195">String</span><span class="sxs-lookup"><span data-stu-id="a6d53-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a6d53-196">Требования</span><span class="sxs-lookup"><span data-stu-id="a6d53-196">Requirements</span></span>

|<span data-ttu-id="a6d53-197">Требование</span><span class="sxs-lookup"><span data-stu-id="a6d53-197">Requirement</span></span>| <span data-ttu-id="a6d53-198">Значение</span><span class="sxs-lookup"><span data-stu-id="a6d53-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6d53-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a6d53-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a6d53-200">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-200">1.1</span></span>|
|[<span data-ttu-id="a6d53-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a6d53-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a6d53-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6d53-203">Пример</span><span class="sxs-lookup"><span data-stu-id="a6d53-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="a6d53-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="a6d53-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="a6d53-205">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="a6d53-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="a6d53-206">Тип</span><span class="sxs-lookup"><span data-stu-id="a6d53-206">Type</span></span>

*   [<span data-ttu-id="a6d53-207">HostType</span><span class="sxs-lookup"><span data-stu-id="a6d53-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="a6d53-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6d53-208">Requirements</span></span>

|<span data-ttu-id="a6d53-209">Требование</span><span class="sxs-lookup"><span data-stu-id="a6d53-209">Requirement</span></span>| <span data-ttu-id="a6d53-210">Значение</span><span class="sxs-lookup"><span data-stu-id="a6d53-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6d53-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a6d53-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a6d53-212">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-212">1.1</span></span>|
|[<span data-ttu-id="a6d53-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a6d53-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a6d53-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6d53-215">Пример</span><span class="sxs-lookup"><span data-stu-id="a6d53-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="a6d53-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="a6d53-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="a6d53-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="a6d53-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="a6d53-218">Тип</span><span class="sxs-lookup"><span data-stu-id="a6d53-218">Type</span></span>

*   [<span data-ttu-id="a6d53-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="a6d53-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="a6d53-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6d53-220">Requirements</span></span>

|<span data-ttu-id="a6d53-221">Требование</span><span class="sxs-lookup"><span data-stu-id="a6d53-221">Requirement</span></span>| <span data-ttu-id="a6d53-222">Значение</span><span class="sxs-lookup"><span data-stu-id="a6d53-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6d53-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a6d53-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a6d53-224">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-224">1.1</span></span>|
|[<span data-ttu-id="a6d53-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a6d53-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a6d53-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6d53-227">Пример</span><span class="sxs-lookup"><span data-stu-id="a6d53-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="a6d53-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="a6d53-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="a6d53-229">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="a6d53-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="a6d53-230">Тип</span><span class="sxs-lookup"><span data-stu-id="a6d53-230">Type</span></span>

*   [<span data-ttu-id="a6d53-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="a6d53-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="a6d53-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6d53-232">Requirements</span></span>

|<span data-ttu-id="a6d53-233">Требование</span><span class="sxs-lookup"><span data-stu-id="a6d53-233">Requirement</span></span>| <span data-ttu-id="a6d53-234">Значение</span><span class="sxs-lookup"><span data-stu-id="a6d53-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6d53-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a6d53-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a6d53-236">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-236">1.1</span></span>|
|[<span data-ttu-id="a6d53-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a6d53-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a6d53-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a6d53-239">Пример</span><span class="sxs-lookup"><span data-stu-id="a6d53-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="a6d53-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="a6d53-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="a6d53-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="a6d53-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="a6d53-242">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="a6d53-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="a6d53-243">Тип</span><span class="sxs-lookup"><span data-stu-id="a6d53-243">Type</span></span>

*   [<span data-ttu-id="a6d53-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a6d53-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="a6d53-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6d53-245">Requirements</span></span>

|<span data-ttu-id="a6d53-246">Требование</span><span class="sxs-lookup"><span data-stu-id="a6d53-246">Requirement</span></span>| <span data-ttu-id="a6d53-247">Значение</span><span class="sxs-lookup"><span data-stu-id="a6d53-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6d53-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a6d53-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a6d53-249">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-249">1.1</span></span>|
|[<span data-ttu-id="a6d53-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a6d53-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="a6d53-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="a6d53-251">Restricted</span></span>|
|[<span data-ttu-id="a6d53-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a6d53-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a6d53-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="a6d53-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="a6d53-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="a6d53-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="a6d53-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="a6d53-256">Тип</span><span class="sxs-lookup"><span data-stu-id="a6d53-256">Type</span></span>

*   [<span data-ttu-id="a6d53-257">UI</span><span class="sxs-lookup"><span data-stu-id="a6d53-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="a6d53-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="a6d53-258">Requirements</span></span>

|<span data-ttu-id="a6d53-259">Требование</span><span class="sxs-lookup"><span data-stu-id="a6d53-259">Requirement</span></span>| <span data-ttu-id="a6d53-260">Значение</span><span class="sxs-lookup"><span data-stu-id="a6d53-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="a6d53-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a6d53-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a6d53-262">1.1</span><span class="sxs-lookup"><span data-stu-id="a6d53-262">1.1</span></span>|
|[<span data-ttu-id="a6d53-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a6d53-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a6d53-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a6d53-264">Compose or Read</span></span>|
