---
title: Office. Context — набор обязательных элементов 1,6
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,6.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: fe36815be5e66b2a0ec5557f55489433c12dd5de
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42891322"
---
# <a name="context-mailbox-requirement-set-16"></a><span data-ttu-id="6fb44-103">контекст (набор требований для почтового ящика 1,6)</span><span class="sxs-lookup"><span data-stu-id="6fb44-103">context (Mailbox requirement set 1.6)</span></span>

### <a name="officecontext"></a><span data-ttu-id="6fb44-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="6fb44-104">[Office](office.md).context</span></span>

<span data-ttu-id="6fb44-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="6fb44-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="6fb44-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.6).</span><span class="sxs-lookup"><span data-stu-id="6fb44-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fb44-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="6fb44-107">Requirements</span></span>

|<span data-ttu-id="6fb44-108">Требование</span><span class="sxs-lookup"><span data-stu-id="6fb44-108">Requirement</span></span>| <span data-ttu-id="6fb44-109">Значение</span><span class="sxs-lookup"><span data-stu-id="6fb44-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fb44-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6fb44-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fb44-111">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-111">1.1</span></span>|
|[<span data-ttu-id="6fb44-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6fb44-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fb44-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="6fb44-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="6fb44-114">Properties</span></span>

| <span data-ttu-id="6fb44-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="6fb44-115">Property</span></span> | <span data-ttu-id="6fb44-116">Способов</span><span class="sxs-lookup"><span data-stu-id="6fb44-116">Modes</span></span> | <span data-ttu-id="6fb44-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="6fb44-117">Return type</span></span> | <span data-ttu-id="6fb44-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="6fb44-118">Minimum</span></span><br><span data-ttu-id="6fb44-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="6fb44-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6fb44-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="6fb44-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="6fb44-121">Создание</span><span class="sxs-lookup"><span data-stu-id="6fb44-121">Compose</span></span><br><span data-ttu-id="6fb44-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-122">Read</span></span> | <span data-ttu-id="6fb44-123">String</span><span class="sxs-lookup"><span data-stu-id="6fb44-123">String</span></span> | [<span data-ttu-id="6fb44-124">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6fb44-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="6fb44-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="6fb44-126">Создание</span><span class="sxs-lookup"><span data-stu-id="6fb44-126">Compose</span></span><br><span data-ttu-id="6fb44-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-127">Read</span></span> | [<span data-ttu-id="6fb44-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="6fb44-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6) | [<span data-ttu-id="6fb44-129">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6fb44-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="6fb44-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="6fb44-131">Создание</span><span class="sxs-lookup"><span data-stu-id="6fb44-131">Compose</span></span><br><span data-ttu-id="6fb44-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-132">Read</span></span> | <span data-ttu-id="6fb44-133">String</span><span class="sxs-lookup"><span data-stu-id="6fb44-133">String</span></span> | [<span data-ttu-id="6fb44-134">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6fb44-135">размещать</span><span class="sxs-lookup"><span data-stu-id="6fb44-135">host</span></span>](#host-hosttype) | <span data-ttu-id="6fb44-136">Создание</span><span class="sxs-lookup"><span data-stu-id="6fb44-136">Compose</span></span><br><span data-ttu-id="6fb44-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-137">Read</span></span> | [<span data-ttu-id="6fb44-138">HostType</span><span class="sxs-lookup"><span data-stu-id="6fb44-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6) | [<span data-ttu-id="6fb44-139">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6fb44-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="6fb44-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="6fb44-141">Создание</span><span class="sxs-lookup"><span data-stu-id="6fb44-141">Compose</span></span><br><span data-ttu-id="6fb44-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-142">Read</span></span> | [<span data-ttu-id="6fb44-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="6fb44-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6) | [<span data-ttu-id="6fb44-144">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6fb44-145">управляем</span><span class="sxs-lookup"><span data-stu-id="6fb44-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="6fb44-146">Создание</span><span class="sxs-lookup"><span data-stu-id="6fb44-146">Compose</span></span><br><span data-ttu-id="6fb44-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-147">Read</span></span> | [<span data-ttu-id="6fb44-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="6fb44-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6) | [<span data-ttu-id="6fb44-149">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6fb44-150">потребность</span><span class="sxs-lookup"><span data-stu-id="6fb44-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="6fb44-151">Создание</span><span class="sxs-lookup"><span data-stu-id="6fb44-151">Compose</span></span><br><span data-ttu-id="6fb44-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-152">Read</span></span> | [<span data-ttu-id="6fb44-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="6fb44-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6) | [<span data-ttu-id="6fb44-154">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6fb44-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="6fb44-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="6fb44-156">Создание</span><span class="sxs-lookup"><span data-stu-id="6fb44-156">Compose</span></span><br><span data-ttu-id="6fb44-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-157">Read</span></span> | [<span data-ttu-id="6fb44-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6fb44-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6) | [<span data-ttu-id="6fb44-159">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6fb44-160">ui</span><span class="sxs-lookup"><span data-stu-id="6fb44-160">ui</span></span>](#ui-ui) | <span data-ttu-id="6fb44-161">Создание</span><span class="sxs-lookup"><span data-stu-id="6fb44-161">Compose</span></span><br><span data-ttu-id="6fb44-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-162">Read</span></span> | [<span data-ttu-id="6fb44-163">UI</span><span class="sxs-lookup"><span data-stu-id="6fb44-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6) | [<span data-ttu-id="6fb44-164">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="6fb44-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="6fb44-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="6fb44-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="6fb44-166">contentLanguage: String</span></span>

<span data-ttu-id="6fb44-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="6fb44-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="6fb44-168">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="6fb44-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="6fb44-169">Тип</span><span class="sxs-lookup"><span data-stu-id="6fb44-169">Type</span></span>

*   <span data-ttu-id="6fb44-170">String</span><span class="sxs-lookup"><span data-stu-id="6fb44-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fb44-171">Требования</span><span class="sxs-lookup"><span data-stu-id="6fb44-171">Requirements</span></span>

|<span data-ttu-id="6fb44-172">Требование</span><span class="sxs-lookup"><span data-stu-id="6fb44-172">Requirement</span></span>| <span data-ttu-id="6fb44-173">Значение</span><span class="sxs-lookup"><span data-stu-id="6fb44-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fb44-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6fb44-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fb44-175">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-175">1.1</span></span>|
|[<span data-ttu-id="6fb44-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6fb44-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fb44-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fb44-178">Пример</span><span class="sxs-lookup"><span data-stu-id="6fb44-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="6fb44-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="6fb44-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="6fb44-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="6fb44-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="6fb44-181">Тип</span><span class="sxs-lookup"><span data-stu-id="6fb44-181">Type</span></span>

*   [<span data-ttu-id="6fb44-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="6fb44-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="6fb44-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="6fb44-183">Requirements</span></span>

|<span data-ttu-id="6fb44-184">Требование</span><span class="sxs-lookup"><span data-stu-id="6fb44-184">Requirement</span></span>| <span data-ttu-id="6fb44-185">Значение</span><span class="sxs-lookup"><span data-stu-id="6fb44-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fb44-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6fb44-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fb44-187">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-187">1.1</span></span>|
|[<span data-ttu-id="6fb44-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6fb44-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fb44-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fb44-190">Пример</span><span class="sxs-lookup"><span data-stu-id="6fb44-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="6fb44-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="6fb44-191">displayLanguage: String</span></span>

<span data-ttu-id="6fb44-192">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="6fb44-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="6fb44-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="6fb44-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="6fb44-194">Тип</span><span class="sxs-lookup"><span data-stu-id="6fb44-194">Type</span></span>

*   <span data-ttu-id="6fb44-195">String</span><span class="sxs-lookup"><span data-stu-id="6fb44-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6fb44-196">Требования</span><span class="sxs-lookup"><span data-stu-id="6fb44-196">Requirements</span></span>

|<span data-ttu-id="6fb44-197">Требование</span><span class="sxs-lookup"><span data-stu-id="6fb44-197">Requirement</span></span>| <span data-ttu-id="6fb44-198">Значение</span><span class="sxs-lookup"><span data-stu-id="6fb44-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fb44-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6fb44-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fb44-200">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-200">1.1</span></span>|
|[<span data-ttu-id="6fb44-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6fb44-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fb44-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fb44-203">Пример</span><span class="sxs-lookup"><span data-stu-id="6fb44-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="6fb44-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="6fb44-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="6fb44-205">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="6fb44-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="6fb44-206">Тип</span><span class="sxs-lookup"><span data-stu-id="6fb44-206">Type</span></span>

*   [<span data-ttu-id="6fb44-207">HostType</span><span class="sxs-lookup"><span data-stu-id="6fb44-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="6fb44-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="6fb44-208">Requirements</span></span>

|<span data-ttu-id="6fb44-209">Требование</span><span class="sxs-lookup"><span data-stu-id="6fb44-209">Requirement</span></span>| <span data-ttu-id="6fb44-210">Значение</span><span class="sxs-lookup"><span data-stu-id="6fb44-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fb44-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6fb44-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fb44-212">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-212">1.1</span></span>|
|[<span data-ttu-id="6fb44-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6fb44-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fb44-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fb44-215">Пример</span><span class="sxs-lookup"><span data-stu-id="6fb44-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="6fb44-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="6fb44-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="6fb44-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="6fb44-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="6fb44-218">Тип</span><span class="sxs-lookup"><span data-stu-id="6fb44-218">Type</span></span>

*   [<span data-ttu-id="6fb44-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="6fb44-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="6fb44-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="6fb44-220">Requirements</span></span>

|<span data-ttu-id="6fb44-221">Требование</span><span class="sxs-lookup"><span data-stu-id="6fb44-221">Requirement</span></span>| <span data-ttu-id="6fb44-222">Значение</span><span class="sxs-lookup"><span data-stu-id="6fb44-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fb44-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6fb44-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fb44-224">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-224">1.1</span></span>|
|[<span data-ttu-id="6fb44-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6fb44-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fb44-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fb44-227">Пример</span><span class="sxs-lookup"><span data-stu-id="6fb44-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="6fb44-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="6fb44-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="6fb44-229">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="6fb44-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="6fb44-230">Тип</span><span class="sxs-lookup"><span data-stu-id="6fb44-230">Type</span></span>

*   [<span data-ttu-id="6fb44-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="6fb44-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="6fb44-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="6fb44-232">Requirements</span></span>

|<span data-ttu-id="6fb44-233">Требование</span><span class="sxs-lookup"><span data-stu-id="6fb44-233">Requirement</span></span>| <span data-ttu-id="6fb44-234">Значение</span><span class="sxs-lookup"><span data-stu-id="6fb44-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fb44-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6fb44-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fb44-236">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-236">1.1</span></span>|
|[<span data-ttu-id="6fb44-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6fb44-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fb44-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6fb44-239">Пример</span><span class="sxs-lookup"><span data-stu-id="6fb44-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="6fb44-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="6fb44-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="6fb44-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="6fb44-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="6fb44-242">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="6fb44-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="6fb44-243">Тип</span><span class="sxs-lookup"><span data-stu-id="6fb44-243">Type</span></span>

*   [<span data-ttu-id="6fb44-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6fb44-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="6fb44-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="6fb44-245">Requirements</span></span>

|<span data-ttu-id="6fb44-246">Требование</span><span class="sxs-lookup"><span data-stu-id="6fb44-246">Requirement</span></span>| <span data-ttu-id="6fb44-247">Значение</span><span class="sxs-lookup"><span data-stu-id="6fb44-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fb44-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6fb44-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fb44-249">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-249">1.1</span></span>|
|[<span data-ttu-id="6fb44-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6fb44-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="6fb44-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="6fb44-251">Restricted</span></span>|
|[<span data-ttu-id="6fb44-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6fb44-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fb44-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="6fb44-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="6fb44-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="6fb44-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="6fb44-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="6fb44-256">Тип</span><span class="sxs-lookup"><span data-stu-id="6fb44-256">Type</span></span>

*   [<span data-ttu-id="6fb44-257">UI</span><span class="sxs-lookup"><span data-stu-id="6fb44-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="6fb44-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="6fb44-258">Requirements</span></span>

|<span data-ttu-id="6fb44-259">Требование</span><span class="sxs-lookup"><span data-stu-id="6fb44-259">Requirement</span></span>| <span data-ttu-id="6fb44-260">Значение</span><span class="sxs-lookup"><span data-stu-id="6fb44-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="6fb44-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6fb44-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6fb44-262">1.1</span><span class="sxs-lookup"><span data-stu-id="6fb44-262">1.1</span></span>|
|[<span data-ttu-id="6fb44-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6fb44-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6fb44-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6fb44-264">Compose or Read</span></span>|
