---
title: Office. Context — набор обязательных элементов 1,4
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,4.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 5d15e3bd70d95693eecab2d124ef43e6bf492446
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44609924"
---
# <a name="context-mailbox-requirement-set-14"></a><span data-ttu-id="eb4ce-103">контекст (набор требований для почтового ящика 1,4)</span><span class="sxs-lookup"><span data-stu-id="eb4ce-103">context (Mailbox requirement set 1.4)</span></span>

### <a name="officecontext"></a><span data-ttu-id="eb4ce-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="eb4ce-104">[Office](office.md).context</span></span>

<span data-ttu-id="eb4ce-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="eb4ce-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.4).</span><span class="sxs-lookup"><span data-stu-id="eb4ce-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4).</span></span>

##### <a name="requirements"></a><span data-ttu-id="eb4ce-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="eb4ce-107">Requirements</span></span>

|<span data-ttu-id="eb4ce-108">Требование</span><span class="sxs-lookup"><span data-stu-id="eb4ce-108">Requirement</span></span>| <span data-ttu-id="eb4ce-109">Значение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4ce-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eb4ce-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb4ce-111">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-111">1.1</span></span>|
|[<span data-ttu-id="eb4ce-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eb4ce-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb4ce-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="eb4ce-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="eb4ce-114">Properties</span></span>

| <span data-ttu-id="eb4ce-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="eb4ce-115">Property</span></span> | <span data-ttu-id="eb4ce-116">Способов</span><span class="sxs-lookup"><span data-stu-id="eb4ce-116">Modes</span></span> | <span data-ttu-id="eb4ce-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="eb4ce-117">Return type</span></span> | <span data-ttu-id="eb4ce-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="eb4ce-118">Minimum</span></span><br><span data-ttu-id="eb4ce-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="eb4ce-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="eb4ce-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="eb4ce-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="eb4ce-121">Создание</span><span class="sxs-lookup"><span data-stu-id="eb4ce-121">Compose</span></span><br><span data-ttu-id="eb4ce-122">Read</span><span class="sxs-lookup"><span data-stu-id="eb4ce-122">Read</span></span> | <span data-ttu-id="eb4ce-123">String</span><span class="sxs-lookup"><span data-stu-id="eb4ce-123">String</span></span> | [<span data-ttu-id="eb4ce-124">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb4ce-125">диагностики</span><span class="sxs-lookup"><span data-stu-id="eb4ce-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="eb4ce-126">Создание</span><span class="sxs-lookup"><span data-stu-id="eb4ce-126">Compose</span></span><br><span data-ttu-id="eb4ce-127">Read</span><span class="sxs-lookup"><span data-stu-id="eb4ce-127">Read</span></span> | [<span data-ttu-id="eb4ce-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="eb4ce-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4) | [<span data-ttu-id="eb4ce-129">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb4ce-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="eb4ce-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="eb4ce-131">Создание</span><span class="sxs-lookup"><span data-stu-id="eb4ce-131">Compose</span></span><br><span data-ttu-id="eb4ce-132">Read</span><span class="sxs-lookup"><span data-stu-id="eb4ce-132">Read</span></span> | <span data-ttu-id="eb4ce-133">String</span><span class="sxs-lookup"><span data-stu-id="eb4ce-133">String</span></span> | [<span data-ttu-id="eb4ce-134">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb4ce-135">размещать</span><span class="sxs-lookup"><span data-stu-id="eb4ce-135">host</span></span>](#host-hosttype) | <span data-ttu-id="eb4ce-136">Создание</span><span class="sxs-lookup"><span data-stu-id="eb4ce-136">Compose</span></span><br><span data-ttu-id="eb4ce-137">Read</span><span class="sxs-lookup"><span data-stu-id="eb4ce-137">Read</span></span> | [<span data-ttu-id="eb4ce-138">HostType</span><span class="sxs-lookup"><span data-stu-id="eb4ce-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.4) | [<span data-ttu-id="eb4ce-139">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb4ce-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="eb4ce-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="eb4ce-141">Создание</span><span class="sxs-lookup"><span data-stu-id="eb4ce-141">Compose</span></span><br><span data-ttu-id="eb4ce-142">Read</span><span class="sxs-lookup"><span data-stu-id="eb4ce-142">Read</span></span> | [<span data-ttu-id="eb4ce-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="eb4ce-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4) | [<span data-ttu-id="eb4ce-144">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb4ce-145">управляем</span><span class="sxs-lookup"><span data-stu-id="eb4ce-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="eb4ce-146">Создание</span><span class="sxs-lookup"><span data-stu-id="eb4ce-146">Compose</span></span><br><span data-ttu-id="eb4ce-147">Read</span><span class="sxs-lookup"><span data-stu-id="eb4ce-147">Read</span></span> | [<span data-ttu-id="eb4ce-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="eb4ce-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.4) | [<span data-ttu-id="eb4ce-149">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb4ce-150">потребность</span><span class="sxs-lookup"><span data-stu-id="eb4ce-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="eb4ce-151">Создание</span><span class="sxs-lookup"><span data-stu-id="eb4ce-151">Compose</span></span><br><span data-ttu-id="eb4ce-152">Read</span><span class="sxs-lookup"><span data-stu-id="eb4ce-152">Read</span></span> | [<span data-ttu-id="eb4ce-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="eb4ce-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4) | [<span data-ttu-id="eb4ce-154">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb4ce-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="eb4ce-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="eb4ce-156">Создание</span><span class="sxs-lookup"><span data-stu-id="eb4ce-156">Compose</span></span><br><span data-ttu-id="eb4ce-157">Read</span><span class="sxs-lookup"><span data-stu-id="eb4ce-157">Read</span></span> | [<span data-ttu-id="eb4ce-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="eb4ce-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4) | [<span data-ttu-id="eb4ce-159">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="eb4ce-160">ui</span><span class="sxs-lookup"><span data-stu-id="eb4ce-160">ui</span></span>](#ui-ui) | <span data-ttu-id="eb4ce-161">Создание</span><span class="sxs-lookup"><span data-stu-id="eb4ce-161">Compose</span></span><br><span data-ttu-id="eb4ce-162">Read</span><span class="sxs-lookup"><span data-stu-id="eb4ce-162">Read</span></span> | [<span data-ttu-id="eb4ce-163">UI</span><span class="sxs-lookup"><span data-stu-id="eb4ce-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4) | [<span data-ttu-id="eb4ce-164">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="eb4ce-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="eb4ce-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="eb4ce-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="eb4ce-166">contentLanguage: String</span></span>

<span data-ttu-id="eb4ce-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="eb4ce-168">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4ce-169">Тип</span><span class="sxs-lookup"><span data-stu-id="eb4ce-169">Type</span></span>

*   <span data-ttu-id="eb4ce-170">String</span><span class="sxs-lookup"><span data-stu-id="eb4ce-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eb4ce-171">Требования</span><span class="sxs-lookup"><span data-stu-id="eb4ce-171">Requirements</span></span>

|<span data-ttu-id="eb4ce-172">Требование</span><span class="sxs-lookup"><span data-stu-id="eb4ce-172">Requirement</span></span>| <span data-ttu-id="eb4ce-173">Значение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4ce-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eb4ce-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb4ce-175">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-175">1.1</span></span>|
|[<span data-ttu-id="eb4ce-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eb4ce-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb4ce-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb4ce-178">Пример</span><span class="sxs-lookup"><span data-stu-id="eb4ce-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="eb4ce-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="eb4ce-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="eb4ce-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4ce-181">Тип</span><span class="sxs-lookup"><span data-stu-id="eb4ce-181">Type</span></span>

*   [<span data-ttu-id="eb4ce-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="eb4ce-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="eb4ce-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="eb4ce-183">Requirements</span></span>

|<span data-ttu-id="eb4ce-184">Требование</span><span class="sxs-lookup"><span data-stu-id="eb4ce-184">Requirement</span></span>| <span data-ttu-id="eb4ce-185">Значение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4ce-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eb4ce-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb4ce-187">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-187">1.1</span></span>|
|[<span data-ttu-id="eb4ce-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eb4ce-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb4ce-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb4ce-190">Пример</span><span class="sxs-lookup"><span data-stu-id="eb4ce-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="eb4ce-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="eb4ce-191">displayLanguage: String</span></span>

<span data-ttu-id="eb4ce-192">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="eb4ce-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4ce-194">Тип</span><span class="sxs-lookup"><span data-stu-id="eb4ce-194">Type</span></span>

*   <span data-ttu-id="eb4ce-195">String</span><span class="sxs-lookup"><span data-stu-id="eb4ce-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="eb4ce-196">Требования</span><span class="sxs-lookup"><span data-stu-id="eb4ce-196">Requirements</span></span>

|<span data-ttu-id="eb4ce-197">Требование</span><span class="sxs-lookup"><span data-stu-id="eb4ce-197">Requirement</span></span>| <span data-ttu-id="eb4ce-198">Значение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4ce-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eb4ce-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb4ce-200">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-200">1.1</span></span>|
|[<span data-ttu-id="eb4ce-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eb4ce-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb4ce-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb4ce-203">Пример</span><span class="sxs-lookup"><span data-stu-id="eb4ce-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="eb4ce-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="eb4ce-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="eb4ce-205">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4ce-206">Тип</span><span class="sxs-lookup"><span data-stu-id="eb4ce-206">Type</span></span>

*   [<span data-ttu-id="eb4ce-207">HostType</span><span class="sxs-lookup"><span data-stu-id="eb4ce-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="eb4ce-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="eb4ce-208">Requirements</span></span>

|<span data-ttu-id="eb4ce-209">Требование</span><span class="sxs-lookup"><span data-stu-id="eb4ce-209">Requirement</span></span>| <span data-ttu-id="eb4ce-210">Значение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4ce-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eb4ce-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb4ce-212">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-212">1.1</span></span>|
|[<span data-ttu-id="eb4ce-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eb4ce-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb4ce-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb4ce-215">Пример</span><span class="sxs-lookup"><span data-stu-id="eb4ce-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="eb4ce-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="eb4ce-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="eb4ce-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4ce-218">Тип</span><span class="sxs-lookup"><span data-stu-id="eb4ce-218">Type</span></span>

*   [<span data-ttu-id="eb4ce-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="eb4ce-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="eb4ce-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="eb4ce-220">Requirements</span></span>

|<span data-ttu-id="eb4ce-221">Требование</span><span class="sxs-lookup"><span data-stu-id="eb4ce-221">Requirement</span></span>| <span data-ttu-id="eb4ce-222">Значение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4ce-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eb4ce-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb4ce-224">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-224">1.1</span></span>|
|[<span data-ttu-id="eb4ce-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eb4ce-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb4ce-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb4ce-227">Пример</span><span class="sxs-lookup"><span data-stu-id="eb4ce-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="eb4ce-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="eb4ce-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="eb4ce-229">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4ce-230">Тип</span><span class="sxs-lookup"><span data-stu-id="eb4ce-230">Type</span></span>

*   [<span data-ttu-id="eb4ce-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="eb4ce-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="eb4ce-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="eb4ce-232">Requirements</span></span>

|<span data-ttu-id="eb4ce-233">Требование</span><span class="sxs-lookup"><span data-stu-id="eb4ce-233">Requirement</span></span>| <span data-ttu-id="eb4ce-234">Значение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4ce-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eb4ce-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb4ce-236">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-236">1.1</span></span>|
|[<span data-ttu-id="eb4ce-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eb4ce-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb4ce-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="eb4ce-239">Пример</span><span class="sxs-lookup"><span data-stu-id="eb4ce-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="eb4ce-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="eb4ce-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="eb4ce-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="eb4ce-242">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4ce-243">Тип</span><span class="sxs-lookup"><span data-stu-id="eb4ce-243">Type</span></span>

*   [<span data-ttu-id="eb4ce-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="eb4ce-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="eb4ce-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="eb4ce-245">Requirements</span></span>

|<span data-ttu-id="eb4ce-246">Требование</span><span class="sxs-lookup"><span data-stu-id="eb4ce-246">Requirement</span></span>| <span data-ttu-id="eb4ce-247">Значение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4ce-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eb4ce-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb4ce-249">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-249">1.1</span></span>|
|[<span data-ttu-id="eb4ce-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="eb4ce-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="eb4ce-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="eb4ce-251">Restricted</span></span>|
|[<span data-ttu-id="eb4ce-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eb4ce-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb4ce-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="eb4ce-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="eb4ce-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="eb4ce-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="eb4ce-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="eb4ce-256">Тип</span><span class="sxs-lookup"><span data-stu-id="eb4ce-256">Type</span></span>

*   [<span data-ttu-id="eb4ce-257">UI</span><span class="sxs-lookup"><span data-stu-id="eb4ce-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="eb4ce-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="eb4ce-258">Requirements</span></span>

|<span data-ttu-id="eb4ce-259">Требование</span><span class="sxs-lookup"><span data-stu-id="eb4ce-259">Requirement</span></span>| <span data-ttu-id="eb4ce-260">Значение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="eb4ce-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="eb4ce-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="eb4ce-262">1.1</span><span class="sxs-lookup"><span data-stu-id="eb4ce-262">1.1</span></span>|
|[<span data-ttu-id="eb4ce-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="eb4ce-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="eb4ce-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="eb4ce-264">Compose or Read</span></span>|
