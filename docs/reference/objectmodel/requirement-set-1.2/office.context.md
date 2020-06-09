---
title: Office. Context — набор обязательных элементов 1,2
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,2.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 4bff6e70b143480b0d4b98925f504870d8a2bbb2
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44610528"
---
# <a name="context-mailbox-requirement-set-12"></a><span data-ttu-id="111b2-103">контекст (набор требований для почтового ящика 1,2)</span><span class="sxs-lookup"><span data-stu-id="111b2-103">context (Mailbox requirement set 1.2)</span></span>

### <a name="officecontext"></a><span data-ttu-id="111b2-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="111b2-104">[Office](office.md).context</span></span>

<span data-ttu-id="111b2-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="111b2-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="111b2-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.2).</span><span class="sxs-lookup"><span data-stu-id="111b2-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.2).</span></span>

##### <a name="requirements"></a><span data-ttu-id="111b2-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="111b2-107">Requirements</span></span>

|<span data-ttu-id="111b2-108">Требование</span><span class="sxs-lookup"><span data-stu-id="111b2-108">Requirement</span></span>| <span data-ttu-id="111b2-109">Значение</span><span class="sxs-lookup"><span data-stu-id="111b2-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="111b2-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="111b2-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="111b2-111">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-111">1.1</span></span>|
|[<span data-ttu-id="111b2-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="111b2-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="111b2-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="111b2-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="111b2-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="111b2-114">Properties</span></span>

| <span data-ttu-id="111b2-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="111b2-115">Property</span></span> | <span data-ttu-id="111b2-116">Способов</span><span class="sxs-lookup"><span data-stu-id="111b2-116">Modes</span></span> | <span data-ttu-id="111b2-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="111b2-117">Return type</span></span> | <span data-ttu-id="111b2-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="111b2-118">Minimum</span></span><br><span data-ttu-id="111b2-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="111b2-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="111b2-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="111b2-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="111b2-121">Создание</span><span class="sxs-lookup"><span data-stu-id="111b2-121">Compose</span></span><br><span data-ttu-id="111b2-122">Read</span><span class="sxs-lookup"><span data-stu-id="111b2-122">Read</span></span> | <span data-ttu-id="111b2-123">String</span><span class="sxs-lookup"><span data-stu-id="111b2-123">String</span></span> | [<span data-ttu-id="111b2-124">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="111b2-125">диагностики</span><span class="sxs-lookup"><span data-stu-id="111b2-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="111b2-126">Создание</span><span class="sxs-lookup"><span data-stu-id="111b2-126">Compose</span></span><br><span data-ttu-id="111b2-127">Read</span><span class="sxs-lookup"><span data-stu-id="111b2-127">Read</span></span> | [<span data-ttu-id="111b2-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="111b2-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.2) | [<span data-ttu-id="111b2-129">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="111b2-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="111b2-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="111b2-131">Создание</span><span class="sxs-lookup"><span data-stu-id="111b2-131">Compose</span></span><br><span data-ttu-id="111b2-132">Read</span><span class="sxs-lookup"><span data-stu-id="111b2-132">Read</span></span> | <span data-ttu-id="111b2-133">String</span><span class="sxs-lookup"><span data-stu-id="111b2-133">String</span></span> | [<span data-ttu-id="111b2-134">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="111b2-135">размещать</span><span class="sxs-lookup"><span data-stu-id="111b2-135">host</span></span>](#host-hosttype) | <span data-ttu-id="111b2-136">Создание</span><span class="sxs-lookup"><span data-stu-id="111b2-136">Compose</span></span><br><span data-ttu-id="111b2-137">Read</span><span class="sxs-lookup"><span data-stu-id="111b2-137">Read</span></span> | [<span data-ttu-id="111b2-138">HostType</span><span class="sxs-lookup"><span data-stu-id="111b2-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.2) | [<span data-ttu-id="111b2-139">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="111b2-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="111b2-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="111b2-141">Создание</span><span class="sxs-lookup"><span data-stu-id="111b2-141">Compose</span></span><br><span data-ttu-id="111b2-142">Read</span><span class="sxs-lookup"><span data-stu-id="111b2-142">Read</span></span> | [<span data-ttu-id="111b2-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="111b2-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.2) | [<span data-ttu-id="111b2-144">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="111b2-145">управляем</span><span class="sxs-lookup"><span data-stu-id="111b2-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="111b2-146">Создание</span><span class="sxs-lookup"><span data-stu-id="111b2-146">Compose</span></span><br><span data-ttu-id="111b2-147">Read</span><span class="sxs-lookup"><span data-stu-id="111b2-147">Read</span></span> | [<span data-ttu-id="111b2-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="111b2-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.2) | [<span data-ttu-id="111b2-149">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="111b2-150">потребность</span><span class="sxs-lookup"><span data-stu-id="111b2-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="111b2-151">Создание</span><span class="sxs-lookup"><span data-stu-id="111b2-151">Compose</span></span><br><span data-ttu-id="111b2-152">Read</span><span class="sxs-lookup"><span data-stu-id="111b2-152">Read</span></span> | [<span data-ttu-id="111b2-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="111b2-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.2) | [<span data-ttu-id="111b2-154">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="111b2-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="111b2-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="111b2-156">Создание</span><span class="sxs-lookup"><span data-stu-id="111b2-156">Compose</span></span><br><span data-ttu-id="111b2-157">Read</span><span class="sxs-lookup"><span data-stu-id="111b2-157">Read</span></span> | [<span data-ttu-id="111b2-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="111b2-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.2) | [<span data-ttu-id="111b2-159">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="111b2-160">ui</span><span class="sxs-lookup"><span data-stu-id="111b2-160">ui</span></span>](#ui-ui) | <span data-ttu-id="111b2-161">Создание</span><span class="sxs-lookup"><span data-stu-id="111b2-161">Compose</span></span><br><span data-ttu-id="111b2-162">Read</span><span class="sxs-lookup"><span data-stu-id="111b2-162">Read</span></span> | [<span data-ttu-id="111b2-163">UI</span><span class="sxs-lookup"><span data-stu-id="111b2-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.2) | [<span data-ttu-id="111b2-164">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="111b2-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="111b2-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="111b2-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="111b2-166">contentLanguage: String</span></span>

<span data-ttu-id="111b2-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="111b2-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="111b2-168">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="111b2-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="111b2-169">Тип</span><span class="sxs-lookup"><span data-stu-id="111b2-169">Type</span></span>

*   <span data-ttu-id="111b2-170">String</span><span class="sxs-lookup"><span data-stu-id="111b2-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="111b2-171">Требования</span><span class="sxs-lookup"><span data-stu-id="111b2-171">Requirements</span></span>

|<span data-ttu-id="111b2-172">Требование</span><span class="sxs-lookup"><span data-stu-id="111b2-172">Requirement</span></span>| <span data-ttu-id="111b2-173">Значение</span><span class="sxs-lookup"><span data-stu-id="111b2-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="111b2-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="111b2-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="111b2-175">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-175">1.1</span></span>|
|[<span data-ttu-id="111b2-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="111b2-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="111b2-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="111b2-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="111b2-178">Пример</span><span class="sxs-lookup"><span data-stu-id="111b2-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="111b2-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="111b2-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="111b2-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="111b2-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="111b2-181">Тип</span><span class="sxs-lookup"><span data-stu-id="111b2-181">Type</span></span>

*   [<span data-ttu-id="111b2-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="111b2-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="111b2-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="111b2-183">Requirements</span></span>

|<span data-ttu-id="111b2-184">Требование</span><span class="sxs-lookup"><span data-stu-id="111b2-184">Requirement</span></span>| <span data-ttu-id="111b2-185">Значение</span><span class="sxs-lookup"><span data-stu-id="111b2-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="111b2-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="111b2-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="111b2-187">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-187">1.1</span></span>|
|[<span data-ttu-id="111b2-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="111b2-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="111b2-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="111b2-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="111b2-190">Пример</span><span class="sxs-lookup"><span data-stu-id="111b2-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="111b2-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="111b2-191">displayLanguage: String</span></span>

<span data-ttu-id="111b2-192">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="111b2-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="111b2-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="111b2-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="111b2-194">Тип</span><span class="sxs-lookup"><span data-stu-id="111b2-194">Type</span></span>

*   <span data-ttu-id="111b2-195">String</span><span class="sxs-lookup"><span data-stu-id="111b2-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="111b2-196">Требования</span><span class="sxs-lookup"><span data-stu-id="111b2-196">Requirements</span></span>

|<span data-ttu-id="111b2-197">Требование</span><span class="sxs-lookup"><span data-stu-id="111b2-197">Requirement</span></span>| <span data-ttu-id="111b2-198">Значение</span><span class="sxs-lookup"><span data-stu-id="111b2-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="111b2-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="111b2-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="111b2-200">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-200">1.1</span></span>|
|[<span data-ttu-id="111b2-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="111b2-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="111b2-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="111b2-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="111b2-203">Пример</span><span class="sxs-lookup"><span data-stu-id="111b2-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="111b2-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="111b2-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="111b2-205">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="111b2-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="111b2-206">Тип</span><span class="sxs-lookup"><span data-stu-id="111b2-206">Type</span></span>

*   [<span data-ttu-id="111b2-207">HostType</span><span class="sxs-lookup"><span data-stu-id="111b2-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="111b2-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="111b2-208">Requirements</span></span>

|<span data-ttu-id="111b2-209">Требование</span><span class="sxs-lookup"><span data-stu-id="111b2-209">Requirement</span></span>| <span data-ttu-id="111b2-210">Значение</span><span class="sxs-lookup"><span data-stu-id="111b2-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="111b2-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="111b2-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="111b2-212">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-212">1.1</span></span>|
|[<span data-ttu-id="111b2-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="111b2-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="111b2-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="111b2-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="111b2-215">Пример</span><span class="sxs-lookup"><span data-stu-id="111b2-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="111b2-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="111b2-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="111b2-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="111b2-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="111b2-218">Тип</span><span class="sxs-lookup"><span data-stu-id="111b2-218">Type</span></span>

*   [<span data-ttu-id="111b2-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="111b2-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="111b2-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="111b2-220">Requirements</span></span>

|<span data-ttu-id="111b2-221">Требование</span><span class="sxs-lookup"><span data-stu-id="111b2-221">Requirement</span></span>| <span data-ttu-id="111b2-222">Значение</span><span class="sxs-lookup"><span data-stu-id="111b2-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="111b2-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="111b2-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="111b2-224">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-224">1.1</span></span>|
|[<span data-ttu-id="111b2-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="111b2-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="111b2-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="111b2-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="111b2-227">Пример</span><span class="sxs-lookup"><span data-stu-id="111b2-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="111b2-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="111b2-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="111b2-229">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="111b2-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="111b2-230">Тип</span><span class="sxs-lookup"><span data-stu-id="111b2-230">Type</span></span>

*   [<span data-ttu-id="111b2-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="111b2-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="111b2-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="111b2-232">Requirements</span></span>

|<span data-ttu-id="111b2-233">Требование</span><span class="sxs-lookup"><span data-stu-id="111b2-233">Requirement</span></span>| <span data-ttu-id="111b2-234">Значение</span><span class="sxs-lookup"><span data-stu-id="111b2-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="111b2-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="111b2-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="111b2-236">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-236">1.1</span></span>|
|[<span data-ttu-id="111b2-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="111b2-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="111b2-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="111b2-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="111b2-239">Пример</span><span class="sxs-lookup"><span data-stu-id="111b2-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="111b2-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="111b2-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="111b2-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="111b2-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="111b2-242">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="111b2-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="111b2-243">Тип</span><span class="sxs-lookup"><span data-stu-id="111b2-243">Type</span></span>

*   [<span data-ttu-id="111b2-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="111b2-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="111b2-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="111b2-245">Requirements</span></span>

|<span data-ttu-id="111b2-246">Требование</span><span class="sxs-lookup"><span data-stu-id="111b2-246">Requirement</span></span>| <span data-ttu-id="111b2-247">Значение</span><span class="sxs-lookup"><span data-stu-id="111b2-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="111b2-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="111b2-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="111b2-249">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-249">1.1</span></span>|
|[<span data-ttu-id="111b2-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="111b2-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="111b2-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="111b2-251">Restricted</span></span>|
|[<span data-ttu-id="111b2-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="111b2-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="111b2-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="111b2-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="111b2-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="111b2-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="111b2-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="111b2-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="111b2-256">Тип</span><span class="sxs-lookup"><span data-stu-id="111b2-256">Type</span></span>

*   [<span data-ttu-id="111b2-257">UI</span><span class="sxs-lookup"><span data-stu-id="111b2-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="111b2-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="111b2-258">Requirements</span></span>

|<span data-ttu-id="111b2-259">Требование</span><span class="sxs-lookup"><span data-stu-id="111b2-259">Requirement</span></span>| <span data-ttu-id="111b2-260">Значение</span><span class="sxs-lookup"><span data-stu-id="111b2-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="111b2-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="111b2-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="111b2-262">1.1</span><span class="sxs-lookup"><span data-stu-id="111b2-262">1.1</span></span>|
|[<span data-ttu-id="111b2-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="111b2-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="111b2-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="111b2-264">Compose or Read</span></span>|
