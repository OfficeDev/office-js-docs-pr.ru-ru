---
title: Office. Context — набор обязательных элементов 1,3
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,3.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: d8edf2568fcb8f9cecf075781bf9b74b3480f8ea
ms.sourcegitcommit: be23b68eb661015508797333915b44381dd29bdb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/08/2020
ms.locfileid: "44612012"
---
# <a name="context-mailbox-requirement-set-13"></a><span data-ttu-id="5e515-103">контекст (набор требований для почтового ящика 1,3)</span><span class="sxs-lookup"><span data-stu-id="5e515-103">context (Mailbox requirement set 1.3)</span></span>

### <a name="officecontext"></a><span data-ttu-id="5e515-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="5e515-104">[Office](office.md).context</span></span>

<span data-ttu-id="5e515-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="5e515-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="5e515-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.3).</span><span class="sxs-lookup"><span data-stu-id="5e515-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3).</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e515-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e515-107">Requirements</span></span>

|<span data-ttu-id="5e515-108">Требование</span><span class="sxs-lookup"><span data-stu-id="5e515-108">Requirement</span></span>| <span data-ttu-id="5e515-109">Значение</span><span class="sxs-lookup"><span data-stu-id="5e515-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e515-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e515-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5e515-111">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-111">1.1</span></span>|
|[<span data-ttu-id="5e515-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e515-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5e515-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e515-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="5e515-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="5e515-114">Properties</span></span>

| <span data-ttu-id="5e515-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="5e515-115">Property</span></span> | <span data-ttu-id="5e515-116">Способов</span><span class="sxs-lookup"><span data-stu-id="5e515-116">Modes</span></span> | <span data-ttu-id="5e515-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="5e515-117">Return type</span></span> | <span data-ttu-id="5e515-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="5e515-118">Minimum</span></span><br><span data-ttu-id="5e515-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="5e515-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="5e515-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="5e515-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="5e515-121">Создание</span><span class="sxs-lookup"><span data-stu-id="5e515-121">Compose</span></span><br><span data-ttu-id="5e515-122">Read</span><span class="sxs-lookup"><span data-stu-id="5e515-122">Read</span></span> | <span data-ttu-id="5e515-123">String</span><span class="sxs-lookup"><span data-stu-id="5e515-123">String</span></span> | [<span data-ttu-id="5e515-124">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5e515-125">диагностики</span><span class="sxs-lookup"><span data-stu-id="5e515-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="5e515-126">Создание</span><span class="sxs-lookup"><span data-stu-id="5e515-126">Compose</span></span><br><span data-ttu-id="5e515-127">Read</span><span class="sxs-lookup"><span data-stu-id="5e515-127">Read</span></span> | [<span data-ttu-id="5e515-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="5e515-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3) | [<span data-ttu-id="5e515-129">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5e515-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="5e515-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="5e515-131">Создание</span><span class="sxs-lookup"><span data-stu-id="5e515-131">Compose</span></span><br><span data-ttu-id="5e515-132">Read</span><span class="sxs-lookup"><span data-stu-id="5e515-132">Read</span></span> | <span data-ttu-id="5e515-133">String</span><span class="sxs-lookup"><span data-stu-id="5e515-133">String</span></span> | [<span data-ttu-id="5e515-134">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5e515-135">размещать</span><span class="sxs-lookup"><span data-stu-id="5e515-135">host</span></span>](#host-hosttype) | <span data-ttu-id="5e515-136">Создание</span><span class="sxs-lookup"><span data-stu-id="5e515-136">Compose</span></span><br><span data-ttu-id="5e515-137">Read</span><span class="sxs-lookup"><span data-stu-id="5e515-137">Read</span></span> | [<span data-ttu-id="5e515-138">HostType</span><span class="sxs-lookup"><span data-stu-id="5e515-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.3) | [<span data-ttu-id="5e515-139">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5e515-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="5e515-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="5e515-141">Создание</span><span class="sxs-lookup"><span data-stu-id="5e515-141">Compose</span></span><br><span data-ttu-id="5e515-142">Read</span><span class="sxs-lookup"><span data-stu-id="5e515-142">Read</span></span> | [<span data-ttu-id="5e515-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="5e515-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3) | [<span data-ttu-id="5e515-144">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5e515-145">управляем</span><span class="sxs-lookup"><span data-stu-id="5e515-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="5e515-146">Создание</span><span class="sxs-lookup"><span data-stu-id="5e515-146">Compose</span></span><br><span data-ttu-id="5e515-147">Read</span><span class="sxs-lookup"><span data-stu-id="5e515-147">Read</span></span> | [<span data-ttu-id="5e515-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="5e515-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.3) | [<span data-ttu-id="5e515-149">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5e515-150">потребность</span><span class="sxs-lookup"><span data-stu-id="5e515-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="5e515-151">Создание</span><span class="sxs-lookup"><span data-stu-id="5e515-151">Compose</span></span><br><span data-ttu-id="5e515-152">Read</span><span class="sxs-lookup"><span data-stu-id="5e515-152">Read</span></span> | [<span data-ttu-id="5e515-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="5e515-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3) | [<span data-ttu-id="5e515-154">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5e515-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="5e515-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="5e515-156">Создание</span><span class="sxs-lookup"><span data-stu-id="5e515-156">Compose</span></span><br><span data-ttu-id="5e515-157">Read</span><span class="sxs-lookup"><span data-stu-id="5e515-157">Read</span></span> | [<span data-ttu-id="5e515-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="5e515-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3) | [<span data-ttu-id="5e515-159">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="5e515-160">ui</span><span class="sxs-lookup"><span data-stu-id="5e515-160">ui</span></span>](#ui-ui) | <span data-ttu-id="5e515-161">Создание</span><span class="sxs-lookup"><span data-stu-id="5e515-161">Compose</span></span><br><span data-ttu-id="5e515-162">Read</span><span class="sxs-lookup"><span data-stu-id="5e515-162">Read</span></span> | [<span data-ttu-id="5e515-163">UI</span><span class="sxs-lookup"><span data-stu-id="5e515-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3) | [<span data-ttu-id="5e515-164">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="5e515-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="5e515-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="5e515-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="5e515-166">contentLanguage: String</span></span>

<span data-ttu-id="5e515-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="5e515-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="5e515-168">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="5e515-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="5e515-169">Тип</span><span class="sxs-lookup"><span data-stu-id="5e515-169">Type</span></span>

*   <span data-ttu-id="5e515-170">String</span><span class="sxs-lookup"><span data-stu-id="5e515-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e515-171">Требования</span><span class="sxs-lookup"><span data-stu-id="5e515-171">Requirements</span></span>

|<span data-ttu-id="5e515-172">Требование</span><span class="sxs-lookup"><span data-stu-id="5e515-172">Requirement</span></span>| <span data-ttu-id="5e515-173">Значение</span><span class="sxs-lookup"><span data-stu-id="5e515-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e515-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e515-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5e515-175">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-175">1.1</span></span>|
|[<span data-ttu-id="5e515-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e515-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5e515-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e515-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e515-178">Пример</span><span class="sxs-lookup"><span data-stu-id="5e515-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="5e515-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="5e515-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="5e515-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="5e515-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="5e515-181">Тип</span><span class="sxs-lookup"><span data-stu-id="5e515-181">Type</span></span>

*   [<span data-ttu-id="5e515-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="5e515-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="5e515-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e515-183">Requirements</span></span>

|<span data-ttu-id="5e515-184">Требование</span><span class="sxs-lookup"><span data-stu-id="5e515-184">Requirement</span></span>| <span data-ttu-id="5e515-185">Значение</span><span class="sxs-lookup"><span data-stu-id="5e515-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e515-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e515-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5e515-187">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-187">1.1</span></span>|
|[<span data-ttu-id="5e515-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e515-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5e515-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e515-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e515-190">Пример</span><span class="sxs-lookup"><span data-stu-id="5e515-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="5e515-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="5e515-191">displayLanguage: String</span></span>

<span data-ttu-id="5e515-192">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="5e515-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="5e515-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="5e515-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="5e515-194">Тип</span><span class="sxs-lookup"><span data-stu-id="5e515-194">Type</span></span>

*   <span data-ttu-id="5e515-195">String</span><span class="sxs-lookup"><span data-stu-id="5e515-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="5e515-196">Требования</span><span class="sxs-lookup"><span data-stu-id="5e515-196">Requirements</span></span>

|<span data-ttu-id="5e515-197">Требование</span><span class="sxs-lookup"><span data-stu-id="5e515-197">Requirement</span></span>| <span data-ttu-id="5e515-198">Значение</span><span class="sxs-lookup"><span data-stu-id="5e515-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e515-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e515-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5e515-200">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-200">1.1</span></span>|
|[<span data-ttu-id="5e515-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e515-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5e515-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e515-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e515-203">Пример</span><span class="sxs-lookup"><span data-stu-id="5e515-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="5e515-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="5e515-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="5e515-205">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="5e515-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="5e515-206">Тип</span><span class="sxs-lookup"><span data-stu-id="5e515-206">Type</span></span>

*   [<span data-ttu-id="5e515-207">HostType</span><span class="sxs-lookup"><span data-stu-id="5e515-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="5e515-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e515-208">Requirements</span></span>

|<span data-ttu-id="5e515-209">Требование</span><span class="sxs-lookup"><span data-stu-id="5e515-209">Requirement</span></span>| <span data-ttu-id="5e515-210">Значение</span><span class="sxs-lookup"><span data-stu-id="5e515-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e515-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e515-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5e515-212">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-212">1.1</span></span>|
|[<span data-ttu-id="5e515-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e515-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5e515-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e515-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e515-215">Пример</span><span class="sxs-lookup"><span data-stu-id="5e515-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="5e515-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="5e515-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="5e515-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="5e515-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="5e515-218">Тип</span><span class="sxs-lookup"><span data-stu-id="5e515-218">Type</span></span>

*   [<span data-ttu-id="5e515-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="5e515-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="5e515-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e515-220">Requirements</span></span>

|<span data-ttu-id="5e515-221">Требование</span><span class="sxs-lookup"><span data-stu-id="5e515-221">Requirement</span></span>| <span data-ttu-id="5e515-222">Значение</span><span class="sxs-lookup"><span data-stu-id="5e515-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e515-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e515-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5e515-224">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-224">1.1</span></span>|
|[<span data-ttu-id="5e515-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e515-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5e515-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e515-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e515-227">Пример</span><span class="sxs-lookup"><span data-stu-id="5e515-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="5e515-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="5e515-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="5e515-229">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="5e515-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="5e515-230">Тип</span><span class="sxs-lookup"><span data-stu-id="5e515-230">Type</span></span>

*   [<span data-ttu-id="5e515-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="5e515-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="5e515-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e515-232">Requirements</span></span>

|<span data-ttu-id="5e515-233">Требование</span><span class="sxs-lookup"><span data-stu-id="5e515-233">Requirement</span></span>| <span data-ttu-id="5e515-234">Значение</span><span class="sxs-lookup"><span data-stu-id="5e515-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e515-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e515-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5e515-236">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-236">1.1</span></span>|
|[<span data-ttu-id="5e515-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e515-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5e515-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e515-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="5e515-239">Пример</span><span class="sxs-lookup"><span data-stu-id="5e515-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="5e515-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="5e515-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="5e515-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="5e515-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="5e515-242">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="5e515-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="5e515-243">Тип</span><span class="sxs-lookup"><span data-stu-id="5e515-243">Type</span></span>

*   [<span data-ttu-id="5e515-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="5e515-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="5e515-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e515-245">Requirements</span></span>

|<span data-ttu-id="5e515-246">Требование</span><span class="sxs-lookup"><span data-stu-id="5e515-246">Requirement</span></span>| <span data-ttu-id="5e515-247">Значение</span><span class="sxs-lookup"><span data-stu-id="5e515-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e515-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e515-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5e515-249">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-249">1.1</span></span>|
|[<span data-ttu-id="5e515-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="5e515-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="5e515-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="5e515-251">Restricted</span></span>|
|[<span data-ttu-id="5e515-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e515-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5e515-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e515-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="5e515-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="5e515-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="5e515-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="5e515-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="5e515-256">Тип</span><span class="sxs-lookup"><span data-stu-id="5e515-256">Type</span></span>

*   [<span data-ttu-id="5e515-257">UI</span><span class="sxs-lookup"><span data-stu-id="5e515-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="5e515-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="5e515-258">Requirements</span></span>

|<span data-ttu-id="5e515-259">Требование</span><span class="sxs-lookup"><span data-stu-id="5e515-259">Requirement</span></span>| <span data-ttu-id="5e515-260">Значение</span><span class="sxs-lookup"><span data-stu-id="5e515-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="5e515-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="5e515-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="5e515-262">1.1</span><span class="sxs-lookup"><span data-stu-id="5e515-262">1.1</span></span>|
|[<span data-ttu-id="5e515-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="5e515-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="5e515-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="5e515-264">Compose or Read</span></span>|
