---
title: Office. Context — набор обязательных элементов 1,3
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: adea7414da7242ba3d2f9d57210c934d61e146df
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42163768"
---
# <a name="context"></a><span data-ttu-id="858d0-102">context</span><span class="sxs-lookup"><span data-stu-id="858d0-102">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="858d0-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="858d0-103">[Office](office.md).context</span></span>

<span data-ttu-id="858d0-104">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="858d0-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="858d0-105">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.3).</span><span class="sxs-lookup"><span data-stu-id="858d0-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3).</span></span>

##### <a name="requirements"></a><span data-ttu-id="858d0-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="858d0-106">Requirements</span></span>

|<span data-ttu-id="858d0-107">Требование</span><span class="sxs-lookup"><span data-stu-id="858d0-107">Requirement</span></span>| <span data-ttu-id="858d0-108">Значение</span><span class="sxs-lookup"><span data-stu-id="858d0-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="858d0-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="858d0-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="858d0-110">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-110">1.1</span></span>|
|[<span data-ttu-id="858d0-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="858d0-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="858d0-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="858d0-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="858d0-113">Properties</span></span>

| <span data-ttu-id="858d0-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="858d0-114">Property</span></span> | <span data-ttu-id="858d0-115">Способов</span><span class="sxs-lookup"><span data-stu-id="858d0-115">Modes</span></span> | <span data-ttu-id="858d0-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="858d0-116">Return type</span></span> | <span data-ttu-id="858d0-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="858d0-117">Minimum</span></span><br><span data-ttu-id="858d0-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="858d0-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="858d0-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="858d0-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="858d0-120">Создание</span><span class="sxs-lookup"><span data-stu-id="858d0-120">Compose</span></span><br><span data-ttu-id="858d0-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-121">Read</span></span> | <span data-ttu-id="858d0-122">String</span><span class="sxs-lookup"><span data-stu-id="858d0-122">String</span></span> | [<span data-ttu-id="858d0-123">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-123">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="858d0-124">diagnostics</span><span class="sxs-lookup"><span data-stu-id="858d0-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="858d0-125">Создание</span><span class="sxs-lookup"><span data-stu-id="858d0-125">Compose</span></span><br><span data-ttu-id="858d0-126">Чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-126">Read</span></span> | [<span data-ttu-id="858d0-127">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="858d0-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3) | [<span data-ttu-id="858d0-128">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="858d0-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="858d0-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="858d0-130">Создание</span><span class="sxs-lookup"><span data-stu-id="858d0-130">Compose</span></span><br><span data-ttu-id="858d0-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-131">Read</span></span> | <span data-ttu-id="858d0-132">String</span><span class="sxs-lookup"><span data-stu-id="858d0-132">String</span></span> | [<span data-ttu-id="858d0-133">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="858d0-134">размещать</span><span class="sxs-lookup"><span data-stu-id="858d0-134">host</span></span>](#host-hosttype) | <span data-ttu-id="858d0-135">Создание</span><span class="sxs-lookup"><span data-stu-id="858d0-135">Compose</span></span><br><span data-ttu-id="858d0-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-136">Read</span></span> | [<span data-ttu-id="858d0-137">HostType</span><span class="sxs-lookup"><span data-stu-id="858d0-137">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.3) | [<span data-ttu-id="858d0-138">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="858d0-139">mailbox</span><span class="sxs-lookup"><span data-stu-id="858d0-139">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="858d0-140">Создание</span><span class="sxs-lookup"><span data-stu-id="858d0-140">Compose</span></span><br><span data-ttu-id="858d0-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-141">Read</span></span> | [<span data-ttu-id="858d0-142">Mailbox</span><span class="sxs-lookup"><span data-stu-id="858d0-142">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3) | [<span data-ttu-id="858d0-143">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="858d0-144">управляем</span><span class="sxs-lookup"><span data-stu-id="858d0-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="858d0-145">Создание</span><span class="sxs-lookup"><span data-stu-id="858d0-145">Compose</span></span><br><span data-ttu-id="858d0-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-146">Read</span></span> | [<span data-ttu-id="858d0-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="858d0-147">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.3) | [<span data-ttu-id="858d0-148">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="858d0-149">потребность</span><span class="sxs-lookup"><span data-stu-id="858d0-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="858d0-150">Создание</span><span class="sxs-lookup"><span data-stu-id="858d0-150">Compose</span></span><br><span data-ttu-id="858d0-151">Чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-151">Read</span></span> | [<span data-ttu-id="858d0-152">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="858d0-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3) | [<span data-ttu-id="858d0-153">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-153">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="858d0-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="858d0-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="858d0-155">Создание</span><span class="sxs-lookup"><span data-stu-id="858d0-155">Compose</span></span><br><span data-ttu-id="858d0-156">Чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-156">Read</span></span> | [<span data-ttu-id="858d0-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="858d0-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3) | [<span data-ttu-id="858d0-158">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="858d0-159">ui</span><span class="sxs-lookup"><span data-stu-id="858d0-159">ui</span></span>](#ui-ui) | <span data-ttu-id="858d0-160">Создание</span><span class="sxs-lookup"><span data-stu-id="858d0-160">Compose</span></span><br><span data-ttu-id="858d0-161">Чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-161">Read</span></span> | [<span data-ttu-id="858d0-162">UI</span><span class="sxs-lookup"><span data-stu-id="858d0-162">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3) | [<span data-ttu-id="858d0-163">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="858d0-164">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="858d0-164">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="858d0-165">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="858d0-165">contentLanguage: String</span></span>

<span data-ttu-id="858d0-166">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="858d0-166">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="858d0-167">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="858d0-167">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="858d0-168">Тип</span><span class="sxs-lookup"><span data-stu-id="858d0-168">Type</span></span>

*   <span data-ttu-id="858d0-169">String</span><span class="sxs-lookup"><span data-stu-id="858d0-169">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="858d0-170">Требования</span><span class="sxs-lookup"><span data-stu-id="858d0-170">Requirements</span></span>

|<span data-ttu-id="858d0-171">Требование</span><span class="sxs-lookup"><span data-stu-id="858d0-171">Requirement</span></span>| <span data-ttu-id="858d0-172">Значение</span><span class="sxs-lookup"><span data-stu-id="858d0-172">Value</span></span>|
|---|---|
|[<span data-ttu-id="858d0-173">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="858d0-173">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="858d0-174">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-174">1.1</span></span>|
|[<span data-ttu-id="858d0-175">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="858d0-175">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="858d0-176">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-176">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="858d0-177">Пример</span><span class="sxs-lookup"><span data-stu-id="858d0-177">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="858d0-178">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="858d0-178">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="858d0-179">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="858d0-179">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="858d0-180">Тип</span><span class="sxs-lookup"><span data-stu-id="858d0-180">Type</span></span>

*   [<span data-ttu-id="858d0-181">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="858d0-181">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="858d0-182">Requirements</span><span class="sxs-lookup"><span data-stu-id="858d0-182">Requirements</span></span>

|<span data-ttu-id="858d0-183">Требование</span><span class="sxs-lookup"><span data-stu-id="858d0-183">Requirement</span></span>| <span data-ttu-id="858d0-184">Значение</span><span class="sxs-lookup"><span data-stu-id="858d0-184">Value</span></span>|
|---|---|
|[<span data-ttu-id="858d0-185">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="858d0-185">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="858d0-186">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-186">1.1</span></span>|
|[<span data-ttu-id="858d0-187">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="858d0-187">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="858d0-188">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-188">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="858d0-189">Пример</span><span class="sxs-lookup"><span data-stu-id="858d0-189">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="858d0-190">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="858d0-190">displayLanguage: String</span></span>

<span data-ttu-id="858d0-191">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="858d0-191">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="858d0-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="858d0-192">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="858d0-193">Тип</span><span class="sxs-lookup"><span data-stu-id="858d0-193">Type</span></span>

*   <span data-ttu-id="858d0-194">String</span><span class="sxs-lookup"><span data-stu-id="858d0-194">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="858d0-195">Требования</span><span class="sxs-lookup"><span data-stu-id="858d0-195">Requirements</span></span>

|<span data-ttu-id="858d0-196">Требование</span><span class="sxs-lookup"><span data-stu-id="858d0-196">Requirement</span></span>| <span data-ttu-id="858d0-197">Значение</span><span class="sxs-lookup"><span data-stu-id="858d0-197">Value</span></span>|
|---|---|
|[<span data-ttu-id="858d0-198">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="858d0-198">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="858d0-199">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-199">1.1</span></span>|
|[<span data-ttu-id="858d0-200">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="858d0-200">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="858d0-201">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-201">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="858d0-202">Пример</span><span class="sxs-lookup"><span data-stu-id="858d0-202">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="858d0-203">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="858d0-203">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="858d0-204">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="858d0-204">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="858d0-205">Тип</span><span class="sxs-lookup"><span data-stu-id="858d0-205">Type</span></span>

*   [<span data-ttu-id="858d0-206">HostType</span><span class="sxs-lookup"><span data-stu-id="858d0-206">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="858d0-207">Requirements</span><span class="sxs-lookup"><span data-stu-id="858d0-207">Requirements</span></span>

|<span data-ttu-id="858d0-208">Требование</span><span class="sxs-lookup"><span data-stu-id="858d0-208">Requirement</span></span>| <span data-ttu-id="858d0-209">Значение</span><span class="sxs-lookup"><span data-stu-id="858d0-209">Value</span></span>|
|---|---|
|[<span data-ttu-id="858d0-210">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="858d0-210">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="858d0-211">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-211">1.1</span></span>|
|[<span data-ttu-id="858d0-212">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="858d0-212">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="858d0-213">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-213">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="858d0-214">Пример</span><span class="sxs-lookup"><span data-stu-id="858d0-214">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="858d0-215">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="858d0-215">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="858d0-216">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="858d0-216">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="858d0-217">Тип</span><span class="sxs-lookup"><span data-stu-id="858d0-217">Type</span></span>

*   [<span data-ttu-id="858d0-218">PlatformType</span><span class="sxs-lookup"><span data-stu-id="858d0-218">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="858d0-219">Requirements</span><span class="sxs-lookup"><span data-stu-id="858d0-219">Requirements</span></span>

|<span data-ttu-id="858d0-220">Требование</span><span class="sxs-lookup"><span data-stu-id="858d0-220">Requirement</span></span>| <span data-ttu-id="858d0-221">Значение</span><span class="sxs-lookup"><span data-stu-id="858d0-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="858d0-222">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="858d0-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="858d0-223">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-223">1.1</span></span>|
|[<span data-ttu-id="858d0-224">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="858d0-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="858d0-225">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="858d0-226">Пример</span><span class="sxs-lookup"><span data-stu-id="858d0-226">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="858d0-227">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="858d0-227">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="858d0-228">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="858d0-228">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="858d0-229">Тип</span><span class="sxs-lookup"><span data-stu-id="858d0-229">Type</span></span>

*   [<span data-ttu-id="858d0-230">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="858d0-230">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="858d0-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="858d0-231">Requirements</span></span>

|<span data-ttu-id="858d0-232">Требование</span><span class="sxs-lookup"><span data-stu-id="858d0-232">Requirement</span></span>| <span data-ttu-id="858d0-233">Значение</span><span class="sxs-lookup"><span data-stu-id="858d0-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="858d0-234">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="858d0-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="858d0-235">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-235">1.1</span></span>|
|[<span data-ttu-id="858d0-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="858d0-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="858d0-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="858d0-238">Пример</span><span class="sxs-lookup"><span data-stu-id="858d0-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="858d0-239">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="858d0-239">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="858d0-240">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="858d0-240">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="858d0-241">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="858d0-241">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="858d0-242">Тип</span><span class="sxs-lookup"><span data-stu-id="858d0-242">Type</span></span>

*   [<span data-ttu-id="858d0-243">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="858d0-243">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="858d0-244">Requirements</span><span class="sxs-lookup"><span data-stu-id="858d0-244">Requirements</span></span>

|<span data-ttu-id="858d0-245">Требование</span><span class="sxs-lookup"><span data-stu-id="858d0-245">Requirement</span></span>| <span data-ttu-id="858d0-246">Значение</span><span class="sxs-lookup"><span data-stu-id="858d0-246">Value</span></span>|
|---|---|
|[<span data-ttu-id="858d0-247">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="858d0-247">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="858d0-248">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-248">1.1</span></span>|
|[<span data-ttu-id="858d0-249">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="858d0-249">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="858d0-250">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="858d0-250">Restricted</span></span>|
|[<span data-ttu-id="858d0-251">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="858d0-251">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="858d0-252">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-252">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="858d0-253">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="858d0-253">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="858d0-254">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="858d0-254">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="858d0-255">Тип</span><span class="sxs-lookup"><span data-stu-id="858d0-255">Type</span></span>

*   [<span data-ttu-id="858d0-256">UI</span><span class="sxs-lookup"><span data-stu-id="858d0-256">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="858d0-257">Requirements</span><span class="sxs-lookup"><span data-stu-id="858d0-257">Requirements</span></span>

|<span data-ttu-id="858d0-258">Требование</span><span class="sxs-lookup"><span data-stu-id="858d0-258">Requirement</span></span>| <span data-ttu-id="858d0-259">Значение</span><span class="sxs-lookup"><span data-stu-id="858d0-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="858d0-260">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="858d0-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="858d0-261">1.1</span><span class="sxs-lookup"><span data-stu-id="858d0-261">1.1</span></span>|
|[<span data-ttu-id="858d0-262">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="858d0-262">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="858d0-263">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="858d0-263">Compose or Read</span></span>|
