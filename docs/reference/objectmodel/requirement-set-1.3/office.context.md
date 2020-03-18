---
title: Office. Context — набор обязательных элементов 1,3
description: Объектная модель для объекта контекста Outlook в API надстроек Outlook (версия API почтовых ящиков 1,3).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: dc524c5657380845744e241a600fbf66f1eed5ce
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42720122"
---
# <a name="context"></a><span data-ttu-id="9db13-103">context</span><span class="sxs-lookup"><span data-stu-id="9db13-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="9db13-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="9db13-104">[Office](office.md).context</span></span>

<span data-ttu-id="9db13-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="9db13-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="9db13-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.3).</span><span class="sxs-lookup"><span data-stu-id="9db13-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.3).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9db13-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="9db13-107">Requirements</span></span>

|<span data-ttu-id="9db13-108">Требование</span><span class="sxs-lookup"><span data-stu-id="9db13-108">Requirement</span></span>| <span data-ttu-id="9db13-109">Значение</span><span class="sxs-lookup"><span data-stu-id="9db13-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9db13-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9db13-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9db13-111">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-111">1.1</span></span>|
|[<span data-ttu-id="9db13-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9db13-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9db13-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="9db13-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="9db13-114">Properties</span></span>

| <span data-ttu-id="9db13-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="9db13-115">Property</span></span> | <span data-ttu-id="9db13-116">Способов</span><span class="sxs-lookup"><span data-stu-id="9db13-116">Modes</span></span> | <span data-ttu-id="9db13-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="9db13-117">Return type</span></span> | <span data-ttu-id="9db13-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="9db13-118">Minimum</span></span><br><span data-ttu-id="9db13-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="9db13-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9db13-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="9db13-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="9db13-121">Создание</span><span class="sxs-lookup"><span data-stu-id="9db13-121">Compose</span></span><br><span data-ttu-id="9db13-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-122">Read</span></span> | <span data-ttu-id="9db13-123">String</span><span class="sxs-lookup"><span data-stu-id="9db13-123">String</span></span> | [<span data-ttu-id="9db13-124">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9db13-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="9db13-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="9db13-126">Создание</span><span class="sxs-lookup"><span data-stu-id="9db13-126">Compose</span></span><br><span data-ttu-id="9db13-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-127">Read</span></span> | [<span data-ttu-id="9db13-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="9db13-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.3) | [<span data-ttu-id="9db13-129">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9db13-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="9db13-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="9db13-131">Создание</span><span class="sxs-lookup"><span data-stu-id="9db13-131">Compose</span></span><br><span data-ttu-id="9db13-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-132">Read</span></span> | <span data-ttu-id="9db13-133">String</span><span class="sxs-lookup"><span data-stu-id="9db13-133">String</span></span> | [<span data-ttu-id="9db13-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9db13-135">размещать</span><span class="sxs-lookup"><span data-stu-id="9db13-135">host</span></span>](#host-hosttype) | <span data-ttu-id="9db13-136">Создание</span><span class="sxs-lookup"><span data-stu-id="9db13-136">Compose</span></span><br><span data-ttu-id="9db13-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-137">Read</span></span> | [<span data-ttu-id="9db13-138">HostType</span><span class="sxs-lookup"><span data-stu-id="9db13-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.3) | [<span data-ttu-id="9db13-139">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9db13-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="9db13-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="9db13-141">Создание</span><span class="sxs-lookup"><span data-stu-id="9db13-141">Compose</span></span><br><span data-ttu-id="9db13-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-142">Read</span></span> | [<span data-ttu-id="9db13-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="9db13-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.3) | [<span data-ttu-id="9db13-144">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9db13-145">управляем</span><span class="sxs-lookup"><span data-stu-id="9db13-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="9db13-146">Создание</span><span class="sxs-lookup"><span data-stu-id="9db13-146">Compose</span></span><br><span data-ttu-id="9db13-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-147">Read</span></span> | [<span data-ttu-id="9db13-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9db13-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.3) | [<span data-ttu-id="9db13-149">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9db13-150">потребность</span><span class="sxs-lookup"><span data-stu-id="9db13-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="9db13-151">Создание</span><span class="sxs-lookup"><span data-stu-id="9db13-151">Compose</span></span><br><span data-ttu-id="9db13-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-152">Read</span></span> | [<span data-ttu-id="9db13-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="9db13-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.3) | [<span data-ttu-id="9db13-154">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9db13-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="9db13-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="9db13-156">Создание</span><span class="sxs-lookup"><span data-stu-id="9db13-156">Compose</span></span><br><span data-ttu-id="9db13-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-157">Read</span></span> | [<span data-ttu-id="9db13-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9db13-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.3) | [<span data-ttu-id="9db13-159">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9db13-160">ui</span><span class="sxs-lookup"><span data-stu-id="9db13-160">ui</span></span>](#ui-ui) | <span data-ttu-id="9db13-161">Создание</span><span class="sxs-lookup"><span data-stu-id="9db13-161">Compose</span></span><br><span data-ttu-id="9db13-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-162">Read</span></span> | [<span data-ttu-id="9db13-163">UI</span><span class="sxs-lookup"><span data-stu-id="9db13-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.3) | [<span data-ttu-id="9db13-164">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="9db13-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="9db13-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="9db13-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="9db13-166">contentLanguage: String</span></span>

<span data-ttu-id="9db13-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="9db13-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="9db13-168">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="9db13-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="9db13-169">Тип</span><span class="sxs-lookup"><span data-stu-id="9db13-169">Type</span></span>

*   <span data-ttu-id="9db13-170">String</span><span class="sxs-lookup"><span data-stu-id="9db13-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9db13-171">Требования</span><span class="sxs-lookup"><span data-stu-id="9db13-171">Requirements</span></span>

|<span data-ttu-id="9db13-172">Требование</span><span class="sxs-lookup"><span data-stu-id="9db13-172">Requirement</span></span>| <span data-ttu-id="9db13-173">Значение</span><span class="sxs-lookup"><span data-stu-id="9db13-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="9db13-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9db13-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9db13-175">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-175">1.1</span></span>|
|[<span data-ttu-id="9db13-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9db13-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9db13-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9db13-178">Пример</span><span class="sxs-lookup"><span data-stu-id="9db13-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="9db13-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="9db13-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="9db13-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="9db13-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9db13-181">Тип</span><span class="sxs-lookup"><span data-stu-id="9db13-181">Type</span></span>

*   [<span data-ttu-id="9db13-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="9db13-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="9db13-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="9db13-183">Requirements</span></span>

|<span data-ttu-id="9db13-184">Требование</span><span class="sxs-lookup"><span data-stu-id="9db13-184">Requirement</span></span>| <span data-ttu-id="9db13-185">Значение</span><span class="sxs-lookup"><span data-stu-id="9db13-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="9db13-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9db13-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9db13-187">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-187">1.1</span></span>|
|[<span data-ttu-id="9db13-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9db13-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9db13-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9db13-190">Пример</span><span class="sxs-lookup"><span data-stu-id="9db13-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="9db13-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="9db13-191">displayLanguage: String</span></span>

<span data-ttu-id="9db13-192">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="9db13-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="9db13-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="9db13-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="9db13-194">Тип</span><span class="sxs-lookup"><span data-stu-id="9db13-194">Type</span></span>

*   <span data-ttu-id="9db13-195">String</span><span class="sxs-lookup"><span data-stu-id="9db13-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9db13-196">Требования</span><span class="sxs-lookup"><span data-stu-id="9db13-196">Requirements</span></span>

|<span data-ttu-id="9db13-197">Требование</span><span class="sxs-lookup"><span data-stu-id="9db13-197">Requirement</span></span>| <span data-ttu-id="9db13-198">Значение</span><span class="sxs-lookup"><span data-stu-id="9db13-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="9db13-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9db13-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9db13-200">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-200">1.1</span></span>|
|[<span data-ttu-id="9db13-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9db13-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9db13-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9db13-203">Пример</span><span class="sxs-lookup"><span data-stu-id="9db13-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="9db13-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="9db13-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="9db13-205">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="9db13-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9db13-206">Тип</span><span class="sxs-lookup"><span data-stu-id="9db13-206">Type</span></span>

*   [<span data-ttu-id="9db13-207">HostType</span><span class="sxs-lookup"><span data-stu-id="9db13-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="9db13-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="9db13-208">Requirements</span></span>

|<span data-ttu-id="9db13-209">Требование</span><span class="sxs-lookup"><span data-stu-id="9db13-209">Requirement</span></span>| <span data-ttu-id="9db13-210">Значение</span><span class="sxs-lookup"><span data-stu-id="9db13-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="9db13-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9db13-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9db13-212">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-212">1.1</span></span>|
|[<span data-ttu-id="9db13-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9db13-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9db13-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9db13-215">Пример</span><span class="sxs-lookup"><span data-stu-id="9db13-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="9db13-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="9db13-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="9db13-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="9db13-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9db13-218">Тип</span><span class="sxs-lookup"><span data-stu-id="9db13-218">Type</span></span>

*   [<span data-ttu-id="9db13-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9db13-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="9db13-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="9db13-220">Requirements</span></span>

|<span data-ttu-id="9db13-221">Требование</span><span class="sxs-lookup"><span data-stu-id="9db13-221">Requirement</span></span>| <span data-ttu-id="9db13-222">Значение</span><span class="sxs-lookup"><span data-stu-id="9db13-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="9db13-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9db13-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9db13-224">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-224">1.1</span></span>|
|[<span data-ttu-id="9db13-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9db13-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9db13-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9db13-227">Пример</span><span class="sxs-lookup"><span data-stu-id="9db13-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="9db13-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="9db13-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="9db13-229">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="9db13-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9db13-230">Тип</span><span class="sxs-lookup"><span data-stu-id="9db13-230">Type</span></span>

*   [<span data-ttu-id="9db13-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="9db13-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="9db13-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="9db13-232">Requirements</span></span>

|<span data-ttu-id="9db13-233">Требование</span><span class="sxs-lookup"><span data-stu-id="9db13-233">Requirement</span></span>| <span data-ttu-id="9db13-234">Значение</span><span class="sxs-lookup"><span data-stu-id="9db13-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="9db13-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9db13-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9db13-236">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-236">1.1</span></span>|
|[<span data-ttu-id="9db13-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9db13-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9db13-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9db13-239">Пример</span><span class="sxs-lookup"><span data-stu-id="9db13-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="9db13-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="9db13-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="9db13-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="9db13-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="9db13-242">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="9db13-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="9db13-243">Тип</span><span class="sxs-lookup"><span data-stu-id="9db13-243">Type</span></span>

*   [<span data-ttu-id="9db13-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9db13-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="9db13-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="9db13-245">Requirements</span></span>

|<span data-ttu-id="9db13-246">Требование</span><span class="sxs-lookup"><span data-stu-id="9db13-246">Requirement</span></span>| <span data-ttu-id="9db13-247">Значение</span><span class="sxs-lookup"><span data-stu-id="9db13-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="9db13-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9db13-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9db13-249">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-249">1.1</span></span>|
|[<span data-ttu-id="9db13-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9db13-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="9db13-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9db13-251">Restricted</span></span>|
|[<span data-ttu-id="9db13-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9db13-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9db13-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="9db13-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="9db13-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="9db13-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="9db13-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="9db13-256">Тип</span><span class="sxs-lookup"><span data-stu-id="9db13-256">Type</span></span>

*   [<span data-ttu-id="9db13-257">UI</span><span class="sxs-lookup"><span data-stu-id="9db13-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="9db13-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="9db13-258">Requirements</span></span>

|<span data-ttu-id="9db13-259">Требование</span><span class="sxs-lookup"><span data-stu-id="9db13-259">Requirement</span></span>| <span data-ttu-id="9db13-260">Значение</span><span class="sxs-lookup"><span data-stu-id="9db13-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="9db13-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9db13-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9db13-262">1.1</span><span class="sxs-lookup"><span data-stu-id="9db13-262">1.1</span></span>|
|[<span data-ttu-id="9db13-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9db13-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9db13-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9db13-264">Compose or Read</span></span>|
