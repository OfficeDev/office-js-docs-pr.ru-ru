---
title: Office. Context — набор обязательных элементов 1,6
description: Объектная модель для объекта контекста Outlook в API надстроек Outlook (версия API почтовых ящиков 1,6).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 8ce9b9c9ce1a0e224f57c899ddf065b22d052cd0
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717644"
---
# <a name="context"></a><span data-ttu-id="c9d17-103">context</span><span class="sxs-lookup"><span data-stu-id="c9d17-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="c9d17-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="c9d17-104">[Office](office.md).context</span></span>

<span data-ttu-id="c9d17-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="c9d17-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="c9d17-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.6).</span><span class="sxs-lookup"><span data-stu-id="c9d17-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9d17-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9d17-107">Requirements</span></span>

|<span data-ttu-id="c9d17-108">Требование</span><span class="sxs-lookup"><span data-stu-id="c9d17-108">Requirement</span></span>| <span data-ttu-id="c9d17-109">Значение</span><span class="sxs-lookup"><span data-stu-id="c9d17-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9d17-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c9d17-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c9d17-111">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-111">1.1</span></span>|
|[<span data-ttu-id="c9d17-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c9d17-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c9d17-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="c9d17-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="c9d17-114">Properties</span></span>

| <span data-ttu-id="c9d17-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="c9d17-115">Property</span></span> | <span data-ttu-id="c9d17-116">Способов</span><span class="sxs-lookup"><span data-stu-id="c9d17-116">Modes</span></span> | <span data-ttu-id="c9d17-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="c9d17-117">Return type</span></span> | <span data-ttu-id="c9d17-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="c9d17-118">Minimum</span></span><br><span data-ttu-id="c9d17-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="c9d17-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c9d17-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="c9d17-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="c9d17-121">Создание</span><span class="sxs-lookup"><span data-stu-id="c9d17-121">Compose</span></span><br><span data-ttu-id="c9d17-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-122">Read</span></span> | <span data-ttu-id="c9d17-123">String</span><span class="sxs-lookup"><span data-stu-id="c9d17-123">String</span></span> | [<span data-ttu-id="c9d17-124">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c9d17-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="c9d17-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="c9d17-126">Создание</span><span class="sxs-lookup"><span data-stu-id="c9d17-126">Compose</span></span><br><span data-ttu-id="c9d17-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-127">Read</span></span> | [<span data-ttu-id="c9d17-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="c9d17-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6) | [<span data-ttu-id="c9d17-129">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c9d17-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="c9d17-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="c9d17-131">Создание</span><span class="sxs-lookup"><span data-stu-id="c9d17-131">Compose</span></span><br><span data-ttu-id="c9d17-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-132">Read</span></span> | <span data-ttu-id="c9d17-133">String</span><span class="sxs-lookup"><span data-stu-id="c9d17-133">String</span></span> | [<span data-ttu-id="c9d17-134">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c9d17-135">размещать</span><span class="sxs-lookup"><span data-stu-id="c9d17-135">host</span></span>](#host-hosttype) | <span data-ttu-id="c9d17-136">Создание</span><span class="sxs-lookup"><span data-stu-id="c9d17-136">Compose</span></span><br><span data-ttu-id="c9d17-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-137">Read</span></span> | [<span data-ttu-id="c9d17-138">HostType</span><span class="sxs-lookup"><span data-stu-id="c9d17-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6) | [<span data-ttu-id="c9d17-139">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c9d17-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="c9d17-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="c9d17-141">Создание</span><span class="sxs-lookup"><span data-stu-id="c9d17-141">Compose</span></span><br><span data-ttu-id="c9d17-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-142">Read</span></span> | [<span data-ttu-id="c9d17-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="c9d17-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6) | [<span data-ttu-id="c9d17-144">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c9d17-145">управляем</span><span class="sxs-lookup"><span data-stu-id="c9d17-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="c9d17-146">Создание</span><span class="sxs-lookup"><span data-stu-id="c9d17-146">Compose</span></span><br><span data-ttu-id="c9d17-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-147">Read</span></span> | [<span data-ttu-id="c9d17-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c9d17-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6) | [<span data-ttu-id="c9d17-149">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c9d17-150">потребность</span><span class="sxs-lookup"><span data-stu-id="c9d17-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="c9d17-151">Создание</span><span class="sxs-lookup"><span data-stu-id="c9d17-151">Compose</span></span><br><span data-ttu-id="c9d17-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-152">Read</span></span> | [<span data-ttu-id="c9d17-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="c9d17-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6) | [<span data-ttu-id="c9d17-154">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c9d17-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="c9d17-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="c9d17-156">Создание</span><span class="sxs-lookup"><span data-stu-id="c9d17-156">Compose</span></span><br><span data-ttu-id="c9d17-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-157">Read</span></span> | [<span data-ttu-id="c9d17-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c9d17-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6) | [<span data-ttu-id="c9d17-159">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c9d17-160">ui</span><span class="sxs-lookup"><span data-stu-id="c9d17-160">ui</span></span>](#ui-ui) | <span data-ttu-id="c9d17-161">Создание</span><span class="sxs-lookup"><span data-stu-id="c9d17-161">Compose</span></span><br><span data-ttu-id="c9d17-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-162">Read</span></span> | [<span data-ttu-id="c9d17-163">UI</span><span class="sxs-lookup"><span data-stu-id="c9d17-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6) | [<span data-ttu-id="c9d17-164">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="c9d17-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="c9d17-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="c9d17-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="c9d17-166">contentLanguage: String</span></span>

<span data-ttu-id="c9d17-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="c9d17-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="c9d17-168">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="c9d17-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c9d17-169">Тип</span><span class="sxs-lookup"><span data-stu-id="c9d17-169">Type</span></span>

*   <span data-ttu-id="c9d17-170">String</span><span class="sxs-lookup"><span data-stu-id="c9d17-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9d17-171">Требования</span><span class="sxs-lookup"><span data-stu-id="c9d17-171">Requirements</span></span>

|<span data-ttu-id="c9d17-172">Требование</span><span class="sxs-lookup"><span data-stu-id="c9d17-172">Requirement</span></span>| <span data-ttu-id="c9d17-173">Значение</span><span class="sxs-lookup"><span data-stu-id="c9d17-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9d17-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c9d17-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c9d17-175">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-175">1.1</span></span>|
|[<span data-ttu-id="c9d17-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c9d17-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c9d17-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9d17-178">Пример</span><span class="sxs-lookup"><span data-stu-id="c9d17-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="c9d17-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="c9d17-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="c9d17-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="c9d17-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c9d17-181">Тип</span><span class="sxs-lookup"><span data-stu-id="c9d17-181">Type</span></span>

*   [<span data-ttu-id="c9d17-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="c9d17-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="c9d17-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9d17-183">Requirements</span></span>

|<span data-ttu-id="c9d17-184">Требование</span><span class="sxs-lookup"><span data-stu-id="c9d17-184">Requirement</span></span>| <span data-ttu-id="c9d17-185">Значение</span><span class="sxs-lookup"><span data-stu-id="c9d17-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9d17-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c9d17-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c9d17-187">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-187">1.1</span></span>|
|[<span data-ttu-id="c9d17-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c9d17-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c9d17-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9d17-190">Пример</span><span class="sxs-lookup"><span data-stu-id="c9d17-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="c9d17-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="c9d17-191">displayLanguage: String</span></span>

<span data-ttu-id="c9d17-192">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="c9d17-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="c9d17-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="c9d17-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="c9d17-194">Тип</span><span class="sxs-lookup"><span data-stu-id="c9d17-194">Type</span></span>

*   <span data-ttu-id="c9d17-195">String</span><span class="sxs-lookup"><span data-stu-id="c9d17-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c9d17-196">Требования</span><span class="sxs-lookup"><span data-stu-id="c9d17-196">Requirements</span></span>

|<span data-ttu-id="c9d17-197">Требование</span><span class="sxs-lookup"><span data-stu-id="c9d17-197">Requirement</span></span>| <span data-ttu-id="c9d17-198">Значение</span><span class="sxs-lookup"><span data-stu-id="c9d17-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9d17-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c9d17-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c9d17-200">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-200">1.1</span></span>|
|[<span data-ttu-id="c9d17-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c9d17-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c9d17-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9d17-203">Пример</span><span class="sxs-lookup"><span data-stu-id="c9d17-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="c9d17-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="c9d17-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="c9d17-205">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="c9d17-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c9d17-206">Тип</span><span class="sxs-lookup"><span data-stu-id="c9d17-206">Type</span></span>

*   [<span data-ttu-id="c9d17-207">HostType</span><span class="sxs-lookup"><span data-stu-id="c9d17-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="c9d17-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9d17-208">Requirements</span></span>

|<span data-ttu-id="c9d17-209">Требование</span><span class="sxs-lookup"><span data-stu-id="c9d17-209">Requirement</span></span>| <span data-ttu-id="c9d17-210">Значение</span><span class="sxs-lookup"><span data-stu-id="c9d17-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9d17-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c9d17-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c9d17-212">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-212">1.1</span></span>|
|[<span data-ttu-id="c9d17-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c9d17-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c9d17-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9d17-215">Пример</span><span class="sxs-lookup"><span data-stu-id="c9d17-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="c9d17-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="c9d17-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="c9d17-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="c9d17-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c9d17-218">Тип</span><span class="sxs-lookup"><span data-stu-id="c9d17-218">Type</span></span>

*   [<span data-ttu-id="c9d17-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="c9d17-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="c9d17-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9d17-220">Requirements</span></span>

|<span data-ttu-id="c9d17-221">Требование</span><span class="sxs-lookup"><span data-stu-id="c9d17-221">Requirement</span></span>| <span data-ttu-id="c9d17-222">Значение</span><span class="sxs-lookup"><span data-stu-id="c9d17-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9d17-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c9d17-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c9d17-224">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-224">1.1</span></span>|
|[<span data-ttu-id="c9d17-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c9d17-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c9d17-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9d17-227">Пример</span><span class="sxs-lookup"><span data-stu-id="c9d17-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="c9d17-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="c9d17-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="c9d17-229">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="c9d17-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="c9d17-230">Тип</span><span class="sxs-lookup"><span data-stu-id="c9d17-230">Type</span></span>

*   [<span data-ttu-id="c9d17-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="c9d17-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="c9d17-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9d17-232">Requirements</span></span>

|<span data-ttu-id="c9d17-233">Требование</span><span class="sxs-lookup"><span data-stu-id="c9d17-233">Requirement</span></span>| <span data-ttu-id="c9d17-234">Значение</span><span class="sxs-lookup"><span data-stu-id="c9d17-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9d17-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c9d17-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c9d17-236">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-236">1.1</span></span>|
|[<span data-ttu-id="c9d17-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c9d17-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c9d17-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c9d17-239">Пример</span><span class="sxs-lookup"><span data-stu-id="c9d17-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="c9d17-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="c9d17-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="c9d17-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="c9d17-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c9d17-242">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="c9d17-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c9d17-243">Тип</span><span class="sxs-lookup"><span data-stu-id="c9d17-243">Type</span></span>

*   [<span data-ttu-id="c9d17-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c9d17-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c9d17-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9d17-245">Requirements</span></span>

|<span data-ttu-id="c9d17-246">Требование</span><span class="sxs-lookup"><span data-stu-id="c9d17-246">Requirement</span></span>| <span data-ttu-id="c9d17-247">Значение</span><span class="sxs-lookup"><span data-stu-id="c9d17-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9d17-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c9d17-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c9d17-249">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-249">1.1</span></span>|
|[<span data-ttu-id="c9d17-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c9d17-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="c9d17-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="c9d17-251">Restricted</span></span>|
|[<span data-ttu-id="c9d17-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c9d17-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c9d17-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="c9d17-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="c9d17-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="c9d17-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="c9d17-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c9d17-256">Тип</span><span class="sxs-lookup"><span data-stu-id="c9d17-256">Type</span></span>

*   [<span data-ttu-id="c9d17-257">UI</span><span class="sxs-lookup"><span data-stu-id="c9d17-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="c9d17-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="c9d17-258">Requirements</span></span>

|<span data-ttu-id="c9d17-259">Требование</span><span class="sxs-lookup"><span data-stu-id="c9d17-259">Requirement</span></span>| <span data-ttu-id="c9d17-260">Значение</span><span class="sxs-lookup"><span data-stu-id="c9d17-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="c9d17-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c9d17-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c9d17-262">1.1</span><span class="sxs-lookup"><span data-stu-id="c9d17-262">1.1</span></span>|
|[<span data-ttu-id="c9d17-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c9d17-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c9d17-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c9d17-264">Compose or Read</span></span>|
