---
title: Office. Context — набор обязательных элементов 1,4
description: Объектная модель для объекта контекста Outlook в API надстроек Outlook (версия API почтовых ящиков 1,4).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: d2d5520d4b48700b2e232e0782583ebd5f2ce6c3
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717728"
---
# <a name="context"></a><span data-ttu-id="9a9f8-103">context</span><span class="sxs-lookup"><span data-stu-id="9a9f8-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="9a9f8-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="9a9f8-104">[Office](office.md).context</span></span>

<span data-ttu-id="9a9f8-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="9a9f8-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.4).</span><span class="sxs-lookup"><span data-stu-id="9a9f8-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a9f8-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a9f8-107">Requirements</span></span>

|<span data-ttu-id="9a9f8-108">Требование</span><span class="sxs-lookup"><span data-stu-id="9a9f8-108">Requirement</span></span>| <span data-ttu-id="9a9f8-109">Значение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a9f8-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9a9f8-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a9f8-111">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-111">1.1</span></span>|
|[<span data-ttu-id="9a9f8-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9a9f8-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a9f8-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="9a9f8-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="9a9f8-114">Properties</span></span>

| <span data-ttu-id="9a9f8-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="9a9f8-115">Property</span></span> | <span data-ttu-id="9a9f8-116">Способов</span><span class="sxs-lookup"><span data-stu-id="9a9f8-116">Modes</span></span> | <span data-ttu-id="9a9f8-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="9a9f8-117">Return type</span></span> | <span data-ttu-id="9a9f8-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="9a9f8-118">Minimum</span></span><br><span data-ttu-id="9a9f8-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="9a9f8-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9a9f8-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="9a9f8-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="9a9f8-121">Создание</span><span class="sxs-lookup"><span data-stu-id="9a9f8-121">Compose</span></span><br><span data-ttu-id="9a9f8-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-122">Read</span></span> | <span data-ttu-id="9a9f8-123">String</span><span class="sxs-lookup"><span data-stu-id="9a9f8-123">String</span></span> | [<span data-ttu-id="9a9f8-124">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9a9f8-125">diagnostics</span><span class="sxs-lookup"><span data-stu-id="9a9f8-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="9a9f8-126">Создание</span><span class="sxs-lookup"><span data-stu-id="9a9f8-126">Compose</span></span><br><span data-ttu-id="9a9f8-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-127">Read</span></span> | [<span data-ttu-id="9a9f8-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="9a9f8-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4) | [<span data-ttu-id="9a9f8-129">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9a9f8-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="9a9f8-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="9a9f8-131">Создание</span><span class="sxs-lookup"><span data-stu-id="9a9f8-131">Compose</span></span><br><span data-ttu-id="9a9f8-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-132">Read</span></span> | <span data-ttu-id="9a9f8-133">String</span><span class="sxs-lookup"><span data-stu-id="9a9f8-133">String</span></span> | [<span data-ttu-id="9a9f8-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9a9f8-135">размещать</span><span class="sxs-lookup"><span data-stu-id="9a9f8-135">host</span></span>](#host-hosttype) | <span data-ttu-id="9a9f8-136">Создание</span><span class="sxs-lookup"><span data-stu-id="9a9f8-136">Compose</span></span><br><span data-ttu-id="9a9f8-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-137">Read</span></span> | [<span data-ttu-id="9a9f8-138">HostType</span><span class="sxs-lookup"><span data-stu-id="9a9f8-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.4) | [<span data-ttu-id="9a9f8-139">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9a9f8-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="9a9f8-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="9a9f8-141">Создание</span><span class="sxs-lookup"><span data-stu-id="9a9f8-141">Compose</span></span><br><span data-ttu-id="9a9f8-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-142">Read</span></span> | [<span data-ttu-id="9a9f8-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="9a9f8-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4) | [<span data-ttu-id="9a9f8-144">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9a9f8-145">управляем</span><span class="sxs-lookup"><span data-stu-id="9a9f8-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="9a9f8-146">Создание</span><span class="sxs-lookup"><span data-stu-id="9a9f8-146">Compose</span></span><br><span data-ttu-id="9a9f8-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-147">Read</span></span> | [<span data-ttu-id="9a9f8-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9a9f8-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.4) | [<span data-ttu-id="9a9f8-149">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9a9f8-150">потребность</span><span class="sxs-lookup"><span data-stu-id="9a9f8-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="9a9f8-151">Создание</span><span class="sxs-lookup"><span data-stu-id="9a9f8-151">Compose</span></span><br><span data-ttu-id="9a9f8-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-152">Read</span></span> | [<span data-ttu-id="9a9f8-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="9a9f8-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4) | [<span data-ttu-id="9a9f8-154">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9a9f8-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="9a9f8-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="9a9f8-156">Создание</span><span class="sxs-lookup"><span data-stu-id="9a9f8-156">Compose</span></span><br><span data-ttu-id="9a9f8-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-157">Read</span></span> | [<span data-ttu-id="9a9f8-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9a9f8-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4) | [<span data-ttu-id="9a9f8-159">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9a9f8-160">ui</span><span class="sxs-lookup"><span data-stu-id="9a9f8-160">ui</span></span>](#ui-ui) | <span data-ttu-id="9a9f8-161">Создание</span><span class="sxs-lookup"><span data-stu-id="9a9f8-161">Compose</span></span><br><span data-ttu-id="9a9f8-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-162">Read</span></span> | [<span data-ttu-id="9a9f8-163">UI</span><span class="sxs-lookup"><span data-stu-id="9a9f8-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4) | [<span data-ttu-id="9a9f8-164">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="9a9f8-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="9a9f8-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="9a9f8-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="9a9f8-166">contentLanguage: String</span></span>

<span data-ttu-id="9a9f8-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="9a9f8-168">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="9a9f8-169">Тип</span><span class="sxs-lookup"><span data-stu-id="9a9f8-169">Type</span></span>

*   <span data-ttu-id="9a9f8-170">String</span><span class="sxs-lookup"><span data-stu-id="9a9f8-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a9f8-171">Требования</span><span class="sxs-lookup"><span data-stu-id="9a9f8-171">Requirements</span></span>

|<span data-ttu-id="9a9f8-172">Требование</span><span class="sxs-lookup"><span data-stu-id="9a9f8-172">Requirement</span></span>| <span data-ttu-id="9a9f8-173">Значение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a9f8-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9a9f8-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a9f8-175">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-175">1.1</span></span>|
|[<span data-ttu-id="9a9f8-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9a9f8-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a9f8-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a9f8-178">Пример</span><span class="sxs-lookup"><span data-stu-id="9a9f8-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="9a9f8-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="9a9f8-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="9a9f8-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9a9f8-181">Тип</span><span class="sxs-lookup"><span data-stu-id="9a9f8-181">Type</span></span>

*   [<span data-ttu-id="9a9f8-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="9a9f8-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="9a9f8-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a9f8-183">Requirements</span></span>

|<span data-ttu-id="9a9f8-184">Требование</span><span class="sxs-lookup"><span data-stu-id="9a9f8-184">Requirement</span></span>| <span data-ttu-id="9a9f8-185">Значение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a9f8-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9a9f8-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a9f8-187">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-187">1.1</span></span>|
|[<span data-ttu-id="9a9f8-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9a9f8-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a9f8-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a9f8-190">Пример</span><span class="sxs-lookup"><span data-stu-id="9a9f8-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="9a9f8-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="9a9f8-191">displayLanguage: String</span></span>

<span data-ttu-id="9a9f8-192">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="9a9f8-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="9a9f8-194">Тип</span><span class="sxs-lookup"><span data-stu-id="9a9f8-194">Type</span></span>

*   <span data-ttu-id="9a9f8-195">String</span><span class="sxs-lookup"><span data-stu-id="9a9f8-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9a9f8-196">Требования</span><span class="sxs-lookup"><span data-stu-id="9a9f8-196">Requirements</span></span>

|<span data-ttu-id="9a9f8-197">Требование</span><span class="sxs-lookup"><span data-stu-id="9a9f8-197">Requirement</span></span>| <span data-ttu-id="9a9f8-198">Значение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a9f8-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9a9f8-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a9f8-200">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-200">1.1</span></span>|
|[<span data-ttu-id="9a9f8-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9a9f8-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a9f8-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a9f8-203">Пример</span><span class="sxs-lookup"><span data-stu-id="9a9f8-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="9a9f8-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="9a9f8-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="9a9f8-205">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-205">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9a9f8-206">Тип</span><span class="sxs-lookup"><span data-stu-id="9a9f8-206">Type</span></span>

*   [<span data-ttu-id="9a9f8-207">HostType</span><span class="sxs-lookup"><span data-stu-id="9a9f8-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="9a9f8-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a9f8-208">Requirements</span></span>

|<span data-ttu-id="9a9f8-209">Требование</span><span class="sxs-lookup"><span data-stu-id="9a9f8-209">Requirement</span></span>| <span data-ttu-id="9a9f8-210">Значение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a9f8-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9a9f8-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a9f8-212">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-212">1.1</span></span>|
|[<span data-ttu-id="9a9f8-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9a9f8-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a9f8-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a9f8-215">Пример</span><span class="sxs-lookup"><span data-stu-id="9a9f8-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="9a9f8-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="9a9f8-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="9a9f8-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9a9f8-218">Тип</span><span class="sxs-lookup"><span data-stu-id="9a9f8-218">Type</span></span>

*   [<span data-ttu-id="9a9f8-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9a9f8-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="9a9f8-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a9f8-220">Requirements</span></span>

|<span data-ttu-id="9a9f8-221">Требование</span><span class="sxs-lookup"><span data-stu-id="9a9f8-221">Requirement</span></span>| <span data-ttu-id="9a9f8-222">Значение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a9f8-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9a9f8-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a9f8-224">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-224">1.1</span></span>|
|[<span data-ttu-id="9a9f8-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9a9f8-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a9f8-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a9f8-227">Пример</span><span class="sxs-lookup"><span data-stu-id="9a9f8-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="9a9f8-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="9a9f8-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="9a9f8-229">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-229">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9a9f8-230">Тип</span><span class="sxs-lookup"><span data-stu-id="9a9f8-230">Type</span></span>

*   [<span data-ttu-id="9a9f8-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="9a9f8-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="9a9f8-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a9f8-232">Requirements</span></span>

|<span data-ttu-id="9a9f8-233">Требование</span><span class="sxs-lookup"><span data-stu-id="9a9f8-233">Requirement</span></span>| <span data-ttu-id="9a9f8-234">Значение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a9f8-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9a9f8-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a9f8-236">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-236">1.1</span></span>|
|[<span data-ttu-id="9a9f8-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9a9f8-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a9f8-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9a9f8-239">Пример</span><span class="sxs-lookup"><span data-stu-id="9a9f8-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="9a9f8-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="9a9f8-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="9a9f8-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="9a9f8-242">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="9a9f8-243">Тип</span><span class="sxs-lookup"><span data-stu-id="9a9f8-243">Type</span></span>

*   [<span data-ttu-id="9a9f8-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9a9f8-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="9a9f8-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a9f8-245">Requirements</span></span>

|<span data-ttu-id="9a9f8-246">Требование</span><span class="sxs-lookup"><span data-stu-id="9a9f8-246">Requirement</span></span>| <span data-ttu-id="9a9f8-247">Значение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a9f8-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9a9f8-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a9f8-249">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-249">1.1</span></span>|
|[<span data-ttu-id="9a9f8-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9a9f8-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="9a9f8-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9a9f8-251">Restricted</span></span>|
|[<span data-ttu-id="9a9f8-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9a9f8-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a9f8-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="9a9f8-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="9a9f8-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="9a9f8-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="9a9f8-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="9a9f8-256">Тип</span><span class="sxs-lookup"><span data-stu-id="9a9f8-256">Type</span></span>

*   [<span data-ttu-id="9a9f8-257">UI</span><span class="sxs-lookup"><span data-stu-id="9a9f8-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="9a9f8-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="9a9f8-258">Requirements</span></span>

|<span data-ttu-id="9a9f8-259">Требование</span><span class="sxs-lookup"><span data-stu-id="9a9f8-259">Requirement</span></span>| <span data-ttu-id="9a9f8-260">Значение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="9a9f8-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9a9f8-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9a9f8-262">1.1</span><span class="sxs-lookup"><span data-stu-id="9a9f8-262">1.1</span></span>|
|[<span data-ttu-id="9a9f8-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9a9f8-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9a9f8-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9a9f8-264">Compose or Read</span></span>|
