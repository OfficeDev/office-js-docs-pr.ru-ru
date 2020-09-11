---
title: Office. Context — набор обязательных элементов 1,6
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,6.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: e8cfb6992b8a654a8f348a61ad8d581ffe887df5
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47430578"
---
# <a name="context-mailbox-requirement-set-16"></a><span data-ttu-id="9efd0-103">контекст (набор требований для почтового ящика 1,6)</span><span class="sxs-lookup"><span data-stu-id="9efd0-103">context (Mailbox requirement set 1.6)</span></span>

### <a name="officecontext"></a><span data-ttu-id="9efd0-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="9efd0-104">[Office](office.md).context</span></span>

<span data-ttu-id="9efd0-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="9efd0-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="9efd0-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="9efd0-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9efd0-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="9efd0-107">Requirements</span></span>

|<span data-ttu-id="9efd0-108">Требование</span><span class="sxs-lookup"><span data-stu-id="9efd0-108">Requirement</span></span>| <span data-ttu-id="9efd0-109">Значение</span><span class="sxs-lookup"><span data-stu-id="9efd0-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9efd0-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9efd0-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9efd0-111">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-111">1.1</span></span>|
|[<span data-ttu-id="9efd0-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9efd0-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9efd0-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="9efd0-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="9efd0-114">Properties</span></span>

| <span data-ttu-id="9efd0-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="9efd0-115">Property</span></span> | <span data-ttu-id="9efd0-116">Способов</span><span class="sxs-lookup"><span data-stu-id="9efd0-116">Modes</span></span> | <span data-ttu-id="9efd0-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="9efd0-117">Return type</span></span> | <span data-ttu-id="9efd0-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="9efd0-118">Minimum</span></span><br><span data-ttu-id="9efd0-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="9efd0-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9efd0-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="9efd0-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="9efd0-121">Создание</span><span class="sxs-lookup"><span data-stu-id="9efd0-121">Compose</span></span><br><span data-ttu-id="9efd0-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-122">Read</span></span> | <span data-ttu-id="9efd0-123">String</span><span class="sxs-lookup"><span data-stu-id="9efd0-123">String</span></span> | [<span data-ttu-id="9efd0-124">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9efd0-125">диагностики</span><span class="sxs-lookup"><span data-stu-id="9efd0-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="9efd0-126">Создание</span><span class="sxs-lookup"><span data-stu-id="9efd0-126">Compose</span></span><br><span data-ttu-id="9efd0-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-127">Read</span></span> | [<span data-ttu-id="9efd0-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="9efd0-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="9efd0-129">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9efd0-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="9efd0-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="9efd0-131">Создание</span><span class="sxs-lookup"><span data-stu-id="9efd0-131">Compose</span></span><br><span data-ttu-id="9efd0-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-132">Read</span></span> | <span data-ttu-id="9efd0-133">String</span><span class="sxs-lookup"><span data-stu-id="9efd0-133">String</span></span> | [<span data-ttu-id="9efd0-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9efd0-135">размещать</span><span class="sxs-lookup"><span data-stu-id="9efd0-135">host</span></span>](#host-hosttype) | <span data-ttu-id="9efd0-136">Создание</span><span class="sxs-lookup"><span data-stu-id="9efd0-136">Compose</span></span><br><span data-ttu-id="9efd0-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-137">Read</span></span> | [<span data-ttu-id="9efd0-138">HostType</span><span class="sxs-lookup"><span data-stu-id="9efd0-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="9efd0-139">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9efd0-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="9efd0-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="9efd0-141">Создание</span><span class="sxs-lookup"><span data-stu-id="9efd0-141">Compose</span></span><br><span data-ttu-id="9efd0-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-142">Read</span></span> | [<span data-ttu-id="9efd0-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="9efd0-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="9efd0-144">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9efd0-145">управляем</span><span class="sxs-lookup"><span data-stu-id="9efd0-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="9efd0-146">Создание</span><span class="sxs-lookup"><span data-stu-id="9efd0-146">Compose</span></span><br><span data-ttu-id="9efd0-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-147">Read</span></span> | [<span data-ttu-id="9efd0-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9efd0-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="9efd0-149">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9efd0-150">потребность</span><span class="sxs-lookup"><span data-stu-id="9efd0-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="9efd0-151">Создание</span><span class="sxs-lookup"><span data-stu-id="9efd0-151">Compose</span></span><br><span data-ttu-id="9efd0-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-152">Read</span></span> | [<span data-ttu-id="9efd0-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="9efd0-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="9efd0-154">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9efd0-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="9efd0-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="9efd0-156">Создание</span><span class="sxs-lookup"><span data-stu-id="9efd0-156">Compose</span></span><br><span data-ttu-id="9efd0-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-157">Read</span></span> | [<span data-ttu-id="9efd0-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9efd0-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="9efd0-159">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9efd0-160">ui</span><span class="sxs-lookup"><span data-stu-id="9efd0-160">ui</span></span>](#ui-ui) | <span data-ttu-id="9efd0-161">Создание</span><span class="sxs-lookup"><span data-stu-id="9efd0-161">Compose</span></span><br><span data-ttu-id="9efd0-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-162">Read</span></span> | [<span data-ttu-id="9efd0-163">UI</span><span class="sxs-lookup"><span data-stu-id="9efd0-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6&preserve-view=true) | [<span data-ttu-id="9efd0-164">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="9efd0-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="9efd0-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="9efd0-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="9efd0-166">contentLanguage: String</span></span>

<span data-ttu-id="9efd0-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="9efd0-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="9efd0-168">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="9efd0-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="9efd0-169">Тип</span><span class="sxs-lookup"><span data-stu-id="9efd0-169">Type</span></span>

*   <span data-ttu-id="9efd0-170">String</span><span class="sxs-lookup"><span data-stu-id="9efd0-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9efd0-171">Требования</span><span class="sxs-lookup"><span data-stu-id="9efd0-171">Requirements</span></span>

|<span data-ttu-id="9efd0-172">Требование</span><span class="sxs-lookup"><span data-stu-id="9efd0-172">Requirement</span></span>| <span data-ttu-id="9efd0-173">Значение</span><span class="sxs-lookup"><span data-stu-id="9efd0-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="9efd0-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9efd0-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9efd0-175">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-175">1.1</span></span>|
|[<span data-ttu-id="9efd0-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9efd0-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9efd0-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9efd0-178">Пример</span><span class="sxs-lookup"><span data-stu-id="9efd0-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="9efd0-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="9efd0-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="9efd0-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="9efd0-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9efd0-181">Тип</span><span class="sxs-lookup"><span data-stu-id="9efd0-181">Type</span></span>

*   [<span data-ttu-id="9efd0-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="9efd0-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="9efd0-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="9efd0-183">Requirements</span></span>

|<span data-ttu-id="9efd0-184">Требование</span><span class="sxs-lookup"><span data-stu-id="9efd0-184">Requirement</span></span>| <span data-ttu-id="9efd0-185">Значение</span><span class="sxs-lookup"><span data-stu-id="9efd0-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="9efd0-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9efd0-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9efd0-187">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-187">1.1</span></span>|
|[<span data-ttu-id="9efd0-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9efd0-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9efd0-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9efd0-190">Пример</span><span class="sxs-lookup"><span data-stu-id="9efd0-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="9efd0-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="9efd0-191">displayLanguage: String</span></span>

<span data-ttu-id="9efd0-192">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="9efd0-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="9efd0-193">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="9efd0-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="9efd0-194">Тип</span><span class="sxs-lookup"><span data-stu-id="9efd0-194">Type</span></span>

*   <span data-ttu-id="9efd0-195">String</span><span class="sxs-lookup"><span data-stu-id="9efd0-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9efd0-196">Требования</span><span class="sxs-lookup"><span data-stu-id="9efd0-196">Requirements</span></span>

|<span data-ttu-id="9efd0-197">Требование</span><span class="sxs-lookup"><span data-stu-id="9efd0-197">Requirement</span></span>| <span data-ttu-id="9efd0-198">Значение</span><span class="sxs-lookup"><span data-stu-id="9efd0-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="9efd0-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9efd0-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9efd0-200">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-200">1.1</span></span>|
|[<span data-ttu-id="9efd0-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9efd0-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9efd0-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9efd0-203">Пример</span><span class="sxs-lookup"><span data-stu-id="9efd0-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="9efd0-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="9efd0-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="9efd0-205">Получает приложение Office, в котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="9efd0-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="9efd0-206">Тип</span><span class="sxs-lookup"><span data-stu-id="9efd0-206">Type</span></span>

*   [<span data-ttu-id="9efd0-207">HostType</span><span class="sxs-lookup"><span data-stu-id="9efd0-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="9efd0-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="9efd0-208">Requirements</span></span>

|<span data-ttu-id="9efd0-209">Требование</span><span class="sxs-lookup"><span data-stu-id="9efd0-209">Requirement</span></span>| <span data-ttu-id="9efd0-210">Значение</span><span class="sxs-lookup"><span data-stu-id="9efd0-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="9efd0-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9efd0-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9efd0-212">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-212">1.1</span></span>|
|[<span data-ttu-id="9efd0-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9efd0-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9efd0-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9efd0-215">Пример</span><span class="sxs-lookup"><span data-stu-id="9efd0-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="9efd0-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="9efd0-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="9efd0-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="9efd0-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9efd0-218">Тип</span><span class="sxs-lookup"><span data-stu-id="9efd0-218">Type</span></span>

*   [<span data-ttu-id="9efd0-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9efd0-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="9efd0-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="9efd0-220">Requirements</span></span>

|<span data-ttu-id="9efd0-221">Требование</span><span class="sxs-lookup"><span data-stu-id="9efd0-221">Requirement</span></span>| <span data-ttu-id="9efd0-222">Значение</span><span class="sxs-lookup"><span data-stu-id="9efd0-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="9efd0-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9efd0-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9efd0-224">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-224">1.1</span></span>|
|[<span data-ttu-id="9efd0-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9efd0-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9efd0-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9efd0-227">Пример</span><span class="sxs-lookup"><span data-stu-id="9efd0-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="9efd0-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="9efd0-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="9efd0-229">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="9efd0-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9efd0-230">Тип</span><span class="sxs-lookup"><span data-stu-id="9efd0-230">Type</span></span>

*   [<span data-ttu-id="9efd0-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="9efd0-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="9efd0-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="9efd0-232">Requirements</span></span>

|<span data-ttu-id="9efd0-233">Требование</span><span class="sxs-lookup"><span data-stu-id="9efd0-233">Requirement</span></span>| <span data-ttu-id="9efd0-234">Значение</span><span class="sxs-lookup"><span data-stu-id="9efd0-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="9efd0-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9efd0-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9efd0-236">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-236">1.1</span></span>|
|[<span data-ttu-id="9efd0-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9efd0-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9efd0-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9efd0-239">Пример</span><span class="sxs-lookup"><span data-stu-id="9efd0-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="9efd0-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="9efd0-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="9efd0-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="9efd0-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="9efd0-242">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="9efd0-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="9efd0-243">Тип</span><span class="sxs-lookup"><span data-stu-id="9efd0-243">Type</span></span>

*   [<span data-ttu-id="9efd0-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9efd0-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="9efd0-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="9efd0-245">Requirements</span></span>

|<span data-ttu-id="9efd0-246">Требование</span><span class="sxs-lookup"><span data-stu-id="9efd0-246">Requirement</span></span>| <span data-ttu-id="9efd0-247">Значение</span><span class="sxs-lookup"><span data-stu-id="9efd0-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="9efd0-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9efd0-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9efd0-249">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-249">1.1</span></span>|
|[<span data-ttu-id="9efd0-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9efd0-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="9efd0-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9efd0-251">Restricted</span></span>|
|[<span data-ttu-id="9efd0-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9efd0-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9efd0-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="9efd0-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="9efd0-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="9efd0-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="9efd0-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="9efd0-256">Тип</span><span class="sxs-lookup"><span data-stu-id="9efd0-256">Type</span></span>

*   [<span data-ttu-id="9efd0-257">UI</span><span class="sxs-lookup"><span data-stu-id="9efd0-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="9efd0-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="9efd0-258">Requirements</span></span>

|<span data-ttu-id="9efd0-259">Требование</span><span class="sxs-lookup"><span data-stu-id="9efd0-259">Requirement</span></span>| <span data-ttu-id="9efd0-260">Значение</span><span class="sxs-lookup"><span data-stu-id="9efd0-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="9efd0-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9efd0-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9efd0-262">1.1</span><span class="sxs-lookup"><span data-stu-id="9efd0-262">1.1</span></span>|
|[<span data-ttu-id="9efd0-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9efd0-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9efd0-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9efd0-264">Compose or Read</span></span>|
