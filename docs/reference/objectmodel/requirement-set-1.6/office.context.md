---
title: Office. Context — набор обязательных элементов 1,6
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,6.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 449ea52b623e3ec491a426096becae7a1b8a334a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293711"
---
# <a name="context-mailbox-requirement-set-16"></a><span data-ttu-id="1fadb-103">контекст (набор требований для почтового ящика 1,6)</span><span class="sxs-lookup"><span data-stu-id="1fadb-103">context (Mailbox requirement set 1.6)</span></span>

### <a name="officecontext"></a><span data-ttu-id="1fadb-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="1fadb-104">[Office](office.md).context</span></span>

<span data-ttu-id="1fadb-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="1fadb-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="1fadb-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.6).</span><span class="sxs-lookup"><span data-stu-id="1fadb-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.6).</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadb-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadb-107">Requirements</span></span>

|<span data-ttu-id="1fadb-108">Требование</span><span class="sxs-lookup"><span data-stu-id="1fadb-108">Requirement</span></span>| <span data-ttu-id="1fadb-109">Значение</span><span class="sxs-lookup"><span data-stu-id="1fadb-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadb-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1fadb-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1fadb-111">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-111">1.1</span></span>|
|[<span data-ttu-id="1fadb-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1fadb-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1fadb-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="1fadb-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="1fadb-114">Properties</span></span>

| <span data-ttu-id="1fadb-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="1fadb-115">Property</span></span> | <span data-ttu-id="1fadb-116">Способов</span><span class="sxs-lookup"><span data-stu-id="1fadb-116">Modes</span></span> | <span data-ttu-id="1fadb-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="1fadb-117">Return type</span></span> | <span data-ttu-id="1fadb-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="1fadb-118">Minimum</span></span><br><span data-ttu-id="1fadb-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="1fadb-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="1fadb-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="1fadb-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="1fadb-121">Создание</span><span class="sxs-lookup"><span data-stu-id="1fadb-121">Compose</span></span><br><span data-ttu-id="1fadb-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-122">Read</span></span> | <span data-ttu-id="1fadb-123">String</span><span class="sxs-lookup"><span data-stu-id="1fadb-123">String</span></span> | [<span data-ttu-id="1fadb-124">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1fadb-125">диагностики</span><span class="sxs-lookup"><span data-stu-id="1fadb-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="1fadb-126">Создание</span><span class="sxs-lookup"><span data-stu-id="1fadb-126">Compose</span></span><br><span data-ttu-id="1fadb-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-127">Read</span></span> | [<span data-ttu-id="1fadb-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="1fadb-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.6) | [<span data-ttu-id="1fadb-129">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1fadb-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="1fadb-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="1fadb-131">Создание</span><span class="sxs-lookup"><span data-stu-id="1fadb-131">Compose</span></span><br><span data-ttu-id="1fadb-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-132">Read</span></span> | <span data-ttu-id="1fadb-133">String</span><span class="sxs-lookup"><span data-stu-id="1fadb-133">String</span></span> | [<span data-ttu-id="1fadb-134">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1fadb-135">размещать</span><span class="sxs-lookup"><span data-stu-id="1fadb-135">host</span></span>](#host-hosttype) | <span data-ttu-id="1fadb-136">Создание</span><span class="sxs-lookup"><span data-stu-id="1fadb-136">Compose</span></span><br><span data-ttu-id="1fadb-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-137">Read</span></span> | [<span data-ttu-id="1fadb-138">HostType</span><span class="sxs-lookup"><span data-stu-id="1fadb-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.6) | [<span data-ttu-id="1fadb-139">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1fadb-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="1fadb-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="1fadb-141">Создание</span><span class="sxs-lookup"><span data-stu-id="1fadb-141">Compose</span></span><br><span data-ttu-id="1fadb-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-142">Read</span></span> | [<span data-ttu-id="1fadb-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="1fadb-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.6) | [<span data-ttu-id="1fadb-144">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1fadb-145">платформа</span><span class="sxs-lookup"><span data-stu-id="1fadb-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="1fadb-146">Создание</span><span class="sxs-lookup"><span data-stu-id="1fadb-146">Compose</span></span><br><span data-ttu-id="1fadb-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-147">Read</span></span> | [<span data-ttu-id="1fadb-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1fadb-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.6) | [<span data-ttu-id="1fadb-149">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1fadb-150">потребность</span><span class="sxs-lookup"><span data-stu-id="1fadb-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="1fadb-151">Создание</span><span class="sxs-lookup"><span data-stu-id="1fadb-151">Compose</span></span><br><span data-ttu-id="1fadb-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-152">Read</span></span> | [<span data-ttu-id="1fadb-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="1fadb-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.6) | [<span data-ttu-id="1fadb-154">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1fadb-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="1fadb-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="1fadb-156">Создание</span><span class="sxs-lookup"><span data-stu-id="1fadb-156">Compose</span></span><br><span data-ttu-id="1fadb-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-157">Read</span></span> | [<span data-ttu-id="1fadb-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1fadb-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.6) | [<span data-ttu-id="1fadb-159">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="1fadb-160">ui</span><span class="sxs-lookup"><span data-stu-id="1fadb-160">ui</span></span>](#ui-ui) | <span data-ttu-id="1fadb-161">Создание</span><span class="sxs-lookup"><span data-stu-id="1fadb-161">Compose</span></span><br><span data-ttu-id="1fadb-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-162">Read</span></span> | [<span data-ttu-id="1fadb-163">UI</span><span class="sxs-lookup"><span data-stu-id="1fadb-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.6) | [<span data-ttu-id="1fadb-164">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="1fadb-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="1fadb-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="1fadb-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="1fadb-166">contentLanguage: String</span></span>

<span data-ttu-id="1fadb-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="1fadb-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="1fadb-168">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="1fadb-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadb-169">Тип</span><span class="sxs-lookup"><span data-stu-id="1fadb-169">Type</span></span>

*   <span data-ttu-id="1fadb-170">String</span><span class="sxs-lookup"><span data-stu-id="1fadb-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadb-171">Требования</span><span class="sxs-lookup"><span data-stu-id="1fadb-171">Requirements</span></span>

|<span data-ttu-id="1fadb-172">Требование</span><span class="sxs-lookup"><span data-stu-id="1fadb-172">Requirement</span></span>| <span data-ttu-id="1fadb-173">Значение</span><span class="sxs-lookup"><span data-stu-id="1fadb-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadb-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1fadb-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1fadb-175">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-175">1.1</span></span>|
|[<span data-ttu-id="1fadb-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1fadb-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1fadb-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadb-178">Пример</span><span class="sxs-lookup"><span data-stu-id="1fadb-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="1fadb-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="1fadb-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="1fadb-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="1fadb-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadb-181">Тип</span><span class="sxs-lookup"><span data-stu-id="1fadb-181">Type</span></span>

*   [<span data-ttu-id="1fadb-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="1fadb-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="1fadb-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadb-183">Requirements</span></span>

|<span data-ttu-id="1fadb-184">Требование</span><span class="sxs-lookup"><span data-stu-id="1fadb-184">Requirement</span></span>| <span data-ttu-id="1fadb-185">Значение</span><span class="sxs-lookup"><span data-stu-id="1fadb-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadb-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1fadb-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1fadb-187">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-187">1.1</span></span>|
|[<span data-ttu-id="1fadb-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1fadb-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1fadb-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadb-190">Пример</span><span class="sxs-lookup"><span data-stu-id="1fadb-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="1fadb-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="1fadb-191">displayLanguage: String</span></span>

<span data-ttu-id="1fadb-192">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="1fadb-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="1fadb-193">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="1fadb-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadb-194">Тип</span><span class="sxs-lookup"><span data-stu-id="1fadb-194">Type</span></span>

*   <span data-ttu-id="1fadb-195">String</span><span class="sxs-lookup"><span data-stu-id="1fadb-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="1fadb-196">Требования</span><span class="sxs-lookup"><span data-stu-id="1fadb-196">Requirements</span></span>

|<span data-ttu-id="1fadb-197">Требование</span><span class="sxs-lookup"><span data-stu-id="1fadb-197">Requirement</span></span>| <span data-ttu-id="1fadb-198">Значение</span><span class="sxs-lookup"><span data-stu-id="1fadb-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadb-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1fadb-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1fadb-200">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-200">1.1</span></span>|
|[<span data-ttu-id="1fadb-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1fadb-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1fadb-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadb-203">Пример</span><span class="sxs-lookup"><span data-stu-id="1fadb-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="1fadb-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="1fadb-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="1fadb-205">Получает приложение Office, в котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="1fadb-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadb-206">Тип</span><span class="sxs-lookup"><span data-stu-id="1fadb-206">Type</span></span>

*   [<span data-ttu-id="1fadb-207">HostType</span><span class="sxs-lookup"><span data-stu-id="1fadb-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="1fadb-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadb-208">Requirements</span></span>

|<span data-ttu-id="1fadb-209">Требование</span><span class="sxs-lookup"><span data-stu-id="1fadb-209">Requirement</span></span>| <span data-ttu-id="1fadb-210">Значение</span><span class="sxs-lookup"><span data-stu-id="1fadb-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadb-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1fadb-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1fadb-212">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-212">1.1</span></span>|
|[<span data-ttu-id="1fadb-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1fadb-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1fadb-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadb-215">Пример</span><span class="sxs-lookup"><span data-stu-id="1fadb-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="1fadb-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="1fadb-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="1fadb-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="1fadb-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadb-218">Тип</span><span class="sxs-lookup"><span data-stu-id="1fadb-218">Type</span></span>

*   [<span data-ttu-id="1fadb-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="1fadb-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="1fadb-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadb-220">Requirements</span></span>

|<span data-ttu-id="1fadb-221">Требование</span><span class="sxs-lookup"><span data-stu-id="1fadb-221">Requirement</span></span>| <span data-ttu-id="1fadb-222">Значение</span><span class="sxs-lookup"><span data-stu-id="1fadb-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadb-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1fadb-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1fadb-224">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-224">1.1</span></span>|
|[<span data-ttu-id="1fadb-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1fadb-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1fadb-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadb-227">Пример</span><span class="sxs-lookup"><span data-stu-id="1fadb-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="1fadb-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="1fadb-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="1fadb-229">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="1fadb-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadb-230">Тип</span><span class="sxs-lookup"><span data-stu-id="1fadb-230">Type</span></span>

*   [<span data-ttu-id="1fadb-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="1fadb-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="1fadb-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadb-232">Requirements</span></span>

|<span data-ttu-id="1fadb-233">Требование</span><span class="sxs-lookup"><span data-stu-id="1fadb-233">Requirement</span></span>| <span data-ttu-id="1fadb-234">Значение</span><span class="sxs-lookup"><span data-stu-id="1fadb-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadb-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1fadb-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1fadb-236">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-236">1.1</span></span>|
|[<span data-ttu-id="1fadb-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1fadb-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1fadb-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="1fadb-239">Пример</span><span class="sxs-lookup"><span data-stu-id="1fadb-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="1fadb-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="1fadb-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="1fadb-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="1fadb-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="1fadb-242">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="1fadb-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadb-243">Тип</span><span class="sxs-lookup"><span data-stu-id="1fadb-243">Type</span></span>

*   [<span data-ttu-id="1fadb-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="1fadb-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="1fadb-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadb-245">Requirements</span></span>

|<span data-ttu-id="1fadb-246">Требование</span><span class="sxs-lookup"><span data-stu-id="1fadb-246">Requirement</span></span>| <span data-ttu-id="1fadb-247">Значение</span><span class="sxs-lookup"><span data-stu-id="1fadb-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadb-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1fadb-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1fadb-249">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-249">1.1</span></span>|
|[<span data-ttu-id="1fadb-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="1fadb-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="1fadb-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="1fadb-251">Restricted</span></span>|
|[<span data-ttu-id="1fadb-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1fadb-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1fadb-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="1fadb-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="1fadb-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="1fadb-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="1fadb-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="1fadb-256">Тип</span><span class="sxs-lookup"><span data-stu-id="1fadb-256">Type</span></span>

*   [<span data-ttu-id="1fadb-257">UI</span><span class="sxs-lookup"><span data-stu-id="1fadb-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="1fadb-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="1fadb-258">Requirements</span></span>

|<span data-ttu-id="1fadb-259">Требование</span><span class="sxs-lookup"><span data-stu-id="1fadb-259">Requirement</span></span>| <span data-ttu-id="1fadb-260">Значение</span><span class="sxs-lookup"><span data-stu-id="1fadb-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="1fadb-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="1fadb-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="1fadb-262">1.1</span><span class="sxs-lookup"><span data-stu-id="1fadb-262">1.1</span></span>|
|[<span data-ttu-id="1fadb-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="1fadb-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="1fadb-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="1fadb-264">Compose or Read</span></span>|
