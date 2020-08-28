---
title: Office. Context — набор обязательных элементов 1,5
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,5.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 4587609e08d48e5c9e0cc74ec2c9050af0c3877a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293753"
---
# <a name="context-mailbox-requirement-set-15"></a><span data-ttu-id="8b13d-103">контекст (набор требований для почтового ящика 1,5)</span><span class="sxs-lookup"><span data-stu-id="8b13d-103">context (Mailbox requirement set 1.5)</span></span>

### <a name="officecontext"></a><span data-ttu-id="8b13d-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="8b13d-104">[Office](office.md).context</span></span>

<span data-ttu-id="8b13d-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="8b13d-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="8b13d-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.5).</span><span class="sxs-lookup"><span data-stu-id="8b13d-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5).</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b13d-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b13d-107">Requirements</span></span>

|<span data-ttu-id="8b13d-108">Требование</span><span class="sxs-lookup"><span data-stu-id="8b13d-108">Requirement</span></span>| <span data-ttu-id="8b13d-109">Значение</span><span class="sxs-lookup"><span data-stu-id="8b13d-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b13d-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8b13d-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b13d-111">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-111">1.1</span></span>|
|[<span data-ttu-id="8b13d-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b13d-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b13d-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="8b13d-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="8b13d-114">Properties</span></span>

| <span data-ttu-id="8b13d-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="8b13d-115">Property</span></span> | <span data-ttu-id="8b13d-116">Способов</span><span class="sxs-lookup"><span data-stu-id="8b13d-116">Modes</span></span> | <span data-ttu-id="8b13d-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="8b13d-117">Return type</span></span> | <span data-ttu-id="8b13d-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="8b13d-118">Minimum</span></span><br><span data-ttu-id="8b13d-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="8b13d-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="8b13d-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="8b13d-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="8b13d-121">Создание</span><span class="sxs-lookup"><span data-stu-id="8b13d-121">Compose</span></span><br><span data-ttu-id="8b13d-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-122">Read</span></span> | <span data-ttu-id="8b13d-123">String</span><span class="sxs-lookup"><span data-stu-id="8b13d-123">String</span></span> | [<span data-ttu-id="8b13d-124">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8b13d-125">диагностики</span><span class="sxs-lookup"><span data-stu-id="8b13d-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="8b13d-126">Создание</span><span class="sxs-lookup"><span data-stu-id="8b13d-126">Compose</span></span><br><span data-ttu-id="8b13d-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-127">Read</span></span> | [<span data-ttu-id="8b13d-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="8b13d-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5) | [<span data-ttu-id="8b13d-129">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8b13d-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="8b13d-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="8b13d-131">Создание</span><span class="sxs-lookup"><span data-stu-id="8b13d-131">Compose</span></span><br><span data-ttu-id="8b13d-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-132">Read</span></span> | <span data-ttu-id="8b13d-133">String</span><span class="sxs-lookup"><span data-stu-id="8b13d-133">String</span></span> | [<span data-ttu-id="8b13d-134">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8b13d-135">размещать</span><span class="sxs-lookup"><span data-stu-id="8b13d-135">host</span></span>](#host-hosttype) | <span data-ttu-id="8b13d-136">Создание</span><span class="sxs-lookup"><span data-stu-id="8b13d-136">Compose</span></span><br><span data-ttu-id="8b13d-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-137">Read</span></span> | [<span data-ttu-id="8b13d-138">HostType</span><span class="sxs-lookup"><span data-stu-id="8b13d-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5) | [<span data-ttu-id="8b13d-139">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8b13d-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="8b13d-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="8b13d-141">Создание</span><span class="sxs-lookup"><span data-stu-id="8b13d-141">Compose</span></span><br><span data-ttu-id="8b13d-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-142">Read</span></span> | [<span data-ttu-id="8b13d-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="8b13d-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5) | [<span data-ttu-id="8b13d-144">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8b13d-145">платформа</span><span class="sxs-lookup"><span data-stu-id="8b13d-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="8b13d-146">Создание</span><span class="sxs-lookup"><span data-stu-id="8b13d-146">Compose</span></span><br><span data-ttu-id="8b13d-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-147">Read</span></span> | [<span data-ttu-id="8b13d-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="8b13d-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5) | [<span data-ttu-id="8b13d-149">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8b13d-150">потребность</span><span class="sxs-lookup"><span data-stu-id="8b13d-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="8b13d-151">Создание</span><span class="sxs-lookup"><span data-stu-id="8b13d-151">Compose</span></span><br><span data-ttu-id="8b13d-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-152">Read</span></span> | [<span data-ttu-id="8b13d-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="8b13d-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5) | [<span data-ttu-id="8b13d-154">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8b13d-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="8b13d-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="8b13d-156">Создание</span><span class="sxs-lookup"><span data-stu-id="8b13d-156">Compose</span></span><br><span data-ttu-id="8b13d-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-157">Read</span></span> | [<span data-ttu-id="8b13d-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8b13d-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5) | [<span data-ttu-id="8b13d-159">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="8b13d-160">ui</span><span class="sxs-lookup"><span data-stu-id="8b13d-160">ui</span></span>](#ui-ui) | <span data-ttu-id="8b13d-161">Создание</span><span class="sxs-lookup"><span data-stu-id="8b13d-161">Compose</span></span><br><span data-ttu-id="8b13d-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-162">Read</span></span> | [<span data-ttu-id="8b13d-163">UI</span><span class="sxs-lookup"><span data-stu-id="8b13d-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5) | [<span data-ttu-id="8b13d-164">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="8b13d-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="8b13d-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="8b13d-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="8b13d-166">contentLanguage: String</span></span>

<span data-ttu-id="8b13d-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="8b13d-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="8b13d-168">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="8b13d-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="8b13d-169">Тип</span><span class="sxs-lookup"><span data-stu-id="8b13d-169">Type</span></span>

*   <span data-ttu-id="8b13d-170">String</span><span class="sxs-lookup"><span data-stu-id="8b13d-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b13d-171">Требования</span><span class="sxs-lookup"><span data-stu-id="8b13d-171">Requirements</span></span>

|<span data-ttu-id="8b13d-172">Требование</span><span class="sxs-lookup"><span data-stu-id="8b13d-172">Requirement</span></span>| <span data-ttu-id="8b13d-173">Значение</span><span class="sxs-lookup"><span data-stu-id="8b13d-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b13d-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8b13d-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b13d-175">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-175">1.1</span></span>|
|[<span data-ttu-id="8b13d-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b13d-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b13d-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b13d-178">Пример</span><span class="sxs-lookup"><span data-stu-id="8b13d-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="8b13d-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="8b13d-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="8b13d-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="8b13d-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="8b13d-181">Тип</span><span class="sxs-lookup"><span data-stu-id="8b13d-181">Type</span></span>

*   [<span data-ttu-id="8b13d-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="8b13d-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="8b13d-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b13d-183">Requirements</span></span>

|<span data-ttu-id="8b13d-184">Требование</span><span class="sxs-lookup"><span data-stu-id="8b13d-184">Requirement</span></span>| <span data-ttu-id="8b13d-185">Значение</span><span class="sxs-lookup"><span data-stu-id="8b13d-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b13d-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8b13d-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b13d-187">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-187">1.1</span></span>|
|[<span data-ttu-id="8b13d-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b13d-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b13d-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b13d-190">Пример</span><span class="sxs-lookup"><span data-stu-id="8b13d-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="8b13d-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="8b13d-191">displayLanguage: String</span></span>

<span data-ttu-id="8b13d-192">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="8b13d-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="8b13d-193">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="8b13d-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="8b13d-194">Тип</span><span class="sxs-lookup"><span data-stu-id="8b13d-194">Type</span></span>

*   <span data-ttu-id="8b13d-195">String</span><span class="sxs-lookup"><span data-stu-id="8b13d-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8b13d-196">Требования</span><span class="sxs-lookup"><span data-stu-id="8b13d-196">Requirements</span></span>

|<span data-ttu-id="8b13d-197">Требование</span><span class="sxs-lookup"><span data-stu-id="8b13d-197">Requirement</span></span>| <span data-ttu-id="8b13d-198">Значение</span><span class="sxs-lookup"><span data-stu-id="8b13d-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b13d-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8b13d-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b13d-200">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-200">1.1</span></span>|
|[<span data-ttu-id="8b13d-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b13d-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b13d-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b13d-203">Пример</span><span class="sxs-lookup"><span data-stu-id="8b13d-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="8b13d-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="8b13d-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="8b13d-205">Получает приложение Office, в котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="8b13d-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="8b13d-206">Тип</span><span class="sxs-lookup"><span data-stu-id="8b13d-206">Type</span></span>

*   [<span data-ttu-id="8b13d-207">HostType</span><span class="sxs-lookup"><span data-stu-id="8b13d-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="8b13d-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b13d-208">Requirements</span></span>

|<span data-ttu-id="8b13d-209">Требование</span><span class="sxs-lookup"><span data-stu-id="8b13d-209">Requirement</span></span>| <span data-ttu-id="8b13d-210">Значение</span><span class="sxs-lookup"><span data-stu-id="8b13d-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b13d-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8b13d-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b13d-212">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-212">1.1</span></span>|
|[<span data-ttu-id="8b13d-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b13d-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b13d-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b13d-215">Пример</span><span class="sxs-lookup"><span data-stu-id="8b13d-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="8b13d-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="8b13d-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="8b13d-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="8b13d-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="8b13d-218">Тип</span><span class="sxs-lookup"><span data-stu-id="8b13d-218">Type</span></span>

*   [<span data-ttu-id="8b13d-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="8b13d-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="8b13d-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b13d-220">Requirements</span></span>

|<span data-ttu-id="8b13d-221">Требование</span><span class="sxs-lookup"><span data-stu-id="8b13d-221">Requirement</span></span>| <span data-ttu-id="8b13d-222">Значение</span><span class="sxs-lookup"><span data-stu-id="8b13d-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b13d-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8b13d-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b13d-224">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-224">1.1</span></span>|
|[<span data-ttu-id="8b13d-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b13d-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b13d-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b13d-227">Пример</span><span class="sxs-lookup"><span data-stu-id="8b13d-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="8b13d-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="8b13d-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="8b13d-229">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="8b13d-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="8b13d-230">Тип</span><span class="sxs-lookup"><span data-stu-id="8b13d-230">Type</span></span>

*   [<span data-ttu-id="8b13d-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="8b13d-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="8b13d-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b13d-232">Requirements</span></span>

|<span data-ttu-id="8b13d-233">Требование</span><span class="sxs-lookup"><span data-stu-id="8b13d-233">Requirement</span></span>| <span data-ttu-id="8b13d-234">Значение</span><span class="sxs-lookup"><span data-stu-id="8b13d-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b13d-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8b13d-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b13d-236">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-236">1.1</span></span>|
|[<span data-ttu-id="8b13d-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b13d-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b13d-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8b13d-239">Пример</span><span class="sxs-lookup"><span data-stu-id="8b13d-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="8b13d-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="8b13d-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="8b13d-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="8b13d-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="8b13d-242">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="8b13d-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="8b13d-243">Тип</span><span class="sxs-lookup"><span data-stu-id="8b13d-243">Type</span></span>

*   [<span data-ttu-id="8b13d-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="8b13d-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="8b13d-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b13d-245">Requirements</span></span>

|<span data-ttu-id="8b13d-246">Требование</span><span class="sxs-lookup"><span data-stu-id="8b13d-246">Requirement</span></span>| <span data-ttu-id="8b13d-247">Значение</span><span class="sxs-lookup"><span data-stu-id="8b13d-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b13d-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8b13d-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b13d-249">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-249">1.1</span></span>|
|[<span data-ttu-id="8b13d-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8b13d-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="8b13d-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="8b13d-251">Restricted</span></span>|
|[<span data-ttu-id="8b13d-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b13d-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b13d-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="8b13d-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="8b13d-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="8b13d-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="8b13d-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="8b13d-256">Тип</span><span class="sxs-lookup"><span data-stu-id="8b13d-256">Type</span></span>

*   [<span data-ttu-id="8b13d-257">UI</span><span class="sxs-lookup"><span data-stu-id="8b13d-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="8b13d-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="8b13d-258">Requirements</span></span>

|<span data-ttu-id="8b13d-259">Требование</span><span class="sxs-lookup"><span data-stu-id="8b13d-259">Requirement</span></span>| <span data-ttu-id="8b13d-260">Значение</span><span class="sxs-lookup"><span data-stu-id="8b13d-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="8b13d-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8b13d-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="8b13d-262">1.1</span><span class="sxs-lookup"><span data-stu-id="8b13d-262">1.1</span></span>|
|[<span data-ttu-id="8b13d-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8b13d-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="8b13d-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8b13d-264">Compose or Read</span></span>|
