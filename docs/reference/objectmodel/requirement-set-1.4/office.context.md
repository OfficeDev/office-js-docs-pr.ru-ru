---
title: Office. Context — набор обязательных элементов 1,4
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,4.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 139ca19d059b303976117cd259fa90c58e5eeb13
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47293739"
---
# <a name="context-mailbox-requirement-set-14"></a><span data-ttu-id="95c35-103">контекст (набор требований для почтового ящика 1,4)</span><span class="sxs-lookup"><span data-stu-id="95c35-103">context (Mailbox requirement set 1.4)</span></span>

### <a name="officecontext"></a><span data-ttu-id="95c35-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="95c35-104">[Office](office.md).context</span></span>

<span data-ttu-id="95c35-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="95c35-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="95c35-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.4).</span><span class="sxs-lookup"><span data-stu-id="95c35-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4).</span></span>

##### <a name="requirements"></a><span data-ttu-id="95c35-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="95c35-107">Requirements</span></span>

|<span data-ttu-id="95c35-108">Требование</span><span class="sxs-lookup"><span data-stu-id="95c35-108">Requirement</span></span>| <span data-ttu-id="95c35-109">Значение</span><span class="sxs-lookup"><span data-stu-id="95c35-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="95c35-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="95c35-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95c35-111">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-111">1.1</span></span>|
|[<span data-ttu-id="95c35-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="95c35-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95c35-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="95c35-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="95c35-114">Properties</span></span>

| <span data-ttu-id="95c35-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="95c35-115">Property</span></span> | <span data-ttu-id="95c35-116">Способов</span><span class="sxs-lookup"><span data-stu-id="95c35-116">Modes</span></span> | <span data-ttu-id="95c35-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="95c35-117">Return type</span></span> | <span data-ttu-id="95c35-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="95c35-118">Minimum</span></span><br><span data-ttu-id="95c35-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="95c35-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="95c35-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="95c35-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="95c35-121">Создание</span><span class="sxs-lookup"><span data-stu-id="95c35-121">Compose</span></span><br><span data-ttu-id="95c35-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-122">Read</span></span> | <span data-ttu-id="95c35-123">String</span><span class="sxs-lookup"><span data-stu-id="95c35-123">String</span></span> | [<span data-ttu-id="95c35-124">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95c35-125">диагностики</span><span class="sxs-lookup"><span data-stu-id="95c35-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="95c35-126">Создание</span><span class="sxs-lookup"><span data-stu-id="95c35-126">Compose</span></span><br><span data-ttu-id="95c35-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-127">Read</span></span> | [<span data-ttu-id="95c35-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="95c35-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4) | [<span data-ttu-id="95c35-129">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95c35-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="95c35-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="95c35-131">Создание</span><span class="sxs-lookup"><span data-stu-id="95c35-131">Compose</span></span><br><span data-ttu-id="95c35-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-132">Read</span></span> | <span data-ttu-id="95c35-133">String</span><span class="sxs-lookup"><span data-stu-id="95c35-133">String</span></span> | [<span data-ttu-id="95c35-134">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95c35-135">размещать</span><span class="sxs-lookup"><span data-stu-id="95c35-135">host</span></span>](#host-hosttype) | <span data-ttu-id="95c35-136">Создание</span><span class="sxs-lookup"><span data-stu-id="95c35-136">Compose</span></span><br><span data-ttu-id="95c35-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-137">Read</span></span> | [<span data-ttu-id="95c35-138">HostType</span><span class="sxs-lookup"><span data-stu-id="95c35-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.4) | [<span data-ttu-id="95c35-139">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95c35-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="95c35-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="95c35-141">Создание</span><span class="sxs-lookup"><span data-stu-id="95c35-141">Compose</span></span><br><span data-ttu-id="95c35-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-142">Read</span></span> | [<span data-ttu-id="95c35-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="95c35-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4) | [<span data-ttu-id="95c35-144">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95c35-145">платформа</span><span class="sxs-lookup"><span data-stu-id="95c35-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="95c35-146">Создание</span><span class="sxs-lookup"><span data-stu-id="95c35-146">Compose</span></span><br><span data-ttu-id="95c35-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-147">Read</span></span> | [<span data-ttu-id="95c35-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="95c35-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.4) | [<span data-ttu-id="95c35-149">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95c35-150">потребность</span><span class="sxs-lookup"><span data-stu-id="95c35-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="95c35-151">Создание</span><span class="sxs-lookup"><span data-stu-id="95c35-151">Compose</span></span><br><span data-ttu-id="95c35-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-152">Read</span></span> | [<span data-ttu-id="95c35-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="95c35-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4) | [<span data-ttu-id="95c35-154">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95c35-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="95c35-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="95c35-156">Создание</span><span class="sxs-lookup"><span data-stu-id="95c35-156">Compose</span></span><br><span data-ttu-id="95c35-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-157">Read</span></span> | [<span data-ttu-id="95c35-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="95c35-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4) | [<span data-ttu-id="95c35-159">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="95c35-160">ui</span><span class="sxs-lookup"><span data-stu-id="95c35-160">ui</span></span>](#ui-ui) | <span data-ttu-id="95c35-161">Создание</span><span class="sxs-lookup"><span data-stu-id="95c35-161">Compose</span></span><br><span data-ttu-id="95c35-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-162">Read</span></span> | [<span data-ttu-id="95c35-163">UI</span><span class="sxs-lookup"><span data-stu-id="95c35-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4) | [<span data-ttu-id="95c35-164">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="95c35-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="95c35-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="95c35-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="95c35-166">contentLanguage: String</span></span>

<span data-ttu-id="95c35-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="95c35-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="95c35-168">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="95c35-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="95c35-169">Тип</span><span class="sxs-lookup"><span data-stu-id="95c35-169">Type</span></span>

*   <span data-ttu-id="95c35-170">String</span><span class="sxs-lookup"><span data-stu-id="95c35-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="95c35-171">Требования</span><span class="sxs-lookup"><span data-stu-id="95c35-171">Requirements</span></span>

|<span data-ttu-id="95c35-172">Требование</span><span class="sxs-lookup"><span data-stu-id="95c35-172">Requirement</span></span>| <span data-ttu-id="95c35-173">Значение</span><span class="sxs-lookup"><span data-stu-id="95c35-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="95c35-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="95c35-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95c35-175">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-175">1.1</span></span>|
|[<span data-ttu-id="95c35-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="95c35-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95c35-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95c35-178">Пример</span><span class="sxs-lookup"><span data-stu-id="95c35-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="95c35-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="95c35-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="95c35-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="95c35-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="95c35-181">Тип</span><span class="sxs-lookup"><span data-stu-id="95c35-181">Type</span></span>

*   [<span data-ttu-id="95c35-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="95c35-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="95c35-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="95c35-183">Requirements</span></span>

|<span data-ttu-id="95c35-184">Требование</span><span class="sxs-lookup"><span data-stu-id="95c35-184">Requirement</span></span>| <span data-ttu-id="95c35-185">Значение</span><span class="sxs-lookup"><span data-stu-id="95c35-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="95c35-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="95c35-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95c35-187">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-187">1.1</span></span>|
|[<span data-ttu-id="95c35-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="95c35-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95c35-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95c35-190">Пример</span><span class="sxs-lookup"><span data-stu-id="95c35-190">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="95c35-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="95c35-191">displayLanguage: String</span></span>

<span data-ttu-id="95c35-192">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="95c35-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="95c35-193">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="95c35-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="95c35-194">Тип</span><span class="sxs-lookup"><span data-stu-id="95c35-194">Type</span></span>

*   <span data-ttu-id="95c35-195">String</span><span class="sxs-lookup"><span data-stu-id="95c35-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="95c35-196">Требования</span><span class="sxs-lookup"><span data-stu-id="95c35-196">Requirements</span></span>

|<span data-ttu-id="95c35-197">Требование</span><span class="sxs-lookup"><span data-stu-id="95c35-197">Requirement</span></span>| <span data-ttu-id="95c35-198">Значение</span><span class="sxs-lookup"><span data-stu-id="95c35-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="95c35-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="95c35-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95c35-200">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-200">1.1</span></span>|
|[<span data-ttu-id="95c35-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="95c35-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95c35-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95c35-203">Пример</span><span class="sxs-lookup"><span data-stu-id="95c35-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="95c35-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="95c35-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="95c35-205">Получает приложение Office, в котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="95c35-205">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="95c35-206">Тип</span><span class="sxs-lookup"><span data-stu-id="95c35-206">Type</span></span>

*   [<span data-ttu-id="95c35-207">HostType</span><span class="sxs-lookup"><span data-stu-id="95c35-207">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="95c35-208">Requirements</span><span class="sxs-lookup"><span data-stu-id="95c35-208">Requirements</span></span>

|<span data-ttu-id="95c35-209">Требование</span><span class="sxs-lookup"><span data-stu-id="95c35-209">Requirement</span></span>| <span data-ttu-id="95c35-210">Значение</span><span class="sxs-lookup"><span data-stu-id="95c35-210">Value</span></span>|
|---|---|
|[<span data-ttu-id="95c35-211">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="95c35-211">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95c35-212">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-212">1.1</span></span>|
|[<span data-ttu-id="95c35-213">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="95c35-213">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95c35-214">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-214">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95c35-215">Пример</span><span class="sxs-lookup"><span data-stu-id="95c35-215">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="95c35-216">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="95c35-216">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="95c35-217">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="95c35-217">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="95c35-218">Тип</span><span class="sxs-lookup"><span data-stu-id="95c35-218">Type</span></span>

*   [<span data-ttu-id="95c35-219">PlatformType</span><span class="sxs-lookup"><span data-stu-id="95c35-219">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="95c35-220">Requirements</span><span class="sxs-lookup"><span data-stu-id="95c35-220">Requirements</span></span>

|<span data-ttu-id="95c35-221">Требование</span><span class="sxs-lookup"><span data-stu-id="95c35-221">Requirement</span></span>| <span data-ttu-id="95c35-222">Значение</span><span class="sxs-lookup"><span data-stu-id="95c35-222">Value</span></span>|
|---|---|
|[<span data-ttu-id="95c35-223">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="95c35-223">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95c35-224">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-224">1.1</span></span>|
|[<span data-ttu-id="95c35-225">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="95c35-225">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95c35-226">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-226">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95c35-227">Пример</span><span class="sxs-lookup"><span data-stu-id="95c35-227">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="95c35-228">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="95c35-228">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="95c35-229">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="95c35-229">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="95c35-230">Тип</span><span class="sxs-lookup"><span data-stu-id="95c35-230">Type</span></span>

*   [<span data-ttu-id="95c35-231">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="95c35-231">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="95c35-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="95c35-232">Requirements</span></span>

|<span data-ttu-id="95c35-233">Требование</span><span class="sxs-lookup"><span data-stu-id="95c35-233">Requirement</span></span>| <span data-ttu-id="95c35-234">Значение</span><span class="sxs-lookup"><span data-stu-id="95c35-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="95c35-235">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="95c35-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95c35-236">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-236">1.1</span></span>|
|[<span data-ttu-id="95c35-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="95c35-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95c35-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="95c35-239">Пример</span><span class="sxs-lookup"><span data-stu-id="95c35-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="95c35-240">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="95c35-240">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="95c35-241">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="95c35-241">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="95c35-242">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="95c35-242">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="95c35-243">Тип</span><span class="sxs-lookup"><span data-stu-id="95c35-243">Type</span></span>

*   [<span data-ttu-id="95c35-244">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="95c35-244">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="95c35-245">Requirements</span><span class="sxs-lookup"><span data-stu-id="95c35-245">Requirements</span></span>

|<span data-ttu-id="95c35-246">Требование</span><span class="sxs-lookup"><span data-stu-id="95c35-246">Requirement</span></span>| <span data-ttu-id="95c35-247">Значение</span><span class="sxs-lookup"><span data-stu-id="95c35-247">Value</span></span>|
|---|---|
|[<span data-ttu-id="95c35-248">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="95c35-248">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95c35-249">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-249">1.1</span></span>|
|[<span data-ttu-id="95c35-250">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="95c35-250">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="95c35-251">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="95c35-251">Restricted</span></span>|
|[<span data-ttu-id="95c35-252">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="95c35-252">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95c35-253">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-253">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="95c35-254">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="95c35-254">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="95c35-255">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="95c35-255">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="95c35-256">Тип</span><span class="sxs-lookup"><span data-stu-id="95c35-256">Type</span></span>

*   [<span data-ttu-id="95c35-257">UI</span><span class="sxs-lookup"><span data-stu-id="95c35-257">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="95c35-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="95c35-258">Requirements</span></span>

|<span data-ttu-id="95c35-259">Требование</span><span class="sxs-lookup"><span data-stu-id="95c35-259">Requirement</span></span>| <span data-ttu-id="95c35-260">Значение</span><span class="sxs-lookup"><span data-stu-id="95c35-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="95c35-261">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="95c35-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="95c35-262">1.1</span><span class="sxs-lookup"><span data-stu-id="95c35-262">1.1</span></span>|
|[<span data-ttu-id="95c35-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="95c35-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="95c35-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="95c35-264">Compose or Read</span></span>|
