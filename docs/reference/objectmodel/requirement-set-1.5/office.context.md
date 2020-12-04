---
title: Office. Context — набор обязательных элементов 1,5
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,5.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 966c2065268d973ac8476fda839d2a6cdf038f4e
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570739"
---
# <a name="context-mailbox-requirement-set-15"></a><span data-ttu-id="a1430-103">контекст (набор требований для почтового ящика 1,5)</span><span class="sxs-lookup"><span data-stu-id="a1430-103">context (Mailbox requirement set 1.5)</span></span>

### <a name="officecontext"></a><span data-ttu-id="a1430-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="a1430-104">[Office](office.md).context</span></span>

<span data-ttu-id="a1430-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="a1430-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="a1430-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="a1430-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.5&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1430-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="a1430-107">Requirements</span></span>

|<span data-ttu-id="a1430-108">Требование</span><span class="sxs-lookup"><span data-stu-id="a1430-108">Requirement</span></span>| <span data-ttu-id="a1430-109">Значение</span><span class="sxs-lookup"><span data-stu-id="a1430-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1430-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a1430-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1430-111">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-111">1.1</span></span>|
|[<span data-ttu-id="a1430-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a1430-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1430-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="a1430-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="a1430-114">Properties</span></span>

| <span data-ttu-id="a1430-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="a1430-115">Property</span></span> | <span data-ttu-id="a1430-116">Способов</span><span class="sxs-lookup"><span data-stu-id="a1430-116">Modes</span></span> | <span data-ttu-id="a1430-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="a1430-117">Return type</span></span> | <span data-ttu-id="a1430-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="a1430-118">Minimum</span></span><br><span data-ttu-id="a1430-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="a1430-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="a1430-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="a1430-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="a1430-121">Создание</span><span class="sxs-lookup"><span data-stu-id="a1430-121">Compose</span></span><br><span data-ttu-id="a1430-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-122">Read</span></span> | <span data-ttu-id="a1430-123">String</span><span class="sxs-lookup"><span data-stu-id="a1430-123">String</span></span> | [<span data-ttu-id="a1430-124">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a1430-125">диагностики</span><span class="sxs-lookup"><span data-stu-id="a1430-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="a1430-126">Создание</span><span class="sxs-lookup"><span data-stu-id="a1430-126">Compose</span></span><br><span data-ttu-id="a1430-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-127">Read</span></span> | [<span data-ttu-id="a1430-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="a1430-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="a1430-129">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a1430-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="a1430-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="a1430-131">Создание</span><span class="sxs-lookup"><span data-stu-id="a1430-131">Compose</span></span><br><span data-ttu-id="a1430-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-132">Read</span></span> | <span data-ttu-id="a1430-133">String</span><span class="sxs-lookup"><span data-stu-id="a1430-133">String</span></span> | [<span data-ttu-id="a1430-134">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a1430-135">размещать</span><span class="sxs-lookup"><span data-stu-id="a1430-135">host</span></span>](#host-hosttype) | <span data-ttu-id="a1430-136">Создание</span><span class="sxs-lookup"><span data-stu-id="a1430-136">Compose</span></span><br><span data-ttu-id="a1430-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-137">Read</span></span> | [<span data-ttu-id="a1430-138">HostType</span><span class="sxs-lookup"><span data-stu-id="a1430-138">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="a1430-139">1,5</span><span class="sxs-lookup"><span data-stu-id="a1430-139">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="a1430-140">mailbox</span><span class="sxs-lookup"><span data-stu-id="a1430-140">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="a1430-141">Создание</span><span class="sxs-lookup"><span data-stu-id="a1430-141">Compose</span></span><br><span data-ttu-id="a1430-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-142">Read</span></span> | [<span data-ttu-id="a1430-143">Mailbox</span><span class="sxs-lookup"><span data-stu-id="a1430-143">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="a1430-144">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a1430-145">платформа</span><span class="sxs-lookup"><span data-stu-id="a1430-145">platform</span></span>](#platform-platformtype) | <span data-ttu-id="a1430-146">Создание</span><span class="sxs-lookup"><span data-stu-id="a1430-146">Compose</span></span><br><span data-ttu-id="a1430-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-147">Read</span></span> | [<span data-ttu-id="a1430-148">PlatformType</span><span class="sxs-lookup"><span data-stu-id="a1430-148">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="a1430-149">1,5</span><span class="sxs-lookup"><span data-stu-id="a1430-149">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="a1430-150">потребность</span><span class="sxs-lookup"><span data-stu-id="a1430-150">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="a1430-151">Создание</span><span class="sxs-lookup"><span data-stu-id="a1430-151">Compose</span></span><br><span data-ttu-id="a1430-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-152">Read</span></span> | [<span data-ttu-id="a1430-153">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="a1430-153">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="a1430-154">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a1430-155">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="a1430-155">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="a1430-156">Создание</span><span class="sxs-lookup"><span data-stu-id="a1430-156">Compose</span></span><br><span data-ttu-id="a1430-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-157">Read</span></span> | [<span data-ttu-id="a1430-158">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a1430-158">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="a1430-159">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="a1430-160">ui</span><span class="sxs-lookup"><span data-stu-id="a1430-160">ui</span></span>](#ui-ui) | <span data-ttu-id="a1430-161">Создание</span><span class="sxs-lookup"><span data-stu-id="a1430-161">Compose</span></span><br><span data-ttu-id="a1430-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-162">Read</span></span> | [<span data-ttu-id="a1430-163">UI</span><span class="sxs-lookup"><span data-stu-id="a1430-163">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.5&preserve-view=true) | [<span data-ttu-id="a1430-164">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="a1430-165">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="a1430-165">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="a1430-166">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="a1430-166">contentLanguage: String</span></span>

<span data-ttu-id="a1430-167">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="a1430-167">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="a1430-168">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="a1430-168">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="a1430-169">Тип</span><span class="sxs-lookup"><span data-stu-id="a1430-169">Type</span></span>

*   <span data-ttu-id="a1430-170">String</span><span class="sxs-lookup"><span data-stu-id="a1430-170">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1430-171">Требования</span><span class="sxs-lookup"><span data-stu-id="a1430-171">Requirements</span></span>

|<span data-ttu-id="a1430-172">Требование</span><span class="sxs-lookup"><span data-stu-id="a1430-172">Requirement</span></span>| <span data-ttu-id="a1430-173">Значение</span><span class="sxs-lookup"><span data-stu-id="a1430-173">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1430-174">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a1430-174">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1430-175">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-175">1.1</span></span>|
|[<span data-ttu-id="a1430-176">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a1430-176">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1430-177">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-177">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1430-178">Пример</span><span class="sxs-lookup"><span data-stu-id="a1430-178">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="a1430-179">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="a1430-179">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="a1430-180">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="a1430-180">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="a1430-181">Type</span><span class="sxs-lookup"><span data-stu-id="a1430-181">Type</span></span>

*   [<span data-ttu-id="a1430-182">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="a1430-182">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="a1430-183">Requirements</span><span class="sxs-lookup"><span data-stu-id="a1430-183">Requirements</span></span>

|<span data-ttu-id="a1430-184">Требование</span><span class="sxs-lookup"><span data-stu-id="a1430-184">Requirement</span></span>| <span data-ttu-id="a1430-185">Значение</span><span class="sxs-lookup"><span data-stu-id="a1430-185">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1430-186">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a1430-186">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1430-187">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-187">1.1</span></span>|
|[<span data-ttu-id="a1430-188">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a1430-188">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1430-189">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-189">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1430-190">Пример</span><span class="sxs-lookup"><span data-stu-id="a1430-190">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="a1430-191">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="a1430-191">displayLanguage: String</span></span>

<span data-ttu-id="a1430-192">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="a1430-192">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="a1430-193">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="a1430-193">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="a1430-194">Тип</span><span class="sxs-lookup"><span data-stu-id="a1430-194">Type</span></span>

*   <span data-ttu-id="a1430-195">String</span><span class="sxs-lookup"><span data-stu-id="a1430-195">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a1430-196">Требования</span><span class="sxs-lookup"><span data-stu-id="a1430-196">Requirements</span></span>

|<span data-ttu-id="a1430-197">Требование</span><span class="sxs-lookup"><span data-stu-id="a1430-197">Requirement</span></span>| <span data-ttu-id="a1430-198">Значение</span><span class="sxs-lookup"><span data-stu-id="a1430-198">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1430-199">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a1430-199">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1430-200">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-200">1.1</span></span>|
|[<span data-ttu-id="a1430-201">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a1430-201">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1430-202">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-202">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1430-203">Пример</span><span class="sxs-lookup"><span data-stu-id="a1430-203">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="a1430-204">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="a1430-204">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="a1430-205">Получает приложение Office, в котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="a1430-205">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="a1430-206">Кроме того, вы можете использовать свойство [Office. Context. Diagnostics](#diagnostics-contextinformation) для получения ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="a1430-206">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="a1430-207">Type</span><span class="sxs-lookup"><span data-stu-id="a1430-207">Type</span></span>

*   [<span data-ttu-id="a1430-208">HostType</span><span class="sxs-lookup"><span data-stu-id="a1430-208">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="a1430-209">Requirements</span><span class="sxs-lookup"><span data-stu-id="a1430-209">Requirements</span></span>

|<span data-ttu-id="a1430-210">Требование</span><span class="sxs-lookup"><span data-stu-id="a1430-210">Requirement</span></span>| <span data-ttu-id="a1430-211">Значение</span><span class="sxs-lookup"><span data-stu-id="a1430-211">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1430-212">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a1430-212">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1430-213">1.5</span><span class="sxs-lookup"><span data-stu-id="a1430-213">1.5</span></span>|
|[<span data-ttu-id="a1430-214">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a1430-214">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1430-215">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-215">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1430-216">Пример</span><span class="sxs-lookup"><span data-stu-id="a1430-216">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="a1430-217">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="a1430-217">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="a1430-218">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="a1430-218">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="a1430-219">Кроме того, вы можете использовать свойство [Office. Context. Diagnostics](#diagnostics-contextinformation) для получения платформы.</span><span class="sxs-lookup"><span data-stu-id="a1430-219">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="a1430-220">Type</span><span class="sxs-lookup"><span data-stu-id="a1430-220">Type</span></span>

*   [<span data-ttu-id="a1430-221">PlatformType</span><span class="sxs-lookup"><span data-stu-id="a1430-221">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="a1430-222">Requirements</span><span class="sxs-lookup"><span data-stu-id="a1430-222">Requirements</span></span>

|<span data-ttu-id="a1430-223">Требование</span><span class="sxs-lookup"><span data-stu-id="a1430-223">Requirement</span></span>| <span data-ttu-id="a1430-224">Значение</span><span class="sxs-lookup"><span data-stu-id="a1430-224">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1430-225">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a1430-225">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1430-226">1.5</span><span class="sxs-lookup"><span data-stu-id="a1430-226">1.5</span></span>|
|[<span data-ttu-id="a1430-227">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a1430-227">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1430-228">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-228">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1430-229">Пример</span><span class="sxs-lookup"><span data-stu-id="a1430-229">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="a1430-230">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="a1430-230">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="a1430-231">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="a1430-231">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="a1430-232">Type</span><span class="sxs-lookup"><span data-stu-id="a1430-232">Type</span></span>

*   [<span data-ttu-id="a1430-233">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="a1430-233">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="a1430-234">Requirements</span><span class="sxs-lookup"><span data-stu-id="a1430-234">Requirements</span></span>

|<span data-ttu-id="a1430-235">Требование</span><span class="sxs-lookup"><span data-stu-id="a1430-235">Requirement</span></span>| <span data-ttu-id="a1430-236">Значение</span><span class="sxs-lookup"><span data-stu-id="a1430-236">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1430-237">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a1430-237">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1430-238">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-238">1.1</span></span>|
|[<span data-ttu-id="a1430-239">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a1430-239">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1430-240">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-240">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="a1430-241">Пример</span><span class="sxs-lookup"><span data-stu-id="a1430-241">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="a1430-242">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="a1430-242">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="a1430-243">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="a1430-243">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="a1430-244">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="a1430-244">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="a1430-245">Type</span><span class="sxs-lookup"><span data-stu-id="a1430-245">Type</span></span>

*   [<span data-ttu-id="a1430-246">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="a1430-246">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="a1430-247">Requirements</span><span class="sxs-lookup"><span data-stu-id="a1430-247">Requirements</span></span>

|<span data-ttu-id="a1430-248">Требование</span><span class="sxs-lookup"><span data-stu-id="a1430-248">Requirement</span></span>| <span data-ttu-id="a1430-249">Значение</span><span class="sxs-lookup"><span data-stu-id="a1430-249">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1430-250">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a1430-250">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1430-251">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-251">1.1</span></span>|
|[<span data-ttu-id="a1430-252">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a1430-252">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="a1430-253">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="a1430-253">Restricted</span></span>|
|[<span data-ttu-id="a1430-254">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a1430-254">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1430-255">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-255">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="a1430-256">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="a1430-256">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="a1430-257">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="a1430-257">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="a1430-258">Type</span><span class="sxs-lookup"><span data-stu-id="a1430-258">Type</span></span>

*   [<span data-ttu-id="a1430-259">UI</span><span class="sxs-lookup"><span data-stu-id="a1430-259">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="a1430-260">Requirements</span><span class="sxs-lookup"><span data-stu-id="a1430-260">Requirements</span></span>

|<span data-ttu-id="a1430-261">Требование</span><span class="sxs-lookup"><span data-stu-id="a1430-261">Requirement</span></span>| <span data-ttu-id="a1430-262">Значение</span><span class="sxs-lookup"><span data-stu-id="a1430-262">Value</span></span>|
|---|---|
|[<span data-ttu-id="a1430-263">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a1430-263">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="a1430-264">1.1</span><span class="sxs-lookup"><span data-stu-id="a1430-264">1.1</span></span>|
|[<span data-ttu-id="a1430-265">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a1430-265">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="a1430-266">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a1430-266">Compose or Read</span></span>|
