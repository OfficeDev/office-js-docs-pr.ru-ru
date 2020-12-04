---
title: Office. Context — набор обязательных элементов 1,4
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,4.
ms.date: 12/02/2020
localization_priority: Normal
ms.openlocfilehash: 0ec84c9d0695871fa3be265c37ce1e682cdfb6af
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570774"
---
# <a name="context-mailbox-requirement-set-14"></a><span data-ttu-id="c68fe-103">контекст (набор требований для почтового ящика 1,4)</span><span class="sxs-lookup"><span data-stu-id="c68fe-103">context (Mailbox requirement set 1.4)</span></span>

### <a name="officecontext"></a><span data-ttu-id="c68fe-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="c68fe-104">[Office](office.md).context</span></span>

<span data-ttu-id="c68fe-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="c68fe-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="c68fe-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="c68fe-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.4&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="c68fe-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="c68fe-107">Requirements</span></span>

|<span data-ttu-id="c68fe-108">Требование</span><span class="sxs-lookup"><span data-stu-id="c68fe-108">Requirement</span></span>| <span data-ttu-id="c68fe-109">Значение</span><span class="sxs-lookup"><span data-stu-id="c68fe-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="c68fe-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c68fe-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c68fe-111">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-111">1.1</span></span>|
|[<span data-ttu-id="c68fe-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c68fe-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c68fe-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="c68fe-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="c68fe-114">Properties</span></span>

| <span data-ttu-id="c68fe-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="c68fe-115">Property</span></span> | <span data-ttu-id="c68fe-116">Способов</span><span class="sxs-lookup"><span data-stu-id="c68fe-116">Modes</span></span> | <span data-ttu-id="c68fe-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="c68fe-117">Return type</span></span> | <span data-ttu-id="c68fe-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="c68fe-118">Minimum</span></span><br><span data-ttu-id="c68fe-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="c68fe-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="c68fe-120">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="c68fe-120">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="c68fe-121">Создание</span><span class="sxs-lookup"><span data-stu-id="c68fe-121">Compose</span></span><br><span data-ttu-id="c68fe-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-122">Read</span></span> | <span data-ttu-id="c68fe-123">String</span><span class="sxs-lookup"><span data-stu-id="c68fe-123">String</span></span> | [<span data-ttu-id="c68fe-124">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-124">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c68fe-125">диагностики</span><span class="sxs-lookup"><span data-stu-id="c68fe-125">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="c68fe-126">Создание</span><span class="sxs-lookup"><span data-stu-id="c68fe-126">Compose</span></span><br><span data-ttu-id="c68fe-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-127">Read</span></span> | [<span data-ttu-id="c68fe-128">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="c68fe-128">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="c68fe-129">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c68fe-130">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="c68fe-130">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="c68fe-131">Создание</span><span class="sxs-lookup"><span data-stu-id="c68fe-131">Compose</span></span><br><span data-ttu-id="c68fe-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-132">Read</span></span> | <span data-ttu-id="c68fe-133">String</span><span class="sxs-lookup"><span data-stu-id="c68fe-133">String</span></span> | [<span data-ttu-id="c68fe-134">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c68fe-135">mailbox</span><span class="sxs-lookup"><span data-stu-id="c68fe-135">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="c68fe-136">Создание</span><span class="sxs-lookup"><span data-stu-id="c68fe-136">Compose</span></span><br><span data-ttu-id="c68fe-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-137">Read</span></span> | [<span data-ttu-id="c68fe-138">Mailbox</span><span class="sxs-lookup"><span data-stu-id="c68fe-138">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="c68fe-139">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c68fe-140">потребность</span><span class="sxs-lookup"><span data-stu-id="c68fe-140">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="c68fe-141">Создание</span><span class="sxs-lookup"><span data-stu-id="c68fe-141">Compose</span></span><br><span data-ttu-id="c68fe-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-142">Read</span></span> | [<span data-ttu-id="c68fe-143">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="c68fe-143">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="c68fe-144">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c68fe-145">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="c68fe-145">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="c68fe-146">Создание</span><span class="sxs-lookup"><span data-stu-id="c68fe-146">Compose</span></span><br><span data-ttu-id="c68fe-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-147">Read</span></span> | [<span data-ttu-id="c68fe-148">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c68fe-148">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="c68fe-149">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="c68fe-150">ui</span><span class="sxs-lookup"><span data-stu-id="c68fe-150">ui</span></span>](#ui-ui) | <span data-ttu-id="c68fe-151">Создание</span><span class="sxs-lookup"><span data-stu-id="c68fe-151">Compose</span></span><br><span data-ttu-id="c68fe-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-152">Read</span></span> | [<span data-ttu-id="c68fe-153">UI</span><span class="sxs-lookup"><span data-stu-id="c68fe-153">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.4&preserve-view=true) | [<span data-ttu-id="c68fe-154">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="c68fe-155">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="c68fe-155">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="c68fe-156">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="c68fe-156">contentLanguage: String</span></span>

<span data-ttu-id="c68fe-157">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="c68fe-157">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="c68fe-158">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="c68fe-158">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="c68fe-159">Тип</span><span class="sxs-lookup"><span data-stu-id="c68fe-159">Type</span></span>

*   <span data-ttu-id="c68fe-160">String</span><span class="sxs-lookup"><span data-stu-id="c68fe-160">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c68fe-161">Требования</span><span class="sxs-lookup"><span data-stu-id="c68fe-161">Requirements</span></span>

|<span data-ttu-id="c68fe-162">Требование</span><span class="sxs-lookup"><span data-stu-id="c68fe-162">Requirement</span></span>| <span data-ttu-id="c68fe-163">Значение</span><span class="sxs-lookup"><span data-stu-id="c68fe-163">Value</span></span>|
|---|---|
|[<span data-ttu-id="c68fe-164">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c68fe-164">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c68fe-165">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-165">1.1</span></span>|
|[<span data-ttu-id="c68fe-166">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c68fe-166">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c68fe-167">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-167">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c68fe-168">Пример</span><span class="sxs-lookup"><span data-stu-id="c68fe-168">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="c68fe-169">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="c68fe-169">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="c68fe-170">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="c68fe-170">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="c68fe-171">Type</span><span class="sxs-lookup"><span data-stu-id="c68fe-171">Type</span></span>

*   [<span data-ttu-id="c68fe-172">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="c68fe-172">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="c68fe-173">Requirements</span><span class="sxs-lookup"><span data-stu-id="c68fe-173">Requirements</span></span>

|<span data-ttu-id="c68fe-174">Требование</span><span class="sxs-lookup"><span data-stu-id="c68fe-174">Requirement</span></span>| <span data-ttu-id="c68fe-175">Значение</span><span class="sxs-lookup"><span data-stu-id="c68fe-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="c68fe-176">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c68fe-176">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c68fe-177">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-177">1.1</span></span>|
|[<span data-ttu-id="c68fe-178">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c68fe-178">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c68fe-179">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c68fe-180">Пример</span><span class="sxs-lookup"><span data-stu-id="c68fe-180">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="c68fe-181">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="c68fe-181">displayLanguage: String</span></span>

<span data-ttu-id="c68fe-182">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="c68fe-182">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="c68fe-183">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="c68fe-183">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="c68fe-184">Тип</span><span class="sxs-lookup"><span data-stu-id="c68fe-184">Type</span></span>

*   <span data-ttu-id="c68fe-185">String</span><span class="sxs-lookup"><span data-stu-id="c68fe-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c68fe-186">Требования</span><span class="sxs-lookup"><span data-stu-id="c68fe-186">Requirements</span></span>

|<span data-ttu-id="c68fe-187">Требование</span><span class="sxs-lookup"><span data-stu-id="c68fe-187">Requirement</span></span>| <span data-ttu-id="c68fe-188">Значение</span><span class="sxs-lookup"><span data-stu-id="c68fe-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="c68fe-189">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c68fe-189">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c68fe-190">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-190">1.1</span></span>|
|[<span data-ttu-id="c68fe-191">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c68fe-191">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c68fe-192">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-192">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c68fe-193">Пример</span><span class="sxs-lookup"><span data-stu-id="c68fe-193">Example</span></span>

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

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="c68fe-194">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="c68fe-194">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="c68fe-195">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="c68fe-195">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="c68fe-196">Type</span><span class="sxs-lookup"><span data-stu-id="c68fe-196">Type</span></span>

*   [<span data-ttu-id="c68fe-197">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="c68fe-197">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="c68fe-198">Requirements</span><span class="sxs-lookup"><span data-stu-id="c68fe-198">Requirements</span></span>

|<span data-ttu-id="c68fe-199">Требование</span><span class="sxs-lookup"><span data-stu-id="c68fe-199">Requirement</span></span>| <span data-ttu-id="c68fe-200">Значение</span><span class="sxs-lookup"><span data-stu-id="c68fe-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="c68fe-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c68fe-201">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c68fe-202">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-202">1.1</span></span>|
|[<span data-ttu-id="c68fe-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c68fe-203">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c68fe-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="c68fe-205">Пример</span><span class="sxs-lookup"><span data-stu-id="c68fe-205">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="c68fe-206">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="c68fe-206">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="c68fe-207">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="c68fe-207">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="c68fe-208">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="c68fe-208">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="c68fe-209">Type</span><span class="sxs-lookup"><span data-stu-id="c68fe-209">Type</span></span>

*   [<span data-ttu-id="c68fe-210">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="c68fe-210">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="c68fe-211">Requirements</span><span class="sxs-lookup"><span data-stu-id="c68fe-211">Requirements</span></span>

|<span data-ttu-id="c68fe-212">Требование</span><span class="sxs-lookup"><span data-stu-id="c68fe-212">Requirement</span></span>| <span data-ttu-id="c68fe-213">Значение</span><span class="sxs-lookup"><span data-stu-id="c68fe-213">Value</span></span>|
|---|---|
|[<span data-ttu-id="c68fe-214">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c68fe-214">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c68fe-215">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-215">1.1</span></span>|
|[<span data-ttu-id="c68fe-216">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c68fe-216">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="c68fe-217">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="c68fe-217">Restricted</span></span>|
|[<span data-ttu-id="c68fe-218">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c68fe-218">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c68fe-219">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-219">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="c68fe-220">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="c68fe-220">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="c68fe-221">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="c68fe-221">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="c68fe-222">Type</span><span class="sxs-lookup"><span data-stu-id="c68fe-222">Type</span></span>

*   [<span data-ttu-id="c68fe-223">UI</span><span class="sxs-lookup"><span data-stu-id="c68fe-223">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="c68fe-224">Requirements</span><span class="sxs-lookup"><span data-stu-id="c68fe-224">Requirements</span></span>

|<span data-ttu-id="c68fe-225">Требование</span><span class="sxs-lookup"><span data-stu-id="c68fe-225">Requirement</span></span>| <span data-ttu-id="c68fe-226">Значение</span><span class="sxs-lookup"><span data-stu-id="c68fe-226">Value</span></span>|
|---|---|
|[<span data-ttu-id="c68fe-227">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c68fe-227">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="c68fe-228">1.1</span><span class="sxs-lookup"><span data-stu-id="c68fe-228">1.1</span></span>|
|[<span data-ttu-id="c68fe-229">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c68fe-229">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="c68fe-230">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c68fe-230">Compose or Read</span></span>|
