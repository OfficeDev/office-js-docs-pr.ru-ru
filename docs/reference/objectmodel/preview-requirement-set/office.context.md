---
title: Office.context — предварительная версия набора обязательных элементов
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 5c34a7a0db5880a94ba5519059a93010a5243978
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629190"
---
# <a name="context"></a><span data-ttu-id="0529e-102">context</span><span class="sxs-lookup"><span data-stu-id="0529e-102">context</span></span>

### <a name="officeofficemdcontext"></a><span data-ttu-id="0529e-103">[Office](Office.md).context</span><span class="sxs-lookup"><span data-stu-id="0529e-103">[Office](Office.md).context</span></span>

<span data-ttu-id="0529e-p101">Пространство имен Office.context содержит общие интерфейсы, которые используются надстройками всех приложений Office. В этот список входят только интерфейсы, используемые надстройками Outlook. Полный список интерфейсов пространства имен Office.context см. в статье [Ссылка на пространство имен Office.context в общем API](/javascript/api/office/office.context).</span><span class="sxs-lookup"><span data-stu-id="0529e-p101">The Office.context namespace provides shared interfaces that are used by add-ins in all of the Office apps. This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0529e-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="0529e-106">Requirements</span></span>

|<span data-ttu-id="0529e-107">Требование</span><span class="sxs-lookup"><span data-stu-id="0529e-107">Requirement</span></span>| <span data-ttu-id="0529e-108">Значение</span><span class="sxs-lookup"><span data-stu-id="0529e-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="0529e-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0529e-109">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0529e-110">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-110">1.0</span></span>|
|[<span data-ttu-id="0529e-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0529e-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0529e-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="0529e-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="0529e-113">Properties</span></span>

| <span data-ttu-id="0529e-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="0529e-114">Property</span></span> | <span data-ttu-id="0529e-115">Способов</span><span class="sxs-lookup"><span data-stu-id="0529e-115">Modes</span></span> | <span data-ttu-id="0529e-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="0529e-116">Return type</span></span> | <span data-ttu-id="0529e-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="0529e-117">Minimum</span></span><br><span data-ttu-id="0529e-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="0529e-118">requirement set</span></span> |
|---|---|---|---|
| [<span data-ttu-id="0529e-119">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="0529e-119">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="0529e-120">Создание</span><span class="sxs-lookup"><span data-stu-id="0529e-120">Compose</span></span><br><span data-ttu-id="0529e-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-121">Read</span></span> | <span data-ttu-id="0529e-122">String</span><span class="sxs-lookup"><span data-stu-id="0529e-122">String</span></span> | <span data-ttu-id="0529e-123">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-123">1.0</span></span> |
| [<span data-ttu-id="0529e-124">diagnostics</span><span class="sxs-lookup"><span data-stu-id="0529e-124">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="0529e-125">Создание</span><span class="sxs-lookup"><span data-stu-id="0529e-125">Compose</span></span><br><span data-ttu-id="0529e-126">Чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-126">Read</span></span> | [<span data-ttu-id="0529e-127">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="0529e-127">ContextInformation</span></span>](/javascript/api/office/office.contextinformation) | <span data-ttu-id="0529e-128">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-128">1.0</span></span> |
| [<span data-ttu-id="0529e-129">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="0529e-129">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="0529e-130">Создание</span><span class="sxs-lookup"><span data-stu-id="0529e-130">Compose</span></span><br><span data-ttu-id="0529e-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-131">Read</span></span> | <span data-ttu-id="0529e-132">String</span><span class="sxs-lookup"><span data-stu-id="0529e-132">String</span></span> | <span data-ttu-id="0529e-133">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-133">1.0</span></span> |
| [<span data-ttu-id="0529e-134">размещать</span><span class="sxs-lookup"><span data-stu-id="0529e-134">host</span></span>](#host-hosttype) | <span data-ttu-id="0529e-135">Создание</span><span class="sxs-lookup"><span data-stu-id="0529e-135">Compose</span></span><br><span data-ttu-id="0529e-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-136">Read</span></span> | [<span data-ttu-id="0529e-137">HostType</span><span class="sxs-lookup"><span data-stu-id="0529e-137">HostType</span></span>](/javascript/api/office/office.hosttype) | <span data-ttu-id="0529e-138">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-138">1.0</span></span> |
| [<span data-ttu-id="0529e-139">officeTheme</span><span class="sxs-lookup"><span data-stu-id="0529e-139">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="0529e-140">Создание</span><span class="sxs-lookup"><span data-stu-id="0529e-140">Compose</span></span><br><span data-ttu-id="0529e-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-141">Read</span></span> | [<span data-ttu-id="0529e-142">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="0529e-142">OfficeTheme</span></span>](/javascript/api/office/office.officetheme) | <span data-ttu-id="0529e-143">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0529e-143">Preview</span></span> |
| [<span data-ttu-id="0529e-144">управляем</span><span class="sxs-lookup"><span data-stu-id="0529e-144">platform</span></span>](#platform-platformtype) | <span data-ttu-id="0529e-145">Создание</span><span class="sxs-lookup"><span data-stu-id="0529e-145">Compose</span></span><br><span data-ttu-id="0529e-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-146">Read</span></span> | [<span data-ttu-id="0529e-147">PlatformType</span><span class="sxs-lookup"><span data-stu-id="0529e-147">PlatformType</span></span>](/javascript/api/office/office.platformtype) | <span data-ttu-id="0529e-148">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-148">1.0</span></span> |
| [<span data-ttu-id="0529e-149">потребность</span><span class="sxs-lookup"><span data-stu-id="0529e-149">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="0529e-150">Создание</span><span class="sxs-lookup"><span data-stu-id="0529e-150">Compose</span></span><br><span data-ttu-id="0529e-151">Чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-151">Read</span></span> | [<span data-ttu-id="0529e-152">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="0529e-152">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport) | <span data-ttu-id="0529e-153">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-153">1.0</span></span> |
| [<span data-ttu-id="0529e-154">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="0529e-154">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="0529e-155">Создание</span><span class="sxs-lookup"><span data-stu-id="0529e-155">Compose</span></span><br><span data-ttu-id="0529e-156">Чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-156">Read</span></span> | [<span data-ttu-id="0529e-157">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="0529e-157">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings) | <span data-ttu-id="0529e-158">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-158">1.0</span></span> |
| [<span data-ttu-id="0529e-159">ui</span><span class="sxs-lookup"><span data-stu-id="0529e-159">ui</span></span>](#ui-ui) | <span data-ttu-id="0529e-160">Создание</span><span class="sxs-lookup"><span data-stu-id="0529e-160">Compose</span></span><br><span data-ttu-id="0529e-161">Чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-161">Read</span></span> | [<span data-ttu-id="0529e-162">UI</span><span class="sxs-lookup"><span data-stu-id="0529e-162">UI</span></span>](/javascript/api/office/office.ui) | <span data-ttu-id="0529e-163">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-163">1.0</span></span> |

### <a name="namespaces"></a><span data-ttu-id="0529e-164">Пространства имен</span><span class="sxs-lookup"><span data-stu-id="0529e-164">Namespaces</span></span>

<span data-ttu-id="0529e-165">[Проверка подлинности](/javascript/api/office/office.auth): предоставляет поддержку [единого входа (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).</span><span class="sxs-lookup"><span data-stu-id="0529e-165">[auth](/javascript/api/office/office.auth): Provides support for [single sign-on (SSO)](/outlook/add-ins/authenticate-a-user-with-an-sso-token).</span></span>

<span data-ttu-id="0529e-166">[почтовый ящик](office.context.mailbox.md): предоставляет доступ к объектной модели надстройки Outlook для Microsoft Outlook.</span><span class="sxs-lookup"><span data-stu-id="0529e-166">[mailbox](office.context.mailbox.md): Provides access to the Outlook add-in object model for Microsoft Outlook.</span></span>

## <a name="property-details"></a><span data-ttu-id="0529e-167">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="0529e-167">Property details</span></span>

#### <a name="contentlanguage-string"></a><span data-ttu-id="0529e-168">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="0529e-168">contentLanguage: String</span></span>

<span data-ttu-id="0529e-169">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="0529e-169">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="0529e-170">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="0529e-170">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="0529e-171">Тип</span><span class="sxs-lookup"><span data-stu-id="0529e-171">Type</span></span>

*   <span data-ttu-id="0529e-172">String</span><span class="sxs-lookup"><span data-stu-id="0529e-172">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0529e-173">Требования</span><span class="sxs-lookup"><span data-stu-id="0529e-173">Requirements</span></span>

|<span data-ttu-id="0529e-174">Требование</span><span class="sxs-lookup"><span data-stu-id="0529e-174">Requirement</span></span>| <span data-ttu-id="0529e-175">Значение</span><span class="sxs-lookup"><span data-stu-id="0529e-175">Value</span></span>|
|---|---|
|[<span data-ttu-id="0529e-176">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0529e-176">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0529e-177">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-177">1.0</span></span>|
|[<span data-ttu-id="0529e-178">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0529e-178">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0529e-179">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-179">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0529e-180">Пример</span><span class="sxs-lookup"><span data-stu-id="0529e-180">Example</span></span>

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

#### <a name="diagnostics-contextinformationjavascriptapiofficeofficecontextinformation"></a><span data-ttu-id="0529e-181">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="0529e-181">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="0529e-182">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="0529e-182">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="0529e-183">Тип</span><span class="sxs-lookup"><span data-stu-id="0529e-183">Type</span></span>

*   [<span data-ttu-id="0529e-184">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="0529e-184">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="0529e-185">Requirements</span><span class="sxs-lookup"><span data-stu-id="0529e-185">Requirements</span></span>

|<span data-ttu-id="0529e-186">Требование</span><span class="sxs-lookup"><span data-stu-id="0529e-186">Requirement</span></span>| <span data-ttu-id="0529e-187">Значение</span><span class="sxs-lookup"><span data-stu-id="0529e-187">Value</span></span>|
|---|---|
|[<span data-ttu-id="0529e-188">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0529e-188">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0529e-189">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-189">1.0</span></span>|
|[<span data-ttu-id="0529e-190">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0529e-190">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0529e-191">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-191">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0529e-192">Пример</span><span class="sxs-lookup"><span data-stu-id="0529e-192">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="0529e-193">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="0529e-193">displayLanguage: String</span></span>

<span data-ttu-id="0529e-194">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="0529e-194">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="0529e-195">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="0529e-195">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="0529e-196">Тип</span><span class="sxs-lookup"><span data-stu-id="0529e-196">Type</span></span>

*   <span data-ttu-id="0529e-197">String</span><span class="sxs-lookup"><span data-stu-id="0529e-197">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0529e-198">Требования</span><span class="sxs-lookup"><span data-stu-id="0529e-198">Requirements</span></span>

|<span data-ttu-id="0529e-199">Требование</span><span class="sxs-lookup"><span data-stu-id="0529e-199">Requirement</span></span>| <span data-ttu-id="0529e-200">Значение</span><span class="sxs-lookup"><span data-stu-id="0529e-200">Value</span></span>|
|---|---|
|[<span data-ttu-id="0529e-201">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0529e-201">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0529e-202">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-202">1.0</span></span>|
|[<span data-ttu-id="0529e-203">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0529e-203">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0529e-204">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-204">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0529e-205">Пример</span><span class="sxs-lookup"><span data-stu-id="0529e-205">Example</span></span>

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

#### <a name="host-hosttypejavascriptapiofficeofficehosttype"></a><span data-ttu-id="0529e-206">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="0529e-206">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="0529e-207">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="0529e-207">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="0529e-208">Тип</span><span class="sxs-lookup"><span data-stu-id="0529e-208">Type</span></span>

*   [<span data-ttu-id="0529e-209">HostType</span><span class="sxs-lookup"><span data-stu-id="0529e-209">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="0529e-210">Requirements</span><span class="sxs-lookup"><span data-stu-id="0529e-210">Requirements</span></span>

|<span data-ttu-id="0529e-211">Требование</span><span class="sxs-lookup"><span data-stu-id="0529e-211">Requirement</span></span>| <span data-ttu-id="0529e-212">Значение</span><span class="sxs-lookup"><span data-stu-id="0529e-212">Value</span></span>|
|---|---|
|[<span data-ttu-id="0529e-213">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0529e-213">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0529e-214">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-214">1.0</span></span>|
|[<span data-ttu-id="0529e-215">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0529e-215">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0529e-216">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-216">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0529e-217">Пример</span><span class="sxs-lookup"><span data-stu-id="0529e-217">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officethemejavascriptapiofficeofficeofficetheme"></a><span data-ttu-id="0529e-218">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="0529e-218">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="0529e-219">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="0529e-219">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="0529e-220">Этот элемент поддерживается только в Outlook для Windows.</span><span class="sxs-lookup"><span data-stu-id="0529e-220">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="0529e-221">Использование цветов тем Office позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с помощью **файла > учетной записи office > пользовательского интерфейса темы**Office, которая применяется ко всем ведущим приложениям Office.</span><span class="sxs-lookup"><span data-stu-id="0529e-221">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="0529e-222">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="0529e-222">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="0529e-223">Тип</span><span class="sxs-lookup"><span data-stu-id="0529e-223">Type</span></span>

*   [<span data-ttu-id="0529e-224">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="0529e-224">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="0529e-225">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0529e-225">Properties:</span></span>

|<span data-ttu-id="0529e-226">Имя</span><span class="sxs-lookup"><span data-stu-id="0529e-226">Name</span></span>| <span data-ttu-id="0529e-227">Тип</span><span class="sxs-lookup"><span data-stu-id="0529e-227">Type</span></span>| <span data-ttu-id="0529e-228">Описание</span><span class="sxs-lookup"><span data-stu-id="0529e-228">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="0529e-229">String</span><span class="sxs-lookup"><span data-stu-id="0529e-229">String</span></span>|<span data-ttu-id="0529e-230">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0529e-230">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="0529e-231">String</span><span class="sxs-lookup"><span data-stu-id="0529e-231">String</span></span>|<span data-ttu-id="0529e-232">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0529e-232">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="0529e-233">String</span><span class="sxs-lookup"><span data-stu-id="0529e-233">String</span></span>|<span data-ttu-id="0529e-234">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0529e-234">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="0529e-235">String</span><span class="sxs-lookup"><span data-stu-id="0529e-235">String</span></span>|<span data-ttu-id="0529e-236">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0529e-236">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0529e-237">Requirements</span><span class="sxs-lookup"><span data-stu-id="0529e-237">Requirements</span></span>

|<span data-ttu-id="0529e-238">Требование</span><span class="sxs-lookup"><span data-stu-id="0529e-238">Requirement</span></span>| <span data-ttu-id="0529e-239">Значение</span><span class="sxs-lookup"><span data-stu-id="0529e-239">Value</span></span>|
|---|---|
|[<span data-ttu-id="0529e-240">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0529e-240">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0529e-241">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0529e-241">Preview</span></span>|
|[<span data-ttu-id="0529e-242">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0529e-242">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0529e-243">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-243">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0529e-244">Пример</span><span class="sxs-lookup"><span data-stu-id="0529e-244">Example</span></span>

```js
function applyOfficeTheme(){
  // Get office theme colors.
  var bodyBackgroundColor = Office.context.officeTheme.bodyBackgroundColor;
  var bodyForegroundColor = Office.context.officeTheme.bodyForegroundColor;
  var controlBackgroundColor = Office.context.officeTheme.controlBackgroundColor
  var controlForegroundColor = Office.context.officeTheme.controlForegroundColor;

  // Apply body background color to a CSS class.
  $('.body').css('background-color', bodyBackgroundColor);
}
```

<br>

---
---

#### <a name="platform-platformtypejavascriptapiofficeofficeplatformtype"></a><span data-ttu-id="0529e-245">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="0529e-245">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="0529e-246">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="0529e-246">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="0529e-247">Тип</span><span class="sxs-lookup"><span data-stu-id="0529e-247">Type</span></span>

*   [<span data-ttu-id="0529e-248">PlatformType</span><span class="sxs-lookup"><span data-stu-id="0529e-248">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="0529e-249">Requirements</span><span class="sxs-lookup"><span data-stu-id="0529e-249">Requirements</span></span>

|<span data-ttu-id="0529e-250">Требование</span><span class="sxs-lookup"><span data-stu-id="0529e-250">Requirement</span></span>| <span data-ttu-id="0529e-251">Значение</span><span class="sxs-lookup"><span data-stu-id="0529e-251">Value</span></span>|
|---|---|
|[<span data-ttu-id="0529e-252">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0529e-252">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0529e-253">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-253">1.0</span></span>|
|[<span data-ttu-id="0529e-254">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0529e-254">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0529e-255">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-255">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0529e-256">Пример</span><span class="sxs-lookup"><span data-stu-id="0529e-256">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupportjavascriptapiofficeofficerequirementsetsupport"></a><span data-ttu-id="0529e-257">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="0529e-257">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="0529e-258">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="0529e-258">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="0529e-259">Тип</span><span class="sxs-lookup"><span data-stu-id="0529e-259">Type</span></span>

*   [<span data-ttu-id="0529e-260">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="0529e-260">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="0529e-261">Requirements</span><span class="sxs-lookup"><span data-stu-id="0529e-261">Requirements</span></span>

|<span data-ttu-id="0529e-262">Требование</span><span class="sxs-lookup"><span data-stu-id="0529e-262">Requirement</span></span>| <span data-ttu-id="0529e-263">Значение</span><span class="sxs-lookup"><span data-stu-id="0529e-263">Value</span></span>|
|---|---|
|[<span data-ttu-id="0529e-264">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0529e-264">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0529e-265">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-265">1.0</span></span>|
|[<span data-ttu-id="0529e-266">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0529e-266">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0529e-267">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-267">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0529e-268">Пример</span><span class="sxs-lookup"><span data-stu-id="0529e-268">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.8")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettingsjavascriptapioutlookofficeroamingsettings"></a><span data-ttu-id="0529e-269">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="0529e-269">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="0529e-270">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="0529e-270">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="0529e-271">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="0529e-271">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="0529e-272">Тип</span><span class="sxs-lookup"><span data-stu-id="0529e-272">Type</span></span>

*   [<span data-ttu-id="0529e-273">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="0529e-273">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="0529e-274">Requirements</span><span class="sxs-lookup"><span data-stu-id="0529e-274">Requirements</span></span>

|<span data-ttu-id="0529e-275">Требование</span><span class="sxs-lookup"><span data-stu-id="0529e-275">Requirement</span></span>| <span data-ttu-id="0529e-276">Значение</span><span class="sxs-lookup"><span data-stu-id="0529e-276">Value</span></span>|
|---|---|
|[<span data-ttu-id="0529e-277">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0529e-277">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0529e-278">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-278">1.0</span></span>|
|[<span data-ttu-id="0529e-279">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0529e-279">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="0529e-280">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="0529e-280">Restricted</span></span>|
|[<span data-ttu-id="0529e-281">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0529e-281">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0529e-282">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-282">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-uijavascriptapiofficeofficeui"></a><span data-ttu-id="0529e-283">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="0529e-283">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="0529e-284">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="0529e-284">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="0529e-285">Тип</span><span class="sxs-lookup"><span data-stu-id="0529e-285">Type</span></span>

*   [<span data-ttu-id="0529e-286">UI</span><span class="sxs-lookup"><span data-stu-id="0529e-286">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="0529e-287">Requirements</span><span class="sxs-lookup"><span data-stu-id="0529e-287">Requirements</span></span>

|<span data-ttu-id="0529e-288">Требование</span><span class="sxs-lookup"><span data-stu-id="0529e-288">Requirement</span></span>| <span data-ttu-id="0529e-289">Значение</span><span class="sxs-lookup"><span data-stu-id="0529e-289">Value</span></span>|
|---|---|
|[<span data-ttu-id="0529e-290">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0529e-290">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="0529e-291">1.0</span><span class="sxs-lookup"><span data-stu-id="0529e-291">1.0</span></span>|
|[<span data-ttu-id="0529e-292">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0529e-292">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="0529e-293">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0529e-293">Compose or Read</span></span>|
