---
title: Office.context — набор требований 1.10
description: Office. Участники объектов context, доступные для Outlook надстройки с помощью API почтовых ящиков, устанавливают 1.10.
ms.date: 05/11/2021
localization_priority: Normal
ms.openlocfilehash: cb189dc3b7b51357dee8ac83bc61795b3ec47ae5
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52592062"
---
# <a name="context-mailbox-requirement-set-110"></a><span data-ttu-id="9d801-103">контекст (требования к почтовым ящикам 1.10)</span><span class="sxs-lookup"><span data-stu-id="9d801-103">context (Mailbox requirement set 1.10)</span></span>

### <a name="officecontext"></a><span data-ttu-id="9d801-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="9d801-104">[Office](office.md).context</span></span>

<span data-ttu-id="9d801-105">Office.context предоставляет общие интерфейсы, используемые надстройки во всех Office приложениях.</span><span class="sxs-lookup"><span data-stu-id="9d801-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="9d801-106">Этот список документов только те интерфейсы, которые используются Outlook надстройки. Полный список пространства имен Office.context см. в [ссылке Office.context в общем API.](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="9d801-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.10&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="9d801-107">Требования</span><span class="sxs-lookup"><span data-stu-id="9d801-107">Requirements</span></span>

|<span data-ttu-id="9d801-108">Требование</span><span class="sxs-lookup"><span data-stu-id="9d801-108">Requirement</span></span>| <span data-ttu-id="9d801-109">Значение</span><span class="sxs-lookup"><span data-stu-id="9d801-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d801-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9d801-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d801-111">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-111">1.1</span></span>|
|[<span data-ttu-id="9d801-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9d801-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d801-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="9d801-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="9d801-114">Properties</span></span>

| <span data-ttu-id="9d801-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="9d801-115">Property</span></span> | <span data-ttu-id="9d801-116">Режимы</span><span class="sxs-lookup"><span data-stu-id="9d801-116">Modes</span></span> | <span data-ttu-id="9d801-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="9d801-117">Return type</span></span> | <span data-ttu-id="9d801-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="9d801-118">Minimum</span></span><br><span data-ttu-id="9d801-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="9d801-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="9d801-120">auth</span><span class="sxs-lookup"><span data-stu-id="9d801-120">auth</span></span>](#auth-auth) | <span data-ttu-id="9d801-121">Создание</span><span class="sxs-lookup"><span data-stu-id="9d801-121">Compose</span></span><br><span data-ttu-id="9d801-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-122">Read</span></span> | [<span data-ttu-id="9d801-123">Auth</span><span class="sxs-lookup"><span data-stu-id="9d801-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9d801-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="9d801-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="9d801-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="9d801-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="9d801-126">Создание</span><span class="sxs-lookup"><span data-stu-id="9d801-126">Compose</span></span><br><span data-ttu-id="9d801-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-127">Read</span></span> | <span data-ttu-id="9d801-128">Строка</span><span class="sxs-lookup"><span data-stu-id="9d801-128">String</span></span> | [<span data-ttu-id="9d801-129">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9d801-130">диагностика</span><span class="sxs-lookup"><span data-stu-id="9d801-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="9d801-131">Создание</span><span class="sxs-lookup"><span data-stu-id="9d801-131">Compose</span></span><br><span data-ttu-id="9d801-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-132">Read</span></span> | [<span data-ttu-id="9d801-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="9d801-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9d801-134">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9d801-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="9d801-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="9d801-136">Создание</span><span class="sxs-lookup"><span data-stu-id="9d801-136">Compose</span></span><br><span data-ttu-id="9d801-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-137">Read</span></span> | <span data-ttu-id="9d801-138">Строка</span><span class="sxs-lookup"><span data-stu-id="9d801-138">String</span></span> | [<span data-ttu-id="9d801-139">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9d801-140">хост</span><span class="sxs-lookup"><span data-stu-id="9d801-140">host</span></span>](#host-hosttype) | <span data-ttu-id="9d801-141">Создание</span><span class="sxs-lookup"><span data-stu-id="9d801-141">Compose</span></span><br><span data-ttu-id="9d801-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-142">Read</span></span> | [<span data-ttu-id="9d801-143">HostType</span><span class="sxs-lookup"><span data-stu-id="9d801-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9d801-144">1.5</span><span class="sxs-lookup"><span data-stu-id="9d801-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9d801-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="9d801-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="9d801-146">Создание</span><span class="sxs-lookup"><span data-stu-id="9d801-146">Compose</span></span><br><span data-ttu-id="9d801-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-147">Read</span></span> | [<span data-ttu-id="9d801-148">Mailbox</span><span class="sxs-lookup"><span data-stu-id="9d801-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9d801-149">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9d801-150">платформа</span><span class="sxs-lookup"><span data-stu-id="9d801-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="9d801-151">Создание</span><span class="sxs-lookup"><span data-stu-id="9d801-151">Compose</span></span><br><span data-ttu-id="9d801-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-152">Read</span></span> | [<span data-ttu-id="9d801-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9d801-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9d801-154">1.5</span><span class="sxs-lookup"><span data-stu-id="9d801-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="9d801-155">требования</span><span class="sxs-lookup"><span data-stu-id="9d801-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="9d801-156">Создание</span><span class="sxs-lookup"><span data-stu-id="9d801-156">Compose</span></span><br><span data-ttu-id="9d801-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-157">Read</span></span> | [<span data-ttu-id="9d801-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="9d801-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9d801-159">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9d801-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="9d801-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="9d801-161">Создание</span><span class="sxs-lookup"><span data-stu-id="9d801-161">Compose</span></span><br><span data-ttu-id="9d801-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-162">Read</span></span> | [<span data-ttu-id="9d801-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9d801-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9d801-164">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="9d801-165">ui</span><span class="sxs-lookup"><span data-stu-id="9d801-165">ui</span></span>](#ui-ui) | <span data-ttu-id="9d801-166">Создание</span><span class="sxs-lookup"><span data-stu-id="9d801-166">Compose</span></span><br><span data-ttu-id="9d801-167">Чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-167">Read</span></span> | [<span data-ttu-id="9d801-168">UI</span><span class="sxs-lookup"><span data-stu-id="9d801-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.10&preserve-view=true) | [<span data-ttu-id="9d801-169">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="9d801-170">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="9d801-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="9d801-171">Auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="9d801-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="9d801-172">Поддерживает один вход [(SSO),](../../../outlook/authenticate-a-user-with-an-sso-token.md) предоставляя метод, который позволяет Office приложению получить маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d801-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="9d801-173">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="9d801-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="9d801-174">Тип</span><span class="sxs-lookup"><span data-stu-id="9d801-174">Type</span></span>

*   [<span data-ttu-id="9d801-175">Auth</span><span class="sxs-lookup"><span data-stu-id="9d801-175">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="9d801-176">Требования</span><span class="sxs-lookup"><span data-stu-id="9d801-176">Requirements</span></span>

|<span data-ttu-id="9d801-177">Требование</span><span class="sxs-lookup"><span data-stu-id="9d801-177">Requirement</span></span>| <span data-ttu-id="9d801-178">Значение</span><span class="sxs-lookup"><span data-stu-id="9d801-178">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d801-179">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9d801-179">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d801-180">1.10</span><span class="sxs-lookup"><span data-stu-id="9d801-180">1.10</span></span>|
|[<span data-ttu-id="9d801-181">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9d801-181">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d801-182">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-182">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9d801-183">Пример</span><span class="sxs-lookup"><span data-stu-id="9d801-183">Example</span></span>

```js
Office.context.auth.getAccessTokenAsync(function(result) {
    if (result.status === "succeeded") {
        var token = result.value;
        // ...
    } else {
        console.log("Error obtaining token", result.error);
    }
});
```

<br>

---
---

#### <a name="contentlanguage-string"></a><span data-ttu-id="9d801-184">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="9d801-184">contentLanguage: String</span></span>

<span data-ttu-id="9d801-185">Получает локализ (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="9d801-185">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="9d801-186">Это значение отражает текущий параметр Язык редактирования, указанный в файле > `contentLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="9d801-186">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="9d801-187">Тип</span><span class="sxs-lookup"><span data-stu-id="9d801-187">Type</span></span>

*   <span data-ttu-id="9d801-188">String</span><span class="sxs-lookup"><span data-stu-id="9d801-188">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9d801-189">Требования</span><span class="sxs-lookup"><span data-stu-id="9d801-189">Requirements</span></span>

|<span data-ttu-id="9d801-190">Требование</span><span class="sxs-lookup"><span data-stu-id="9d801-190">Requirement</span></span>| <span data-ttu-id="9d801-191">Значение</span><span class="sxs-lookup"><span data-stu-id="9d801-191">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d801-192">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9d801-192">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d801-193">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-193">1.1</span></span>|
|[<span data-ttu-id="9d801-194">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9d801-194">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d801-195">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-195">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9d801-196">Пример</span><span class="sxs-lookup"><span data-stu-id="9d801-196">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="9d801-197">диагностика: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="9d801-197">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="9d801-198">Получает сведения об среде, в которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="9d801-198">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="9d801-199">Тип</span><span class="sxs-lookup"><span data-stu-id="9d801-199">Type</span></span>

*   [<span data-ttu-id="9d801-200">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="9d801-200">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="9d801-201">Требования</span><span class="sxs-lookup"><span data-stu-id="9d801-201">Requirements</span></span>

|<span data-ttu-id="9d801-202">Требование</span><span class="sxs-lookup"><span data-stu-id="9d801-202">Requirement</span></span>| <span data-ttu-id="9d801-203">Значение</span><span class="sxs-lookup"><span data-stu-id="9d801-203">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d801-204">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9d801-204">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d801-205">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-205">1.1</span></span>|
|[<span data-ttu-id="9d801-206">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9d801-206">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d801-207">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-207">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9d801-208">Пример</span><span class="sxs-lookup"><span data-stu-id="9d801-208">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="9d801-209">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="9d801-209">displayLanguage: String</span></span>

<span data-ttu-id="9d801-210">Получает локализ (язык) в формате языковых тегов RFC 1766, заданный пользователем для пользовательского интерфейса Office клиентского приложения.</span><span class="sxs-lookup"><span data-stu-id="9d801-210">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="9d801-211">Это значение отражает текущий параметр Язык отображения, указанный в файле > `displayLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="9d801-211">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="9d801-212">Тип</span><span class="sxs-lookup"><span data-stu-id="9d801-212">Type</span></span>

*   <span data-ttu-id="9d801-213">String</span><span class="sxs-lookup"><span data-stu-id="9d801-213">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="9d801-214">Требования</span><span class="sxs-lookup"><span data-stu-id="9d801-214">Requirements</span></span>

|<span data-ttu-id="9d801-215">Требование</span><span class="sxs-lookup"><span data-stu-id="9d801-215">Requirement</span></span>| <span data-ttu-id="9d801-216">Значение</span><span class="sxs-lookup"><span data-stu-id="9d801-216">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d801-217">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9d801-217">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d801-218">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-218">1.1</span></span>|
|[<span data-ttu-id="9d801-219">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9d801-219">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d801-220">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-220">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9d801-221">Пример</span><span class="sxs-lookup"><span data-stu-id="9d801-221">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="9d801-222">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="9d801-222">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="9d801-223">Получает Office приложение, в которое размещена надстройка.</span><span class="sxs-lookup"><span data-stu-id="9d801-223">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="9d801-224">Кроме того, для получения хоста можно использовать [свойство Office.context.diagnostics.](#diagnostics-contextinformation)</span><span class="sxs-lookup"><span data-stu-id="9d801-224">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="9d801-225">Тип</span><span class="sxs-lookup"><span data-stu-id="9d801-225">Type</span></span>

*   [<span data-ttu-id="9d801-226">HostType</span><span class="sxs-lookup"><span data-stu-id="9d801-226">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="9d801-227">Требования</span><span class="sxs-lookup"><span data-stu-id="9d801-227">Requirements</span></span>

|<span data-ttu-id="9d801-228">Требование</span><span class="sxs-lookup"><span data-stu-id="9d801-228">Requirement</span></span>| <span data-ttu-id="9d801-229">Значение</span><span class="sxs-lookup"><span data-stu-id="9d801-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d801-230">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9d801-230">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d801-231">1.5</span><span class="sxs-lookup"><span data-stu-id="9d801-231">1.5</span></span>|
|[<span data-ttu-id="9d801-232">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9d801-232">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d801-233">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9d801-234">Пример</span><span class="sxs-lookup"><span data-stu-id="9d801-234">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="9d801-235">платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="9d801-235">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="9d801-236">Предоставляет платформу, на которой запущена надстройка.</span><span class="sxs-lookup"><span data-stu-id="9d801-236">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="9d801-237">Кроме того, для получения [платформы можно использовать свойство Office.context.diagnostics.](#diagnostics-contextinformation)</span><span class="sxs-lookup"><span data-stu-id="9d801-237">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9d801-238">Тип</span><span class="sxs-lookup"><span data-stu-id="9d801-238">Type</span></span>

*   [<span data-ttu-id="9d801-239">PlatformType</span><span class="sxs-lookup"><span data-stu-id="9d801-239">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="9d801-240">Требования</span><span class="sxs-lookup"><span data-stu-id="9d801-240">Requirements</span></span>

|<span data-ttu-id="9d801-241">Требование</span><span class="sxs-lookup"><span data-stu-id="9d801-241">Requirement</span></span>| <span data-ttu-id="9d801-242">Значение</span><span class="sxs-lookup"><span data-stu-id="9d801-242">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d801-243">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="9d801-243">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d801-244">1.5</span><span class="sxs-lookup"><span data-stu-id="9d801-244">1.5</span></span>|
|[<span data-ttu-id="9d801-245">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9d801-245">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d801-246">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-246">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9d801-247">Пример</span><span class="sxs-lookup"><span data-stu-id="9d801-247">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="9d801-248">требования: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="9d801-248">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="9d801-249">Предоставляет метод определения, какие наборы требований поддерживаются в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="9d801-249">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="9d801-250">Тип</span><span class="sxs-lookup"><span data-stu-id="9d801-250">Type</span></span>

*   [<span data-ttu-id="9d801-251">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="9d801-251">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="9d801-252">Требования</span><span class="sxs-lookup"><span data-stu-id="9d801-252">Requirements</span></span>

|<span data-ttu-id="9d801-253">Требование</span><span class="sxs-lookup"><span data-stu-id="9d801-253">Requirement</span></span>| <span data-ttu-id="9d801-254">Значение</span><span class="sxs-lookup"><span data-stu-id="9d801-254">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d801-255">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9d801-255">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d801-256">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-256">1.1</span></span>|
|[<span data-ttu-id="9d801-257">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9d801-257">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d801-258">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-258">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="9d801-259">Пример</span><span class="sxs-lookup"><span data-stu-id="9d801-259">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="9d801-260">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="9d801-260">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="9d801-261">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="9d801-261">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="9d801-262">Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранимой в почтовом ящике пользователя, чтобы она была доступна этой надстройке, когда она запущена из любого клиента Outlook, используемого для доступа к этому `RoamingSettings` почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="9d801-262">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="9d801-263">Тип</span><span class="sxs-lookup"><span data-stu-id="9d801-263">Type</span></span>

*   [<span data-ttu-id="9d801-264">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="9d801-264">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="9d801-265">Требования</span><span class="sxs-lookup"><span data-stu-id="9d801-265">Requirements</span></span>

|<span data-ttu-id="9d801-266">Требование</span><span class="sxs-lookup"><span data-stu-id="9d801-266">Requirement</span></span>| <span data-ttu-id="9d801-267">Значение</span><span class="sxs-lookup"><span data-stu-id="9d801-267">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d801-268">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9d801-268">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d801-269">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-269">1.1</span></span>|
|[<span data-ttu-id="9d801-270">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="9d801-270">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="9d801-271">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="9d801-271">Restricted</span></span>|
|[<span data-ttu-id="9d801-272">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9d801-272">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d801-273">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-273">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="9d801-274">ui: [пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="9d801-274">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="9d801-275">Предоставляет объекты и методы, которые можно использовать для создания и управления компонентами пользовательского интерфейса, такими как диалоговое окно, в Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="9d801-275">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="9d801-276">Тип</span><span class="sxs-lookup"><span data-stu-id="9d801-276">Type</span></span>

*   [<span data-ttu-id="9d801-277">UI</span><span class="sxs-lookup"><span data-stu-id="9d801-277">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="9d801-278">Требования</span><span class="sxs-lookup"><span data-stu-id="9d801-278">Requirements</span></span>

|<span data-ttu-id="9d801-279">Требование</span><span class="sxs-lookup"><span data-stu-id="9d801-279">Requirement</span></span>| <span data-ttu-id="9d801-280">Значение</span><span class="sxs-lookup"><span data-stu-id="9d801-280">Value</span></span>|
|---|---|
|[<span data-ttu-id="9d801-281">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="9d801-281">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="9d801-282">1.1</span><span class="sxs-lookup"><span data-stu-id="9d801-282">1.1</span></span>|
|[<span data-ttu-id="9d801-283">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="9d801-283">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="9d801-284">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="9d801-284">Compose or Read</span></span>|
