---
title: Office.context — предварительная версия набора обязательных элементов
description: Office. Участники объектов Context, доступные для Outlook надстройки с помощью набора требований к API API почтовых ящиков.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 59b1cce579afe69384e41a6f31cc70c8cec25bea
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591074"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="6d376-103">контекст (набор требований предварительного просмотра почтовых ящиков)</span><span class="sxs-lookup"><span data-stu-id="6d376-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="6d376-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="6d376-104">[Office](office.md).context</span></span>

<span data-ttu-id="6d376-105">Office.context предоставляет общие интерфейсы, используемые надстройки во всех Office приложениях.</span><span class="sxs-lookup"><span data-stu-id="6d376-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="6d376-106">Этот список документов только те интерфейсы, которые используются Outlook надстройки. Полный список пространства имен Office.context см. в [ссылке Office.context в общем API.](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="6d376-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d376-107">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-107">Requirements</span></span>

|<span data-ttu-id="6d376-108">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-108">Requirement</span></span>| <span data-ttu-id="6d376-109">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6d376-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-111">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-111">1.1</span></span>|
|[<span data-ttu-id="6d376-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="6d376-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="6d376-114">Properties</span></span>

| <span data-ttu-id="6d376-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="6d376-115">Property</span></span> | <span data-ttu-id="6d376-116">Режимы</span><span class="sxs-lookup"><span data-stu-id="6d376-116">Modes</span></span> | <span data-ttu-id="6d376-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="6d376-117">Return type</span></span> | <span data-ttu-id="6d376-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="6d376-118">Minimum</span></span><br><span data-ttu-id="6d376-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="6d376-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="6d376-120">auth</span><span class="sxs-lookup"><span data-stu-id="6d376-120">auth</span></span>](#auth-auth) | <span data-ttu-id="6d376-121">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-121">Compose</span></span><br><span data-ttu-id="6d376-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-122">Read</span></span> | [<span data-ttu-id="6d376-123">Auth</span><span class="sxs-lookup"><span data-stu-id="6d376-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="6d376-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="6d376-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="6d376-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="6d376-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="6d376-126">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-126">Compose</span></span><br><span data-ttu-id="6d376-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-127">Read</span></span> | <span data-ttu-id="6d376-128">Строка</span><span class="sxs-lookup"><span data-stu-id="6d376-128">String</span></span> | [<span data-ttu-id="6d376-129">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6d376-130">диагностика</span><span class="sxs-lookup"><span data-stu-id="6d376-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="6d376-131">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-131">Compose</span></span><br><span data-ttu-id="6d376-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-132">Read</span></span> | [<span data-ttu-id="6d376-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="6d376-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="6d376-134">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6d376-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="6d376-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="6d376-136">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-136">Compose</span></span><br><span data-ttu-id="6d376-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-137">Read</span></span> | <span data-ttu-id="6d376-138">Строка</span><span class="sxs-lookup"><span data-stu-id="6d376-138">String</span></span> | [<span data-ttu-id="6d376-139">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6d376-140">хост</span><span class="sxs-lookup"><span data-stu-id="6d376-140">host</span></span>](#host-hosttype) | <span data-ttu-id="6d376-141">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-141">Compose</span></span><br><span data-ttu-id="6d376-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-142">Read</span></span> | [<span data-ttu-id="6d376-143">HostType</span><span class="sxs-lookup"><span data-stu-id="6d376-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="6d376-144">1.5</span><span class="sxs-lookup"><span data-stu-id="6d376-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="6d376-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="6d376-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="6d376-146">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-146">Compose</span></span><br><span data-ttu-id="6d376-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-147">Read</span></span> | [<span data-ttu-id="6d376-148">Mailbox</span><span class="sxs-lookup"><span data-stu-id="6d376-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="6d376-149">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6d376-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="6d376-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="6d376-151">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-151">Compose</span></span><br><span data-ttu-id="6d376-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-152">Read</span></span> | [<span data-ttu-id="6d376-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="6d376-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="6d376-154">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="6d376-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="6d376-155">платформа</span><span class="sxs-lookup"><span data-stu-id="6d376-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="6d376-156">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-156">Compose</span></span><br><span data-ttu-id="6d376-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-157">Read</span></span> | [<span data-ttu-id="6d376-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="6d376-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="6d376-159">1.5</span><span class="sxs-lookup"><span data-stu-id="6d376-159">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="6d376-160">требования</span><span class="sxs-lookup"><span data-stu-id="6d376-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="6d376-161">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-161">Compose</span></span><br><span data-ttu-id="6d376-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-162">Read</span></span> | [<span data-ttu-id="6d376-163">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="6d376-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="6d376-164">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6d376-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="6d376-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="6d376-166">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-166">Compose</span></span><br><span data-ttu-id="6d376-167">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-167">Read</span></span> | [<span data-ttu-id="6d376-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6d376-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="6d376-169">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="6d376-170">ui</span><span class="sxs-lookup"><span data-stu-id="6d376-170">ui</span></span>](#ui-ui) | <span data-ttu-id="6d376-171">Создание</span><span class="sxs-lookup"><span data-stu-id="6d376-171">Compose</span></span><br><span data-ttu-id="6d376-172">Чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-172">Read</span></span> | [<span data-ttu-id="6d376-173">UI</span><span class="sxs-lookup"><span data-stu-id="6d376-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="6d376-174">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="6d376-175">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="6d376-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="6d376-176">Auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="6d376-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="6d376-177">Поддерживает один вход [(SSO),](../../../outlook/authenticate-a-user-with-an-sso-token.md) предоставляя метод, который позволяет Office приложению получить маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="6d376-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="6d376-178">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="6d376-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="6d376-179">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-179">Type</span></span>

*   [<span data-ttu-id="6d376-180">Auth</span><span class="sxs-lookup"><span data-stu-id="6d376-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="6d376-181">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-181">Requirements</span></span>

|<span data-ttu-id="6d376-182">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-182">Requirement</span></span>| <span data-ttu-id="6d376-183">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="6d376-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-185">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="6d376-185">Preview</span></span>|
|[<span data-ttu-id="6d376-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d376-188">Пример</span><span class="sxs-lookup"><span data-stu-id="6d376-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="6d376-189">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="6d376-189">contentLanguage: String</span></span>

<span data-ttu-id="6d376-190">Получает локализ (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="6d376-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="6d376-191">Это значение отражает текущий параметр Язык редактирования, указанный в файле > `contentLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="6d376-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="6d376-192">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-192">Type</span></span>

*   <span data-ttu-id="6d376-193">String</span><span class="sxs-lookup"><span data-stu-id="6d376-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d376-194">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-194">Requirements</span></span>

|<span data-ttu-id="6d376-195">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-195">Requirement</span></span>| <span data-ttu-id="6d376-196">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6d376-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-198">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-198">1.1</span></span>|
|[<span data-ttu-id="6d376-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d376-201">Пример</span><span class="sxs-lookup"><span data-stu-id="6d376-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="6d376-202">диагностика: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="6d376-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="6d376-203">Получает сведения об среде, в которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="6d376-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="6d376-204">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-204">Type</span></span>

*   [<span data-ttu-id="6d376-205">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="6d376-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="6d376-206">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-206">Requirements</span></span>

|<span data-ttu-id="6d376-207">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-207">Requirement</span></span>| <span data-ttu-id="6d376-208">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-209">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6d376-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-210">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-210">1.1</span></span>|
|[<span data-ttu-id="6d376-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-212">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d376-213">Пример</span><span class="sxs-lookup"><span data-stu-id="6d376-213">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="6d376-214">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="6d376-214">displayLanguage: String</span></span>

<span data-ttu-id="6d376-215">Получает локализ (язык) в формате языковых тегов RFC 1766, заданный пользователем для пользовательского интерфейса Office клиентского приложения.</span><span class="sxs-lookup"><span data-stu-id="6d376-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="6d376-216">Это значение отражает текущий параметр Язык отображения, указанный в файле > `displayLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="6d376-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="6d376-217">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-217">Type</span></span>

*   <span data-ttu-id="6d376-218">String</span><span class="sxs-lookup"><span data-stu-id="6d376-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="6d376-219">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-219">Requirements</span></span>

|<span data-ttu-id="6d376-220">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-220">Requirement</span></span>| <span data-ttu-id="6d376-221">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-222">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6d376-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-223">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-223">1.1</span></span>|
|[<span data-ttu-id="6d376-224">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-225">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d376-226">Пример</span><span class="sxs-lookup"><span data-stu-id="6d376-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="6d376-227">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="6d376-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="6d376-228">Получает Office приложение, в которое размещена надстройка.</span><span class="sxs-lookup"><span data-stu-id="6d376-228">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="6d376-229">Кроме того, для получения хоста можно использовать [свойство Office.context.diagnostics.](#diagnostics-contextinformation)</span><span class="sxs-lookup"><span data-stu-id="6d376-229">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="6d376-230">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-230">Type</span></span>

*   [<span data-ttu-id="6d376-231">HostType</span><span class="sxs-lookup"><span data-stu-id="6d376-231">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="6d376-232">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-232">Requirements</span></span>

|<span data-ttu-id="6d376-233">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-233">Requirement</span></span>| <span data-ttu-id="6d376-234">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-235">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="6d376-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-236">1.5</span><span class="sxs-lookup"><span data-stu-id="6d376-236">1.5</span></span>|
|[<span data-ttu-id="6d376-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d376-239">Пример</span><span class="sxs-lookup"><span data-stu-id="6d376-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="6d376-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="6d376-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="6d376-241">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="6d376-241">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="6d376-242">Этот член поддерживается только в Outlook на Windows.</span><span class="sxs-lookup"><span data-stu-id="6d376-242">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="6d376-243">Использование Office тем позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с пользовательским интерфейсом **File > Office Account > Office Theme,** который применяется во всех Office клиентских приложениях.</span><span class="sxs-lookup"><span data-stu-id="6d376-243">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="6d376-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="6d376-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="6d376-245">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-245">Type</span></span>

*   [<span data-ttu-id="6d376-246">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="6d376-246">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="6d376-247">Свойства</span><span class="sxs-lookup"><span data-stu-id="6d376-247">Properties</span></span>

|<span data-ttu-id="6d376-248">Имя</span><span class="sxs-lookup"><span data-stu-id="6d376-248">Name</span></span>| <span data-ttu-id="6d376-249">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-249">Type</span></span>| <span data-ttu-id="6d376-250">Описание</span><span class="sxs-lookup"><span data-stu-id="6d376-250">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="6d376-251">Строка</span><span class="sxs-lookup"><span data-stu-id="6d376-251">String</span></span>|<span data-ttu-id="6d376-252">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="6d376-252">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="6d376-253">String</span><span class="sxs-lookup"><span data-stu-id="6d376-253">String</span></span>|<span data-ttu-id="6d376-254">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="6d376-254">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="6d376-255">String</span><span class="sxs-lookup"><span data-stu-id="6d376-255">String</span></span>|<span data-ttu-id="6d376-256">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="6d376-256">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="6d376-257">String</span><span class="sxs-lookup"><span data-stu-id="6d376-257">String</span></span>|<span data-ttu-id="6d376-258">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="6d376-258">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="6d376-259">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-259">Requirements</span></span>

|<span data-ttu-id="6d376-260">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-260">Requirement</span></span>| <span data-ttu-id="6d376-261">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-261">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-262">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="6d376-262">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-263">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="6d376-263">Preview</span></span>|
|[<span data-ttu-id="6d376-264">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-264">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-265">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-265">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d376-266">Пример</span><span class="sxs-lookup"><span data-stu-id="6d376-266">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="6d376-267">платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="6d376-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="6d376-268">Предоставляет платформу, на которой запущена надстройка.</span><span class="sxs-lookup"><span data-stu-id="6d376-268">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="6d376-269">Кроме того, для получения [платформы можно использовать свойство Office.context.diagnostics.](#diagnostics-contextinformation)</span><span class="sxs-lookup"><span data-stu-id="6d376-269">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="6d376-270">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-270">Type</span></span>

*   [<span data-ttu-id="6d376-271">PlatformType</span><span class="sxs-lookup"><span data-stu-id="6d376-271">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="6d376-272">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-272">Requirements</span></span>

|<span data-ttu-id="6d376-273">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-273">Requirement</span></span>| <span data-ttu-id="6d376-274">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-274">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-275">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="6d376-275">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-276">1.5</span><span class="sxs-lookup"><span data-stu-id="6d376-276">1.5</span></span>|
|[<span data-ttu-id="6d376-277">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-277">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-278">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d376-279">Пример</span><span class="sxs-lookup"><span data-stu-id="6d376-279">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="6d376-280">требования: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="6d376-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="6d376-281">Предоставляет метод определения, какие наборы требований поддерживаются в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="6d376-281">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="6d376-282">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-282">Type</span></span>

*   [<span data-ttu-id="6d376-283">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="6d376-283">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="6d376-284">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-284">Requirements</span></span>

|<span data-ttu-id="6d376-285">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-285">Requirement</span></span>| <span data-ttu-id="6d376-286">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-287">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6d376-287">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-288">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-288">1.1</span></span>|
|[<span data-ttu-id="6d376-289">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-289">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-290">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="6d376-291">Пример</span><span class="sxs-lookup"><span data-stu-id="6d376-291">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="6d376-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="6d376-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="6d376-293">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="6d376-293">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="6d376-294">Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранимой в почтовом ящике пользователя, чтобы она была доступна этой надстройке, когда она запущена из любого клиента Outlook, используемого для доступа к этому `RoamingSettings` почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="6d376-294">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="6d376-295">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-295">Type</span></span>

*   [<span data-ttu-id="6d376-296">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="6d376-296">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="6d376-297">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-297">Requirements</span></span>

|<span data-ttu-id="6d376-298">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-298">Requirement</span></span>| <span data-ttu-id="6d376-299">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-300">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6d376-300">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-301">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-301">1.1</span></span>|
|[<span data-ttu-id="6d376-302">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="6d376-302">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="6d376-303">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="6d376-303">Restricted</span></span>|
|[<span data-ttu-id="6d376-304">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-304">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-305">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-305">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="6d376-306">ui: [пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="6d376-306">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="6d376-307">Предоставляет объекты и методы, которые можно использовать для создания и управления компонентами пользовательского интерфейса, такими как диалоговое окно, в Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="6d376-307">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="6d376-308">Тип</span><span class="sxs-lookup"><span data-stu-id="6d376-308">Type</span></span>

*   [<span data-ttu-id="6d376-309">UI</span><span class="sxs-lookup"><span data-stu-id="6d376-309">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="6d376-310">Требования</span><span class="sxs-lookup"><span data-stu-id="6d376-310">Requirements</span></span>

|<span data-ttu-id="6d376-311">Требование</span><span class="sxs-lookup"><span data-stu-id="6d376-311">Requirement</span></span>| <span data-ttu-id="6d376-312">Значение</span><span class="sxs-lookup"><span data-stu-id="6d376-312">Value</span></span>|
|---|---|
|[<span data-ttu-id="6d376-313">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="6d376-313">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="6d376-314">1.1</span><span class="sxs-lookup"><span data-stu-id="6d376-314">1.1</span></span>|
|[<span data-ttu-id="6d376-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="6d376-315">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="6d376-316">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="6d376-316">Compose or Read</span></span>|
