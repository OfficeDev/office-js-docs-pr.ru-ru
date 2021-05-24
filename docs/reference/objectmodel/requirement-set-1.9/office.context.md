---
title: Office.context — набор требований 1.9
description: Office. Участники объектов Context, доступные для Outlook надстройки с помощью API почтовых ящиков, устанавливают 1.9.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: f45eec7ce638f4bbb97ad4be9f2ba089905c631d
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52590521"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="b9d44-103">контекст (набор требований к почтовым ящикам 1.9)</span><span class="sxs-lookup"><span data-stu-id="b9d44-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="b9d44-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="b9d44-104">[Office](office.md).context</span></span>

<span data-ttu-id="b9d44-105">Office.context предоставляет общие интерфейсы, используемые надстройки во всех Office приложениях.</span><span class="sxs-lookup"><span data-stu-id="b9d44-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="b9d44-106">Этот список документов только те интерфейсы, которые используются Outlook надстройки. Полный список пространства имен Office.context см. в [ссылке Office.context в общем API.](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true)</span><span class="sxs-lookup"><span data-stu-id="b9d44-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d44-107">Требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-107">Requirements</span></span>

|<span data-ttu-id="b9d44-108">Требование</span><span class="sxs-lookup"><span data-stu-id="b9d44-108">Requirement</span></span>| <span data-ttu-id="b9d44-109">Значение</span><span class="sxs-lookup"><span data-stu-id="b9d44-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d44-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9d44-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9d44-111">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-111">1.1</span></span>|
|[<span data-ttu-id="b9d44-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9d44-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9d44-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-113">Compose or Read</span></span>|

## <a name="properties"></a><span data-ttu-id="b9d44-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="b9d44-114">Properties</span></span>

| <span data-ttu-id="b9d44-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="b9d44-115">Property</span></span> | <span data-ttu-id="b9d44-116">Режимы</span><span class="sxs-lookup"><span data-stu-id="b9d44-116">Modes</span></span> | <span data-ttu-id="b9d44-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="b9d44-117">Return type</span></span> | <span data-ttu-id="b9d44-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="b9d44-118">Minimum</span></span><br><span data-ttu-id="b9d44-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="b9d44-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="b9d44-120">auth</span><span class="sxs-lookup"><span data-stu-id="b9d44-120">auth</span></span>](#auth-auth) | <span data-ttu-id="b9d44-121">Создание</span><span class="sxs-lookup"><span data-stu-id="b9d44-121">Compose</span></span><br><span data-ttu-id="b9d44-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-122">Read</span></span> | [<span data-ttu-id="b9d44-123">Auth</span><span class="sxs-lookup"><span data-stu-id="b9d44-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b9d44-124">IdentityAPI 1.3</span><span class="sxs-lookup"><span data-stu-id="b9d44-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="b9d44-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="b9d44-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="b9d44-126">Создание</span><span class="sxs-lookup"><span data-stu-id="b9d44-126">Compose</span></span><br><span data-ttu-id="b9d44-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-127">Read</span></span> | <span data-ttu-id="b9d44-128">Строка</span><span class="sxs-lookup"><span data-stu-id="b9d44-128">String</span></span> | [<span data-ttu-id="b9d44-129">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b9d44-130">диагностика</span><span class="sxs-lookup"><span data-stu-id="b9d44-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="b9d44-131">Создание</span><span class="sxs-lookup"><span data-stu-id="b9d44-131">Compose</span></span><br><span data-ttu-id="b9d44-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-132">Read</span></span> | [<span data-ttu-id="b9d44-133">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b9d44-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b9d44-134">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b9d44-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="b9d44-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="b9d44-136">Создание</span><span class="sxs-lookup"><span data-stu-id="b9d44-136">Compose</span></span><br><span data-ttu-id="b9d44-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-137">Read</span></span> | <span data-ttu-id="b9d44-138">Строка</span><span class="sxs-lookup"><span data-stu-id="b9d44-138">String</span></span> | [<span data-ttu-id="b9d44-139">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b9d44-140">хост</span><span class="sxs-lookup"><span data-stu-id="b9d44-140">host</span></span>](#host-hosttype) | <span data-ttu-id="b9d44-141">Создание</span><span class="sxs-lookup"><span data-stu-id="b9d44-141">Compose</span></span><br><span data-ttu-id="b9d44-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-142">Read</span></span> | [<span data-ttu-id="b9d44-143">HostType</span><span class="sxs-lookup"><span data-stu-id="b9d44-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b9d44-144">1.5</span><span class="sxs-lookup"><span data-stu-id="b9d44-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b9d44-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="b9d44-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="b9d44-146">Создание</span><span class="sxs-lookup"><span data-stu-id="b9d44-146">Compose</span></span><br><span data-ttu-id="b9d44-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-147">Read</span></span> | [<span data-ttu-id="b9d44-148">Mailbox</span><span class="sxs-lookup"><span data-stu-id="b9d44-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b9d44-149">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b9d44-150">платформа</span><span class="sxs-lookup"><span data-stu-id="b9d44-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="b9d44-151">Создание</span><span class="sxs-lookup"><span data-stu-id="b9d44-151">Compose</span></span><br><span data-ttu-id="b9d44-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-152">Read</span></span> | [<span data-ttu-id="b9d44-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b9d44-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b9d44-154">1.5</span><span class="sxs-lookup"><span data-stu-id="b9d44-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="b9d44-155">требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="b9d44-156">Создание</span><span class="sxs-lookup"><span data-stu-id="b9d44-156">Compose</span></span><br><span data-ttu-id="b9d44-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-157">Read</span></span> | [<span data-ttu-id="b9d44-158">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b9d44-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b9d44-159">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b9d44-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="b9d44-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="b9d44-161">Создание</span><span class="sxs-lookup"><span data-stu-id="b9d44-161">Compose</span></span><br><span data-ttu-id="b9d44-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-162">Read</span></span> | [<span data-ttu-id="b9d44-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b9d44-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b9d44-164">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="b9d44-165">ui</span><span class="sxs-lookup"><span data-stu-id="b9d44-165">ui</span></span>](#ui-ui) | <span data-ttu-id="b9d44-166">Создание</span><span class="sxs-lookup"><span data-stu-id="b9d44-166">Compose</span></span><br><span data-ttu-id="b9d44-167">Чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-167">Read</span></span> | [<span data-ttu-id="b9d44-168">UI</span><span class="sxs-lookup"><span data-stu-id="b9d44-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="b9d44-169">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="b9d44-170">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="b9d44-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="b9d44-171">Auth: [Auth](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="b9d44-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="b9d44-172">Поддерживает один вход [(SSO),](../../../outlook/authenticate-a-user-with-an-sso-token.md) предоставляя метод, который позволяет Office приложению получить маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="b9d44-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="b9d44-173">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="b9d44-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="b9d44-174">См. набор требований [IdentityAPI 1.3.](../../requirement-sets/identity-api-requirement-sets.md)</span><span class="sxs-lookup"><span data-stu-id="b9d44-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="b9d44-175">Тип</span><span class="sxs-lookup"><span data-stu-id="b9d44-175">Type</span></span>

*   [<span data-ttu-id="b9d44-176">Auth</span><span class="sxs-lookup"><span data-stu-id="b9d44-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="b9d44-177">Требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-177">Requirements</span></span>

|<span data-ttu-id="b9d44-178">Требование</span><span class="sxs-lookup"><span data-stu-id="b9d44-178">Requirement</span></span>| <span data-ttu-id="b9d44-179">Значение</span><span class="sxs-lookup"><span data-stu-id="b9d44-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d44-180">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b9d44-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9d44-181">Недоступно</span><span class="sxs-lookup"><span data-stu-id="b9d44-181">N/A</span></span>|
|[<span data-ttu-id="b9d44-182">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9d44-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9d44-183">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d44-184">Пример</span><span class="sxs-lookup"><span data-stu-id="b9d44-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="b9d44-185">contentLanguage: String</span><span class="sxs-lookup"><span data-stu-id="b9d44-185">contentLanguage: String</span></span>

<span data-ttu-id="b9d44-186">Получает локализ (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="b9d44-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="b9d44-187">Это значение отражает текущий параметр Язык редактирования, указанный в файле > `contentLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="b9d44-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d44-188">Тип</span><span class="sxs-lookup"><span data-stu-id="b9d44-188">Type</span></span>

*   <span data-ttu-id="b9d44-189">String</span><span class="sxs-lookup"><span data-stu-id="b9d44-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d44-190">Требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-190">Requirements</span></span>

|<span data-ttu-id="b9d44-191">Требование</span><span class="sxs-lookup"><span data-stu-id="b9d44-191">Requirement</span></span>| <span data-ttu-id="b9d44-192">Значение</span><span class="sxs-lookup"><span data-stu-id="b9d44-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d44-193">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9d44-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9d44-194">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-194">1.1</span></span>|
|[<span data-ttu-id="b9d44-195">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9d44-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9d44-196">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d44-197">Пример</span><span class="sxs-lookup"><span data-stu-id="b9d44-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="b9d44-198">диагностика: [ContextInformation](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b9d44-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="b9d44-199">Получает сведения об среде, в которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="b9d44-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d44-200">Тип</span><span class="sxs-lookup"><span data-stu-id="b9d44-200">Type</span></span>

*   [<span data-ttu-id="b9d44-201">ContextInformation</span><span class="sxs-lookup"><span data-stu-id="b9d44-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="b9d44-202">Требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-202">Requirements</span></span>

|<span data-ttu-id="b9d44-203">Требование</span><span class="sxs-lookup"><span data-stu-id="b9d44-203">Requirement</span></span>| <span data-ttu-id="b9d44-204">Значение</span><span class="sxs-lookup"><span data-stu-id="b9d44-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d44-205">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9d44-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9d44-206">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-206">1.1</span></span>|
|[<span data-ttu-id="b9d44-207">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9d44-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9d44-208">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d44-209">Пример</span><span class="sxs-lookup"><span data-stu-id="b9d44-209">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="b9d44-210">displayLanguage: String</span><span class="sxs-lookup"><span data-stu-id="b9d44-210">displayLanguage: String</span></span>

<span data-ttu-id="b9d44-211">Получает локализ (язык) в формате языковых тегов RFC 1766, заданный пользователем для пользовательского интерфейса Office клиентского приложения.</span><span class="sxs-lookup"><span data-stu-id="b9d44-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="b9d44-212">Это значение отражает текущий параметр Язык отображения, указанный в файле > `displayLanguage` **Параметры > язык** в клиентском приложении Office. </span><span class="sxs-lookup"><span data-stu-id="b9d44-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d44-213">Тип</span><span class="sxs-lookup"><span data-stu-id="b9d44-213">Type</span></span>

*   <span data-ttu-id="b9d44-214">String</span><span class="sxs-lookup"><span data-stu-id="b9d44-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="b9d44-215">Требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-215">Requirements</span></span>

|<span data-ttu-id="b9d44-216">Требование</span><span class="sxs-lookup"><span data-stu-id="b9d44-216">Requirement</span></span>| <span data-ttu-id="b9d44-217">Значение</span><span class="sxs-lookup"><span data-stu-id="b9d44-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d44-218">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9d44-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9d44-219">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-219">1.1</span></span>|
|[<span data-ttu-id="b9d44-220">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9d44-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9d44-221">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d44-222">Пример</span><span class="sxs-lookup"><span data-stu-id="b9d44-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="b9d44-223">host: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="b9d44-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="b9d44-224">Получает Office приложение, в которое размещена надстройка.</span><span class="sxs-lookup"><span data-stu-id="b9d44-224">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d44-225">Кроме того, для получения [платформы можно использовать свойство Office.context.diagnostics.](#diagnostics-contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b9d44-225">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d44-226">Тип</span><span class="sxs-lookup"><span data-stu-id="b9d44-226">Type</span></span>

*   [<span data-ttu-id="b9d44-227">HostType</span><span class="sxs-lookup"><span data-stu-id="b9d44-227">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="b9d44-228">Требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-228">Requirements</span></span>

|<span data-ttu-id="b9d44-229">Требование</span><span class="sxs-lookup"><span data-stu-id="b9d44-229">Requirement</span></span>| <span data-ttu-id="b9d44-230">Значение</span><span class="sxs-lookup"><span data-stu-id="b9d44-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d44-231">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b9d44-231">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9d44-232">1.5</span><span class="sxs-lookup"><span data-stu-id="b9d44-232">1.5</span></span>|
|[<span data-ttu-id="b9d44-233">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9d44-233">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9d44-234">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-234">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d44-235">Пример</span><span class="sxs-lookup"><span data-stu-id="b9d44-235">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="b9d44-236">платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="b9d44-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="b9d44-237">Предоставляет платформу, на которой запущена надстройка.</span><span class="sxs-lookup"><span data-stu-id="b9d44-237">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="b9d44-238">Кроме того, для получения [платформы можно использовать свойство Office.context.diagnostics.](#diagnostics-contextinformation)</span><span class="sxs-lookup"><span data-stu-id="b9d44-238">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d44-239">Тип</span><span class="sxs-lookup"><span data-stu-id="b9d44-239">Type</span></span>

*   [<span data-ttu-id="b9d44-240">PlatformType</span><span class="sxs-lookup"><span data-stu-id="b9d44-240">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="b9d44-241">Требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-241">Requirements</span></span>

|<span data-ttu-id="b9d44-242">Требование</span><span class="sxs-lookup"><span data-stu-id="b9d44-242">Requirement</span></span>| <span data-ttu-id="b9d44-243">Значение</span><span class="sxs-lookup"><span data-stu-id="b9d44-243">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d44-244">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="b9d44-244">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9d44-245">1.5</span><span class="sxs-lookup"><span data-stu-id="b9d44-245">1.5</span></span>|
|[<span data-ttu-id="b9d44-246">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9d44-246">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9d44-247">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-247">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d44-248">Пример</span><span class="sxs-lookup"><span data-stu-id="b9d44-248">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="b9d44-249">требования: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="b9d44-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="b9d44-250">Предоставляет метод определения, какие наборы требований поддерживаются в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="b9d44-250">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d44-251">Тип</span><span class="sxs-lookup"><span data-stu-id="b9d44-251">Type</span></span>

*   [<span data-ttu-id="b9d44-252">RequirementSetSupport</span><span class="sxs-lookup"><span data-stu-id="b9d44-252">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="b9d44-253">Требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-253">Requirements</span></span>

|<span data-ttu-id="b9d44-254">Требование</span><span class="sxs-lookup"><span data-stu-id="b9d44-254">Requirement</span></span>| <span data-ttu-id="b9d44-255">Значение</span><span class="sxs-lookup"><span data-stu-id="b9d44-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d44-256">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9d44-256">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9d44-257">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-257">1.1</span></span>|
|[<span data-ttu-id="b9d44-258">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9d44-258">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9d44-259">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="b9d44-260">Пример</span><span class="sxs-lookup"><span data-stu-id="b9d44-260">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="b9d44-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="b9d44-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="b9d44-262">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="b9d44-262">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="b9d44-263">Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранимой в почтовом ящике пользователя, чтобы она была доступна этой надстройке, когда она запущена из любого клиента Outlook, используемого для доступа к этому `RoamingSettings` почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="b9d44-263">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d44-264">Тип</span><span class="sxs-lookup"><span data-stu-id="b9d44-264">Type</span></span>

*   [<span data-ttu-id="b9d44-265">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="b9d44-265">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="b9d44-266">Требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-266">Requirements</span></span>

|<span data-ttu-id="b9d44-267">Требование</span><span class="sxs-lookup"><span data-stu-id="b9d44-267">Requirement</span></span>| <span data-ttu-id="b9d44-268">Значение</span><span class="sxs-lookup"><span data-stu-id="b9d44-268">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d44-269">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9d44-269">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9d44-270">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-270">1.1</span></span>|
|[<span data-ttu-id="b9d44-271">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="b9d44-271">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="b9d44-272">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="b9d44-272">Restricted</span></span>|
|[<span data-ttu-id="b9d44-273">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9d44-273">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9d44-274">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-274">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="b9d44-275">ui: [пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="b9d44-275">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="b9d44-276">Предоставляет объекты и методы, которые можно использовать для создания и управления компонентами пользовательского интерфейса, такими как диалоговое окно, в Office надстройки.</span><span class="sxs-lookup"><span data-stu-id="b9d44-276">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="b9d44-277">Тип</span><span class="sxs-lookup"><span data-stu-id="b9d44-277">Type</span></span>

*   [<span data-ttu-id="b9d44-278">UI</span><span class="sxs-lookup"><span data-stu-id="b9d44-278">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="b9d44-279">Требования</span><span class="sxs-lookup"><span data-stu-id="b9d44-279">Requirements</span></span>

|<span data-ttu-id="b9d44-280">Требование</span><span class="sxs-lookup"><span data-stu-id="b9d44-280">Requirement</span></span>| <span data-ttu-id="b9d44-281">Значение</span><span class="sxs-lookup"><span data-stu-id="b9d44-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="b9d44-282">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="b9d44-282">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="b9d44-283">1.1</span><span class="sxs-lookup"><span data-stu-id="b9d44-283">1.1</span></span>|
|[<span data-ttu-id="b9d44-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="b9d44-284">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="b9d44-285">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="b9d44-285">Compose or Read</span></span>|
