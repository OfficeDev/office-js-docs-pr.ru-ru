---
title: Office.context — предварительная версия набора обязательных элементов
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора обязательных элементов API почтового ящика.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 8370df907aa3ab0534254057860c187cec583e6c
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570788"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="01abf-103">контекст (набор требований Preview для предварительного просмотра почтового ящика)</span><span class="sxs-lookup"><span data-stu-id="01abf-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="01abf-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="01abf-104">[Office](office.md).context</span></span>

<span data-ttu-id="01abf-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="01abf-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="01abf-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="01abf-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="01abf-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="01abf-107">Requirements</span></span>

|<span data-ttu-id="01abf-108">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-108">Requirement</span></span>| <span data-ttu-id="01abf-109">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="01abf-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-111">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-111">1.1</span></span>|
|[<span data-ttu-id="01abf-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="01abf-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="01abf-114">Properties</span></span>

| <span data-ttu-id="01abf-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="01abf-115">Property</span></span> | <span data-ttu-id="01abf-116">Способов</span><span class="sxs-lookup"><span data-stu-id="01abf-116">Modes</span></span> | <span data-ttu-id="01abf-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="01abf-117">Return type</span></span> | <span data-ttu-id="01abf-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="01abf-118">Minimum</span></span><br><span data-ttu-id="01abf-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="01abf-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="01abf-120">auth</span><span class="sxs-lookup"><span data-stu-id="01abf-120">auth</span></span>](#auth-auth) | <span data-ttu-id="01abf-121">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-121">Compose</span></span><br><span data-ttu-id="01abf-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-122">Read</span></span> | [<span data-ttu-id="01abf-123">Auth</span><span class="sxs-lookup"><span data-stu-id="01abf-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="01abf-124">IdentityAPI 1,3</span><span class="sxs-lookup"><span data-stu-id="01abf-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="01abf-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="01abf-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="01abf-126">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-126">Compose</span></span><br><span data-ttu-id="01abf-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-127">Read</span></span> | <span data-ttu-id="01abf-128">String</span><span class="sxs-lookup"><span data-stu-id="01abf-128">String</span></span> | [<span data-ttu-id="01abf-129">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="01abf-130">диагностики</span><span class="sxs-lookup"><span data-stu-id="01abf-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="01abf-131">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-131">Compose</span></span><br><span data-ttu-id="01abf-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-132">Read</span></span> | [<span data-ttu-id="01abf-133">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="01abf-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="01abf-134">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="01abf-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="01abf-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="01abf-136">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-136">Compose</span></span><br><span data-ttu-id="01abf-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-137">Read</span></span> | <span data-ttu-id="01abf-138">String</span><span class="sxs-lookup"><span data-stu-id="01abf-138">String</span></span> | [<span data-ttu-id="01abf-139">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="01abf-140">размещать</span><span class="sxs-lookup"><span data-stu-id="01abf-140">host</span></span>](#host-hosttype) | <span data-ttu-id="01abf-141">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-141">Compose</span></span><br><span data-ttu-id="01abf-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-142">Read</span></span> | [<span data-ttu-id="01abf-143">HostType</span><span class="sxs-lookup"><span data-stu-id="01abf-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="01abf-144">1,5</span><span class="sxs-lookup"><span data-stu-id="01abf-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="01abf-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="01abf-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="01abf-146">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-146">Compose</span></span><br><span data-ttu-id="01abf-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-147">Read</span></span> | [<span data-ttu-id="01abf-148">Mailbox</span><span class="sxs-lookup"><span data-stu-id="01abf-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="01abf-149">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="01abf-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="01abf-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="01abf-151">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-151">Compose</span></span><br><span data-ttu-id="01abf-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-152">Read</span></span> | [<span data-ttu-id="01abf-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="01abf-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="01abf-154">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="01abf-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="01abf-155">платформа</span><span class="sxs-lookup"><span data-stu-id="01abf-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="01abf-156">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-156">Compose</span></span><br><span data-ttu-id="01abf-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-157">Read</span></span> | [<span data-ttu-id="01abf-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="01abf-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="01abf-159">1,5</span><span class="sxs-lookup"><span data-stu-id="01abf-159">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="01abf-160">потребность</span><span class="sxs-lookup"><span data-stu-id="01abf-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="01abf-161">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-161">Compose</span></span><br><span data-ttu-id="01abf-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-162">Read</span></span> | [<span data-ttu-id="01abf-163">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="01abf-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="01abf-164">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="01abf-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="01abf-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="01abf-166">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-166">Compose</span></span><br><span data-ttu-id="01abf-167">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-167">Read</span></span> | [<span data-ttu-id="01abf-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="01abf-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="01abf-169">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="01abf-170">ui</span><span class="sxs-lookup"><span data-stu-id="01abf-170">ui</span></span>](#ui-ui) | <span data-ttu-id="01abf-171">Создание</span><span class="sxs-lookup"><span data-stu-id="01abf-171">Compose</span></span><br><span data-ttu-id="01abf-172">Чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-172">Read</span></span> | [<span data-ttu-id="01abf-173">UI</span><span class="sxs-lookup"><span data-stu-id="01abf-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="01abf-174">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="01abf-175">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="01abf-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="01abf-176">Проверка подлинности: [AUTH](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="01abf-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="01abf-177">Поддерживает [единый вход (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , предоставляя метод, позволяющий приложению Office получать маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="01abf-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="01abf-178">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="01abf-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="01abf-179">Type</span><span class="sxs-lookup"><span data-stu-id="01abf-179">Type</span></span>

*   [<span data-ttu-id="01abf-180">Auth</span><span class="sxs-lookup"><span data-stu-id="01abf-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="01abf-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="01abf-181">Requirements</span></span>

|<span data-ttu-id="01abf-182">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-182">Requirement</span></span>| <span data-ttu-id="01abf-183">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="01abf-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-185">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="01abf-185">Preview</span></span>|
|[<span data-ttu-id="01abf-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01abf-188">Пример</span><span class="sxs-lookup"><span data-stu-id="01abf-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="01abf-189">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="01abf-189">contentLanguage: String</span></span>

<span data-ttu-id="01abf-190">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="01abf-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="01abf-191">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="01abf-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="01abf-192">Тип</span><span class="sxs-lookup"><span data-stu-id="01abf-192">Type</span></span>

*   <span data-ttu-id="01abf-193">String</span><span class="sxs-lookup"><span data-stu-id="01abf-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="01abf-194">Требования</span><span class="sxs-lookup"><span data-stu-id="01abf-194">Requirements</span></span>

|<span data-ttu-id="01abf-195">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-195">Requirement</span></span>| <span data-ttu-id="01abf-196">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="01abf-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-198">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-198">1.1</span></span>|
|[<span data-ttu-id="01abf-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01abf-201">Пример</span><span class="sxs-lookup"><span data-stu-id="01abf-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="01abf-202">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="01abf-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="01abf-203">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="01abf-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="01abf-204">Type</span><span class="sxs-lookup"><span data-stu-id="01abf-204">Type</span></span>

*   [<span data-ttu-id="01abf-205">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="01abf-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="01abf-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="01abf-206">Requirements</span></span>

|<span data-ttu-id="01abf-207">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-207">Requirement</span></span>| <span data-ttu-id="01abf-208">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-209">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="01abf-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-210">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-210">1.1</span></span>|
|[<span data-ttu-id="01abf-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-212">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01abf-213">Пример</span><span class="sxs-lookup"><span data-stu-id="01abf-213">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="01abf-214">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="01abf-214">displayLanguage: String</span></span>

<span data-ttu-id="01abf-215">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="01abf-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="01abf-216">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="01abf-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="01abf-217">Тип</span><span class="sxs-lookup"><span data-stu-id="01abf-217">Type</span></span>

*   <span data-ttu-id="01abf-218">String</span><span class="sxs-lookup"><span data-stu-id="01abf-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="01abf-219">Требования</span><span class="sxs-lookup"><span data-stu-id="01abf-219">Requirements</span></span>

|<span data-ttu-id="01abf-220">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-220">Requirement</span></span>| <span data-ttu-id="01abf-221">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-222">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="01abf-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-223">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-223">1.1</span></span>|
|[<span data-ttu-id="01abf-224">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-225">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01abf-226">Пример</span><span class="sxs-lookup"><span data-stu-id="01abf-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="01abf-227">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="01abf-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="01abf-228">Получает приложение Office, в котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="01abf-228">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="01abf-229">Кроме того, вы можете использовать свойство [Office. Context. Diagnostics](#diagnostics-contextinformation) для получения ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="01abf-229">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the host.</span></span>

##### <a name="type"></a><span data-ttu-id="01abf-230">Type</span><span class="sxs-lookup"><span data-stu-id="01abf-230">Type</span></span>

*   [<span data-ttu-id="01abf-231">HostType</span><span class="sxs-lookup"><span data-stu-id="01abf-231">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="01abf-232">Requirements</span><span class="sxs-lookup"><span data-stu-id="01abf-232">Requirements</span></span>

|<span data-ttu-id="01abf-233">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-233">Requirement</span></span>| <span data-ttu-id="01abf-234">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-234">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-235">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="01abf-235">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-236">1.5</span><span class="sxs-lookup"><span data-stu-id="01abf-236">1.5</span></span>|
|[<span data-ttu-id="01abf-237">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-237">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-238">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-238">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01abf-239">Пример</span><span class="sxs-lookup"><span data-stu-id="01abf-239">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="01abf-240">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="01abf-240">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="01abf-241">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="01abf-241">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="01abf-242">Этот элемент поддерживается только в Outlook для Windows.</span><span class="sxs-lookup"><span data-stu-id="01abf-242">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="01abf-243">Использование цветов тем Office позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с помощью **файла > учетной записи office > пользовательского интерфейса темы** Office, которая применяется ко всем клиентским приложениям Office.</span><span class="sxs-lookup"><span data-stu-id="01abf-243">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="01abf-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="01abf-244">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="01abf-245">Type</span><span class="sxs-lookup"><span data-stu-id="01abf-245">Type</span></span>

*   [<span data-ttu-id="01abf-246">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="01abf-246">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="01abf-247">Свойства:</span><span class="sxs-lookup"><span data-stu-id="01abf-247">Properties:</span></span>

|<span data-ttu-id="01abf-248">Имя</span><span class="sxs-lookup"><span data-stu-id="01abf-248">Name</span></span>| <span data-ttu-id="01abf-249">Тип</span><span class="sxs-lookup"><span data-stu-id="01abf-249">Type</span></span>| <span data-ttu-id="01abf-250">Описание</span><span class="sxs-lookup"><span data-stu-id="01abf-250">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="01abf-251">String</span><span class="sxs-lookup"><span data-stu-id="01abf-251">String</span></span>|<span data-ttu-id="01abf-252">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="01abf-252">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="01abf-253">String</span><span class="sxs-lookup"><span data-stu-id="01abf-253">String</span></span>|<span data-ttu-id="01abf-254">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="01abf-254">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="01abf-255">String</span><span class="sxs-lookup"><span data-stu-id="01abf-255">String</span></span>|<span data-ttu-id="01abf-256">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="01abf-256">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="01abf-257">String</span><span class="sxs-lookup"><span data-stu-id="01abf-257">String</span></span>|<span data-ttu-id="01abf-258">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="01abf-258">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="01abf-259">Requirements</span><span class="sxs-lookup"><span data-stu-id="01abf-259">Requirements</span></span>

|<span data-ttu-id="01abf-260">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-260">Requirement</span></span>| <span data-ttu-id="01abf-261">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-261">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-262">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="01abf-262">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-263">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="01abf-263">Preview</span></span>|
|[<span data-ttu-id="01abf-264">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-264">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-265">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-265">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01abf-266">Пример</span><span class="sxs-lookup"><span data-stu-id="01abf-266">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="01abf-267">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="01abf-267">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="01abf-268">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="01abf-268">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="01abf-269">Кроме того, вы можете использовать свойство [Office. Context. Diagnostics](#diagnostics-contextinformation) для получения платформы.</span><span class="sxs-lookup"><span data-stu-id="01abf-269">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="01abf-270">Type</span><span class="sxs-lookup"><span data-stu-id="01abf-270">Type</span></span>

*   [<span data-ttu-id="01abf-271">PlatformType</span><span class="sxs-lookup"><span data-stu-id="01abf-271">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="01abf-272">Requirements</span><span class="sxs-lookup"><span data-stu-id="01abf-272">Requirements</span></span>

|<span data-ttu-id="01abf-273">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-273">Requirement</span></span>| <span data-ttu-id="01abf-274">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-274">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-275">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="01abf-275">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-276">1.5</span><span class="sxs-lookup"><span data-stu-id="01abf-276">1.5</span></span>|
|[<span data-ttu-id="01abf-277">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-277">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-278">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-278">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01abf-279">Пример</span><span class="sxs-lookup"><span data-stu-id="01abf-279">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="01abf-280">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="01abf-280">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="01abf-281">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="01abf-281">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="01abf-282">Type</span><span class="sxs-lookup"><span data-stu-id="01abf-282">Type</span></span>

*   [<span data-ttu-id="01abf-283">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="01abf-283">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="01abf-284">Requirements</span><span class="sxs-lookup"><span data-stu-id="01abf-284">Requirements</span></span>

|<span data-ttu-id="01abf-285">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-285">Requirement</span></span>| <span data-ttu-id="01abf-286">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-286">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-287">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="01abf-287">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-288">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-288">1.1</span></span>|
|[<span data-ttu-id="01abf-289">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-289">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-290">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-290">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="01abf-291">Пример</span><span class="sxs-lookup"><span data-stu-id="01abf-291">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="01abf-292">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="01abf-292">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="01abf-293">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="01abf-293">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="01abf-294">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="01abf-294">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="01abf-295">Type</span><span class="sxs-lookup"><span data-stu-id="01abf-295">Type</span></span>

*   [<span data-ttu-id="01abf-296">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="01abf-296">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="01abf-297">Requirements</span><span class="sxs-lookup"><span data-stu-id="01abf-297">Requirements</span></span>

|<span data-ttu-id="01abf-298">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-298">Requirement</span></span>| <span data-ttu-id="01abf-299">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-299">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-300">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="01abf-300">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-301">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-301">1.1</span></span>|
|[<span data-ttu-id="01abf-302">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="01abf-302">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="01abf-303">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="01abf-303">Restricted</span></span>|
|[<span data-ttu-id="01abf-304">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-304">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-305">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-305">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="01abf-306">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="01abf-306">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="01abf-307">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="01abf-307">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="01abf-308">Type</span><span class="sxs-lookup"><span data-stu-id="01abf-308">Type</span></span>

*   [<span data-ttu-id="01abf-309">UI</span><span class="sxs-lookup"><span data-stu-id="01abf-309">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="01abf-310">Requirements</span><span class="sxs-lookup"><span data-stu-id="01abf-310">Requirements</span></span>

|<span data-ttu-id="01abf-311">Требование</span><span class="sxs-lookup"><span data-stu-id="01abf-311">Requirement</span></span>| <span data-ttu-id="01abf-312">Значение</span><span class="sxs-lookup"><span data-stu-id="01abf-312">Value</span></span>|
|---|---|
|[<span data-ttu-id="01abf-313">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="01abf-313">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="01abf-314">1.1</span><span class="sxs-lookup"><span data-stu-id="01abf-314">1.1</span></span>|
|[<span data-ttu-id="01abf-315">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="01abf-315">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="01abf-316">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="01abf-316">Compose or Read</span></span>|
