---
title: Office. Context — набор обязательных элементов 1,9
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,9.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 6b2657d1e608bd1820d3814d9a6bfab67681824c
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/20/2020
ms.locfileid: "48628081"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="36134-103">контекст (набор требований для почтового ящика 1,9)</span><span class="sxs-lookup"><span data-stu-id="36134-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="36134-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="36134-104">[Office](office.md).context</span></span>

<span data-ttu-id="36134-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="36134-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="36134-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="36134-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="36134-107">Требования</span><span class="sxs-lookup"><span data-stu-id="36134-107">Requirements</span></span>

|<span data-ttu-id="36134-108">Требование</span><span class="sxs-lookup"><span data-stu-id="36134-108">Requirement</span></span>| <span data-ttu-id="36134-109">Значение</span><span class="sxs-lookup"><span data-stu-id="36134-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="36134-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="36134-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36134-111">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-111">1.1</span></span>|
|[<span data-ttu-id="36134-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="36134-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36134-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="36134-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="36134-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="36134-114">Properties</span></span>

| <span data-ttu-id="36134-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="36134-115">Property</span></span> | <span data-ttu-id="36134-116">Способов</span><span class="sxs-lookup"><span data-stu-id="36134-116">Modes</span></span> | <span data-ttu-id="36134-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="36134-117">Return type</span></span> | <span data-ttu-id="36134-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="36134-118">Minimum</span></span><br><span data-ttu-id="36134-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="36134-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="36134-120">auth</span><span class="sxs-lookup"><span data-stu-id="36134-120">auth</span></span>](#auth-auth) | <span data-ttu-id="36134-121">Создание</span><span class="sxs-lookup"><span data-stu-id="36134-121">Compose</span></span><br><span data-ttu-id="36134-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="36134-122">Read</span></span> | [<span data-ttu-id="36134-123">Auth</span><span class="sxs-lookup"><span data-stu-id="36134-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="36134-124">IdentityAPI 1,3</span><span class="sxs-lookup"><span data-stu-id="36134-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="36134-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="36134-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="36134-126">Создание</span><span class="sxs-lookup"><span data-stu-id="36134-126">Compose</span></span><br><span data-ttu-id="36134-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="36134-127">Read</span></span> | <span data-ttu-id="36134-128">String</span><span class="sxs-lookup"><span data-stu-id="36134-128">String</span></span> | [<span data-ttu-id="36134-129">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36134-130">диагностики</span><span class="sxs-lookup"><span data-stu-id="36134-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="36134-131">Создание</span><span class="sxs-lookup"><span data-stu-id="36134-131">Compose</span></span><br><span data-ttu-id="36134-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="36134-132">Read</span></span> | [<span data-ttu-id="36134-133">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="36134-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="36134-134">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36134-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="36134-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="36134-136">Создание</span><span class="sxs-lookup"><span data-stu-id="36134-136">Compose</span></span><br><span data-ttu-id="36134-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="36134-137">Read</span></span> | <span data-ttu-id="36134-138">String</span><span class="sxs-lookup"><span data-stu-id="36134-138">String</span></span> | [<span data-ttu-id="36134-139">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36134-140">размещать</span><span class="sxs-lookup"><span data-stu-id="36134-140">host</span></span>](#host-hosttype) | <span data-ttu-id="36134-141">Создание</span><span class="sxs-lookup"><span data-stu-id="36134-141">Compose</span></span><br><span data-ttu-id="36134-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="36134-142">Read</span></span> | [<span data-ttu-id="36134-143">HostType</span><span class="sxs-lookup"><span data-stu-id="36134-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="36134-144">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36134-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="36134-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="36134-146">Создание</span><span class="sxs-lookup"><span data-stu-id="36134-146">Compose</span></span><br><span data-ttu-id="36134-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="36134-147">Read</span></span> | [<span data-ttu-id="36134-148">Mailbox</span><span class="sxs-lookup"><span data-stu-id="36134-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="36134-149">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36134-150">платформа</span><span class="sxs-lookup"><span data-stu-id="36134-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="36134-151">Создание</span><span class="sxs-lookup"><span data-stu-id="36134-151">Compose</span></span><br><span data-ttu-id="36134-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="36134-152">Read</span></span> | [<span data-ttu-id="36134-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="36134-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="36134-154">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-154">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36134-155">потребность</span><span class="sxs-lookup"><span data-stu-id="36134-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="36134-156">Создание</span><span class="sxs-lookup"><span data-stu-id="36134-156">Compose</span></span><br><span data-ttu-id="36134-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="36134-157">Read</span></span> | [<span data-ttu-id="36134-158">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="36134-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="36134-159">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36134-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="36134-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="36134-161">Создание</span><span class="sxs-lookup"><span data-stu-id="36134-161">Compose</span></span><br><span data-ttu-id="36134-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="36134-162">Read</span></span> | [<span data-ttu-id="36134-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="36134-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="36134-164">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="36134-165">ui</span><span class="sxs-lookup"><span data-stu-id="36134-165">ui</span></span>](#ui-ui) | <span data-ttu-id="36134-166">Создание</span><span class="sxs-lookup"><span data-stu-id="36134-166">Compose</span></span><br><span data-ttu-id="36134-167">Чтение</span><span class="sxs-lookup"><span data-stu-id="36134-167">Read</span></span> | [<span data-ttu-id="36134-168">UI</span><span class="sxs-lookup"><span data-stu-id="36134-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="36134-169">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="36134-170">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="36134-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="36134-171">Проверка подлинности: [AUTH](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="36134-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="36134-172">Поддерживает [единый вход (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , предоставляя метод, позволяющий приложению Office получать маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="36134-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="36134-173">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="36134-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="36134-174">Просмотрите набор обязательных элементов [IdentityAPI 1,3](../../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="36134-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="36134-175">Тип</span><span class="sxs-lookup"><span data-stu-id="36134-175">Type</span></span>

*   [<span data-ttu-id="36134-176">Auth</span><span class="sxs-lookup"><span data-stu-id="36134-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="36134-177">Требования</span><span class="sxs-lookup"><span data-stu-id="36134-177">Requirements</span></span>

|<span data-ttu-id="36134-178">Требование</span><span class="sxs-lookup"><span data-stu-id="36134-178">Requirement</span></span>| <span data-ttu-id="36134-179">Значение</span><span class="sxs-lookup"><span data-stu-id="36134-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="36134-180">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="36134-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36134-181">Недоступно</span><span class="sxs-lookup"><span data-stu-id="36134-181">N/A</span></span>|
|[<span data-ttu-id="36134-182">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="36134-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36134-183">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="36134-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36134-184">Пример</span><span class="sxs-lookup"><span data-stu-id="36134-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="36134-185">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="36134-185">contentLanguage: String</span></span>

<span data-ttu-id="36134-186">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="36134-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="36134-187">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="36134-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="36134-188">Тип</span><span class="sxs-lookup"><span data-stu-id="36134-188">Type</span></span>

*   <span data-ttu-id="36134-189">String</span><span class="sxs-lookup"><span data-stu-id="36134-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="36134-190">Требования</span><span class="sxs-lookup"><span data-stu-id="36134-190">Requirements</span></span>

|<span data-ttu-id="36134-191">Требование</span><span class="sxs-lookup"><span data-stu-id="36134-191">Requirement</span></span>| <span data-ttu-id="36134-192">Значение</span><span class="sxs-lookup"><span data-stu-id="36134-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="36134-193">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="36134-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36134-194">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-194">1.1</span></span>|
|[<span data-ttu-id="36134-195">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="36134-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36134-196">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="36134-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36134-197">Пример</span><span class="sxs-lookup"><span data-stu-id="36134-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="36134-198">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="36134-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="36134-199">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="36134-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="36134-200">Тип</span><span class="sxs-lookup"><span data-stu-id="36134-200">Type</span></span>

*   [<span data-ttu-id="36134-201">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="36134-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="36134-202">Требования</span><span class="sxs-lookup"><span data-stu-id="36134-202">Requirements</span></span>

|<span data-ttu-id="36134-203">Требование</span><span class="sxs-lookup"><span data-stu-id="36134-203">Requirement</span></span>| <span data-ttu-id="36134-204">Значение</span><span class="sxs-lookup"><span data-stu-id="36134-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="36134-205">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="36134-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36134-206">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-206">1.1</span></span>|
|[<span data-ttu-id="36134-207">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="36134-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36134-208">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="36134-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36134-209">Пример</span><span class="sxs-lookup"><span data-stu-id="36134-209">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="36134-210">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="36134-210">displayLanguage: String</span></span>

<span data-ttu-id="36134-211">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="36134-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="36134-212">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="36134-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="36134-213">Тип</span><span class="sxs-lookup"><span data-stu-id="36134-213">Type</span></span>

*   <span data-ttu-id="36134-214">String</span><span class="sxs-lookup"><span data-stu-id="36134-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="36134-215">Требования</span><span class="sxs-lookup"><span data-stu-id="36134-215">Requirements</span></span>

|<span data-ttu-id="36134-216">Требование</span><span class="sxs-lookup"><span data-stu-id="36134-216">Requirement</span></span>| <span data-ttu-id="36134-217">Значение</span><span class="sxs-lookup"><span data-stu-id="36134-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="36134-218">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="36134-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36134-219">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-219">1.1</span></span>|
|[<span data-ttu-id="36134-220">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="36134-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36134-221">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="36134-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36134-222">Пример</span><span class="sxs-lookup"><span data-stu-id="36134-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="36134-223">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="36134-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="36134-224">Получает приложение Office, в котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="36134-224">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="36134-225">Тип</span><span class="sxs-lookup"><span data-stu-id="36134-225">Type</span></span>

*   [<span data-ttu-id="36134-226">HostType</span><span class="sxs-lookup"><span data-stu-id="36134-226">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="36134-227">Требования</span><span class="sxs-lookup"><span data-stu-id="36134-227">Requirements</span></span>

|<span data-ttu-id="36134-228">Требование</span><span class="sxs-lookup"><span data-stu-id="36134-228">Requirement</span></span>| <span data-ttu-id="36134-229">Значение</span><span class="sxs-lookup"><span data-stu-id="36134-229">Value</span></span>|
|---|---|
|[<span data-ttu-id="36134-230">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="36134-230">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36134-231">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-231">1.1</span></span>|
|[<span data-ttu-id="36134-232">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="36134-232">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36134-233">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="36134-233">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36134-234">Пример</span><span class="sxs-lookup"><span data-stu-id="36134-234">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="36134-235">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="36134-235">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="36134-236">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="36134-236">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="36134-237">Тип</span><span class="sxs-lookup"><span data-stu-id="36134-237">Type</span></span>

*   [<span data-ttu-id="36134-238">PlatformType</span><span class="sxs-lookup"><span data-stu-id="36134-238">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="36134-239">Требования</span><span class="sxs-lookup"><span data-stu-id="36134-239">Requirements</span></span>

|<span data-ttu-id="36134-240">Требование</span><span class="sxs-lookup"><span data-stu-id="36134-240">Requirement</span></span>| <span data-ttu-id="36134-241">Значение</span><span class="sxs-lookup"><span data-stu-id="36134-241">Value</span></span>|
|---|---|
|[<span data-ttu-id="36134-242">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="36134-242">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36134-243">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-243">1.1</span></span>|
|[<span data-ttu-id="36134-244">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="36134-244">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36134-245">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="36134-245">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36134-246">Пример</span><span class="sxs-lookup"><span data-stu-id="36134-246">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="36134-247">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="36134-247">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="36134-248">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="36134-248">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="36134-249">Тип</span><span class="sxs-lookup"><span data-stu-id="36134-249">Type</span></span>

*   [<span data-ttu-id="36134-250">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="36134-250">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="36134-251">Требования</span><span class="sxs-lookup"><span data-stu-id="36134-251">Requirements</span></span>

|<span data-ttu-id="36134-252">Требование</span><span class="sxs-lookup"><span data-stu-id="36134-252">Requirement</span></span>| <span data-ttu-id="36134-253">Значение</span><span class="sxs-lookup"><span data-stu-id="36134-253">Value</span></span>|
|---|---|
|[<span data-ttu-id="36134-254">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="36134-254">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36134-255">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-255">1.1</span></span>|
|[<span data-ttu-id="36134-256">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="36134-256">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36134-257">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="36134-257">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="36134-258">Пример</span><span class="sxs-lookup"><span data-stu-id="36134-258">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="36134-259">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="36134-259">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="36134-260">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="36134-260">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="36134-261">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="36134-261">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="36134-262">Тип</span><span class="sxs-lookup"><span data-stu-id="36134-262">Type</span></span>

*   [<span data-ttu-id="36134-263">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="36134-263">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="36134-264">Требования</span><span class="sxs-lookup"><span data-stu-id="36134-264">Requirements</span></span>

|<span data-ttu-id="36134-265">Требование</span><span class="sxs-lookup"><span data-stu-id="36134-265">Requirement</span></span>| <span data-ttu-id="36134-266">Значение</span><span class="sxs-lookup"><span data-stu-id="36134-266">Value</span></span>|
|---|---|
|[<span data-ttu-id="36134-267">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="36134-267">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36134-268">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-268">1.1</span></span>|
|[<span data-ttu-id="36134-269">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="36134-269">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="36134-270">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="36134-270">Restricted</span></span>|
|[<span data-ttu-id="36134-271">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="36134-271">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36134-272">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="36134-272">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="36134-273">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="36134-273">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="36134-274">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="36134-274">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="36134-275">Тип</span><span class="sxs-lookup"><span data-stu-id="36134-275">Type</span></span>

*   [<span data-ttu-id="36134-276">UI</span><span class="sxs-lookup"><span data-stu-id="36134-276">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="36134-277">Требования</span><span class="sxs-lookup"><span data-stu-id="36134-277">Requirements</span></span>

|<span data-ttu-id="36134-278">Требование</span><span class="sxs-lookup"><span data-stu-id="36134-278">Requirement</span></span>| <span data-ttu-id="36134-279">Значение</span><span class="sxs-lookup"><span data-stu-id="36134-279">Value</span></span>|
|---|---|
|[<span data-ttu-id="36134-280">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="36134-280">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="36134-281">1.1</span><span class="sxs-lookup"><span data-stu-id="36134-281">1.1</span></span>|
|[<span data-ttu-id="36134-282">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="36134-282">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="36134-283">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="36134-283">Compose or Read</span></span>|
