---
title: Office. Context — набор обязательных элементов 1,9
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора требований API почтовых ящиков 1,9.
ms.date: 12/03/2020
localization_priority: Normal
ms.openlocfilehash: 3a8a9fe65ebf3c5a5ee63766f71dfce8e3f8d905
ms.sourcegitcommit: 1737026df569b62957d38b62c0b16caee4f0cdfe
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/04/2020
ms.locfileid: "49570725"
---
# <a name="context-mailbox-requirement-set-19"></a><span data-ttu-id="e23ef-103">контекст (набор требований для почтового ящика 1,9)</span><span class="sxs-lookup"><span data-stu-id="e23ef-103">context (Mailbox requirement set 1.9)</span></span>

### <a name="officecontext"></a><span data-ttu-id="e23ef-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="e23ef-104">[Office](office.md).context</span></span>

<span data-ttu-id="e23ef-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="e23ef-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="e23ef-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="e23ef-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-1.9&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="e23ef-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="e23ef-107">Requirements</span></span>

|<span data-ttu-id="e23ef-108">Требование</span><span class="sxs-lookup"><span data-stu-id="e23ef-108">Requirement</span></span>| <span data-ttu-id="e23ef-109">Значение</span><span class="sxs-lookup"><span data-stu-id="e23ef-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="e23ef-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e23ef-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e23ef-111">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-111">1.1</span></span>|
|[<span data-ttu-id="e23ef-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e23ef-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e23ef-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="e23ef-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="e23ef-114">Properties</span></span>

| <span data-ttu-id="e23ef-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="e23ef-115">Property</span></span> | <span data-ttu-id="e23ef-116">Способов</span><span class="sxs-lookup"><span data-stu-id="e23ef-116">Modes</span></span> | <span data-ttu-id="e23ef-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="e23ef-117">Return type</span></span> | <span data-ttu-id="e23ef-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="e23ef-118">Minimum</span></span><br><span data-ttu-id="e23ef-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="e23ef-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="e23ef-120">auth</span><span class="sxs-lookup"><span data-stu-id="e23ef-120">auth</span></span>](#auth-auth) | <span data-ttu-id="e23ef-121">Создание</span><span class="sxs-lookup"><span data-stu-id="e23ef-121">Compose</span></span><br><span data-ttu-id="e23ef-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-122">Read</span></span> | [<span data-ttu-id="e23ef-123">Auth</span><span class="sxs-lookup"><span data-stu-id="e23ef-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="e23ef-124">IdentityAPI 1,3</span><span class="sxs-lookup"><span data-stu-id="e23ef-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="e23ef-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="e23ef-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="e23ef-126">Создание</span><span class="sxs-lookup"><span data-stu-id="e23ef-126">Compose</span></span><br><span data-ttu-id="e23ef-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-127">Read</span></span> | <span data-ttu-id="e23ef-128">String</span><span class="sxs-lookup"><span data-stu-id="e23ef-128">String</span></span> | [<span data-ttu-id="e23ef-129">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e23ef-130">диагностики</span><span class="sxs-lookup"><span data-stu-id="e23ef-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="e23ef-131">Создание</span><span class="sxs-lookup"><span data-stu-id="e23ef-131">Compose</span></span><br><span data-ttu-id="e23ef-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-132">Read</span></span> | [<span data-ttu-id="e23ef-133">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="e23ef-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="e23ef-134">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e23ef-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="e23ef-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="e23ef-136">Создание</span><span class="sxs-lookup"><span data-stu-id="e23ef-136">Compose</span></span><br><span data-ttu-id="e23ef-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-137">Read</span></span> | <span data-ttu-id="e23ef-138">String</span><span class="sxs-lookup"><span data-stu-id="e23ef-138">String</span></span> | [<span data-ttu-id="e23ef-139">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e23ef-140">размещать</span><span class="sxs-lookup"><span data-stu-id="e23ef-140">host</span></span>](#host-hosttype) | <span data-ttu-id="e23ef-141">Создание</span><span class="sxs-lookup"><span data-stu-id="e23ef-141">Compose</span></span><br><span data-ttu-id="e23ef-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-142">Read</span></span> | [<span data-ttu-id="e23ef-143">HostType</span><span class="sxs-lookup"><span data-stu-id="e23ef-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="e23ef-144">1,5</span><span class="sxs-lookup"><span data-stu-id="e23ef-144">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e23ef-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="e23ef-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="e23ef-146">Создание</span><span class="sxs-lookup"><span data-stu-id="e23ef-146">Compose</span></span><br><span data-ttu-id="e23ef-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-147">Read</span></span> | [<span data-ttu-id="e23ef-148">Mailbox</span><span class="sxs-lookup"><span data-stu-id="e23ef-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="e23ef-149">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e23ef-150">платформа</span><span class="sxs-lookup"><span data-stu-id="e23ef-150">platform</span></span>](#platform-platformtype) | <span data-ttu-id="e23ef-151">Создание</span><span class="sxs-lookup"><span data-stu-id="e23ef-151">Compose</span></span><br><span data-ttu-id="e23ef-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-152">Read</span></span> | [<span data-ttu-id="e23ef-153">PlatformType</span><span class="sxs-lookup"><span data-stu-id="e23ef-153">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="e23ef-154">1,5</span><span class="sxs-lookup"><span data-stu-id="e23ef-154">1.5</span></span>](../requirement-set-1.5/outlook-requirement-set-1.5.md) |
| [<span data-ttu-id="e23ef-155">потребность</span><span class="sxs-lookup"><span data-stu-id="e23ef-155">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="e23ef-156">Создание</span><span class="sxs-lookup"><span data-stu-id="e23ef-156">Compose</span></span><br><span data-ttu-id="e23ef-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-157">Read</span></span> | [<span data-ttu-id="e23ef-158">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="e23ef-158">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="e23ef-159">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e23ef-160">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="e23ef-160">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="e23ef-161">Создание</span><span class="sxs-lookup"><span data-stu-id="e23ef-161">Compose</span></span><br><span data-ttu-id="e23ef-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-162">Read</span></span> | [<span data-ttu-id="e23ef-163">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e23ef-163">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="e23ef-164">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="e23ef-165">ui</span><span class="sxs-lookup"><span data-stu-id="e23ef-165">ui</span></span>](#ui-ui) | <span data-ttu-id="e23ef-166">Создание</span><span class="sxs-lookup"><span data-stu-id="e23ef-166">Compose</span></span><br><span data-ttu-id="e23ef-167">Чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-167">Read</span></span> | [<span data-ttu-id="e23ef-168">UI</span><span class="sxs-lookup"><span data-stu-id="e23ef-168">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-1.9&preserve-view=true) | [<span data-ttu-id="e23ef-169">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="e23ef-170">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="e23ef-170">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="e23ef-171">Проверка подлинности: [AUTH](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="e23ef-171">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="e23ef-172">Поддерживает [единый вход (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , предоставляя метод, позволяющий приложению Office получать маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="e23ef-172">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="e23ef-173">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="e23ef-173">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span> <span data-ttu-id="e23ef-174">Просмотрите набор обязательных элементов [IdentityAPI 1,3](../../requirement-sets/identity-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="e23ef-174">See [IdentityAPI 1.3 requirement set](../../requirement-sets/identity-api-requirement-sets.md).</span></span>

##### <a name="type"></a><span data-ttu-id="e23ef-175">Type</span><span class="sxs-lookup"><span data-stu-id="e23ef-175">Type</span></span>

*   [<span data-ttu-id="e23ef-176">Auth</span><span class="sxs-lookup"><span data-stu-id="e23ef-176">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="e23ef-177">Requirements</span><span class="sxs-lookup"><span data-stu-id="e23ef-177">Requirements</span></span>

|<span data-ttu-id="e23ef-178">Требование</span><span class="sxs-lookup"><span data-stu-id="e23ef-178">Requirement</span></span>| <span data-ttu-id="e23ef-179">Значение</span><span class="sxs-lookup"><span data-stu-id="e23ef-179">Value</span></span>|
|---|---|
|[<span data-ttu-id="e23ef-180">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e23ef-180">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e23ef-181">Недоступно</span><span class="sxs-lookup"><span data-stu-id="e23ef-181">N/A</span></span>|
|[<span data-ttu-id="e23ef-182">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e23ef-182">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e23ef-183">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-183">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e23ef-184">Пример</span><span class="sxs-lookup"><span data-stu-id="e23ef-184">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="e23ef-185">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="e23ef-185">contentLanguage: String</span></span>

<span data-ttu-id="e23ef-186">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="e23ef-186">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="e23ef-187">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="e23ef-187">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="e23ef-188">Тип</span><span class="sxs-lookup"><span data-stu-id="e23ef-188">Type</span></span>

*   <span data-ttu-id="e23ef-189">String</span><span class="sxs-lookup"><span data-stu-id="e23ef-189">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e23ef-190">Требования</span><span class="sxs-lookup"><span data-stu-id="e23ef-190">Requirements</span></span>

|<span data-ttu-id="e23ef-191">Требование</span><span class="sxs-lookup"><span data-stu-id="e23ef-191">Requirement</span></span>| <span data-ttu-id="e23ef-192">Значение</span><span class="sxs-lookup"><span data-stu-id="e23ef-192">Value</span></span>|
|---|---|
|[<span data-ttu-id="e23ef-193">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e23ef-193">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e23ef-194">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-194">1.1</span></span>|
|[<span data-ttu-id="e23ef-195">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e23ef-195">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e23ef-196">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-196">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e23ef-197">Пример</span><span class="sxs-lookup"><span data-stu-id="e23ef-197">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="e23ef-198">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="e23ef-198">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="e23ef-199">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="e23ef-199">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="e23ef-200">Type</span><span class="sxs-lookup"><span data-stu-id="e23ef-200">Type</span></span>

*   [<span data-ttu-id="e23ef-201">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="e23ef-201">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="e23ef-202">Requirements</span><span class="sxs-lookup"><span data-stu-id="e23ef-202">Requirements</span></span>

|<span data-ttu-id="e23ef-203">Требование</span><span class="sxs-lookup"><span data-stu-id="e23ef-203">Requirement</span></span>| <span data-ttu-id="e23ef-204">Значение</span><span class="sxs-lookup"><span data-stu-id="e23ef-204">Value</span></span>|
|---|---|
|[<span data-ttu-id="e23ef-205">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e23ef-205">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e23ef-206">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-206">1.1</span></span>|
|[<span data-ttu-id="e23ef-207">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e23ef-207">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e23ef-208">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e23ef-209">Пример</span><span class="sxs-lookup"><span data-stu-id="e23ef-209">Example</span></span>

```js
var contextInfo = Office.context.diagnostics;
console.log("Office application: " + contextInfo.host);
console.log("Office version: " + contextInfo.version);
console.log("Platform: " + contextInfo.platform);
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="e23ef-210">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="e23ef-210">displayLanguage: String</span></span>

<span data-ttu-id="e23ef-211">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="e23ef-211">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="e23ef-212">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="e23ef-212">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="e23ef-213">Тип</span><span class="sxs-lookup"><span data-stu-id="e23ef-213">Type</span></span>

*   <span data-ttu-id="e23ef-214">String</span><span class="sxs-lookup"><span data-stu-id="e23ef-214">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="e23ef-215">Требования</span><span class="sxs-lookup"><span data-stu-id="e23ef-215">Requirements</span></span>

|<span data-ttu-id="e23ef-216">Требование</span><span class="sxs-lookup"><span data-stu-id="e23ef-216">Requirement</span></span>| <span data-ttu-id="e23ef-217">Значение</span><span class="sxs-lookup"><span data-stu-id="e23ef-217">Value</span></span>|
|---|---|
|[<span data-ttu-id="e23ef-218">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e23ef-218">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e23ef-219">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-219">1.1</span></span>|
|[<span data-ttu-id="e23ef-220">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e23ef-220">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e23ef-221">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-221">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e23ef-222">Пример</span><span class="sxs-lookup"><span data-stu-id="e23ef-222">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="e23ef-223">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="e23ef-223">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="e23ef-224">Получает приложение Office, в котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="e23ef-224">Gets the Office application that is hosting the add-in.</span></span>

> [!NOTE]
> <span data-ttu-id="e23ef-225">Кроме того, вы можете использовать свойство [Office. Context. Diagnostics](#diagnostics-contextinformation) для получения платформы.</span><span class="sxs-lookup"><span data-stu-id="e23ef-225">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="e23ef-226">Type</span><span class="sxs-lookup"><span data-stu-id="e23ef-226">Type</span></span>

*   [<span data-ttu-id="e23ef-227">HostType</span><span class="sxs-lookup"><span data-stu-id="e23ef-227">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="e23ef-228">Requirements</span><span class="sxs-lookup"><span data-stu-id="e23ef-228">Requirements</span></span>

|<span data-ttu-id="e23ef-229">Требование</span><span class="sxs-lookup"><span data-stu-id="e23ef-229">Requirement</span></span>| <span data-ttu-id="e23ef-230">Значение</span><span class="sxs-lookup"><span data-stu-id="e23ef-230">Value</span></span>|
|---|---|
|[<span data-ttu-id="e23ef-231">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e23ef-231">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e23ef-232">1.5</span><span class="sxs-lookup"><span data-stu-id="e23ef-232">1.5</span></span>|
|[<span data-ttu-id="e23ef-233">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e23ef-233">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e23ef-234">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-234">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e23ef-235">Пример</span><span class="sxs-lookup"><span data-stu-id="e23ef-235">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="platform-platformtype"></a><span data-ttu-id="e23ef-236">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="e23ef-236">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="e23ef-237">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="e23ef-237">Provides the platform on which the add-in is running.</span></span>

> [!NOTE]
> <span data-ttu-id="e23ef-238">Кроме того, вы можете использовать свойство [Office. Context. Diagnostics](#diagnostics-contextinformation) для получения платформы.</span><span class="sxs-lookup"><span data-stu-id="e23ef-238">Alternatively, you can use the [Office.context.diagnostics](#diagnostics-contextinformation) property to get the platform.</span></span>

##### <a name="type"></a><span data-ttu-id="e23ef-239">Type</span><span class="sxs-lookup"><span data-stu-id="e23ef-239">Type</span></span>

*   [<span data-ttu-id="e23ef-240">PlatformType</span><span class="sxs-lookup"><span data-stu-id="e23ef-240">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="e23ef-241">Requirements</span><span class="sxs-lookup"><span data-stu-id="e23ef-241">Requirements</span></span>

|<span data-ttu-id="e23ef-242">Требование</span><span class="sxs-lookup"><span data-stu-id="e23ef-242">Requirement</span></span>| <span data-ttu-id="e23ef-243">Значение</span><span class="sxs-lookup"><span data-stu-id="e23ef-243">Value</span></span>|
|---|---|
|[<span data-ttu-id="e23ef-244">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="e23ef-244">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e23ef-245">1.5</span><span class="sxs-lookup"><span data-stu-id="e23ef-245">1.5</span></span>|
|[<span data-ttu-id="e23ef-246">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e23ef-246">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e23ef-247">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-247">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e23ef-248">Пример</span><span class="sxs-lookup"><span data-stu-id="e23ef-248">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="e23ef-249">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="e23ef-249">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="e23ef-250">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="e23ef-250">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="e23ef-251">Type</span><span class="sxs-lookup"><span data-stu-id="e23ef-251">Type</span></span>

*   [<span data-ttu-id="e23ef-252">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="e23ef-252">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="e23ef-253">Requirements</span><span class="sxs-lookup"><span data-stu-id="e23ef-253">Requirements</span></span>

|<span data-ttu-id="e23ef-254">Требование</span><span class="sxs-lookup"><span data-stu-id="e23ef-254">Requirement</span></span>| <span data-ttu-id="e23ef-255">Значение</span><span class="sxs-lookup"><span data-stu-id="e23ef-255">Value</span></span>|
|---|---|
|[<span data-ttu-id="e23ef-256">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e23ef-256">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e23ef-257">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-257">1.1</span></span>|
|[<span data-ttu-id="e23ef-258">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e23ef-258">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e23ef-259">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-259">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="e23ef-260">Пример</span><span class="sxs-lookup"><span data-stu-id="e23ef-260">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="e23ef-261">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="e23ef-261">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="e23ef-262">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="e23ef-262">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="e23ef-263">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="e23ef-263">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="e23ef-264">Type</span><span class="sxs-lookup"><span data-stu-id="e23ef-264">Type</span></span>

*   [<span data-ttu-id="e23ef-265">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="e23ef-265">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="e23ef-266">Requirements</span><span class="sxs-lookup"><span data-stu-id="e23ef-266">Requirements</span></span>

|<span data-ttu-id="e23ef-267">Требование</span><span class="sxs-lookup"><span data-stu-id="e23ef-267">Requirement</span></span>| <span data-ttu-id="e23ef-268">Значение</span><span class="sxs-lookup"><span data-stu-id="e23ef-268">Value</span></span>|
|---|---|
|[<span data-ttu-id="e23ef-269">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e23ef-269">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e23ef-270">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-270">1.1</span></span>|
|[<span data-ttu-id="e23ef-271">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="e23ef-271">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="e23ef-272">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="e23ef-272">Restricted</span></span>|
|[<span data-ttu-id="e23ef-273">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e23ef-273">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e23ef-274">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-274">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="e23ef-275">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="e23ef-275">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="e23ef-276">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="e23ef-276">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="e23ef-277">Type</span><span class="sxs-lookup"><span data-stu-id="e23ef-277">Type</span></span>

*   [<span data-ttu-id="e23ef-278">UI</span><span class="sxs-lookup"><span data-stu-id="e23ef-278">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="e23ef-279">Requirements</span><span class="sxs-lookup"><span data-stu-id="e23ef-279">Requirements</span></span>

|<span data-ttu-id="e23ef-280">Требование</span><span class="sxs-lookup"><span data-stu-id="e23ef-280">Requirement</span></span>| <span data-ttu-id="e23ef-281">Значение</span><span class="sxs-lookup"><span data-stu-id="e23ef-281">Value</span></span>|
|---|---|
|[<span data-ttu-id="e23ef-282">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="e23ef-282">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="e23ef-283">1.1</span><span class="sxs-lookup"><span data-stu-id="e23ef-283">1.1</span></span>|
|[<span data-ttu-id="e23ef-284">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="e23ef-284">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="e23ef-285">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="e23ef-285">Compose or Read</span></span>|
