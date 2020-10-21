---
title: Office.context — предварительная версия набора обязательных элементов
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора обязательных элементов API почтового ящика.
ms.date: 10/14/2020
localization_priority: Normal
ms.openlocfilehash: 8286434d2cbfc11cf0d16f8bd014b4760f0337ff
ms.sourcegitcommit: 4e7c74ad67ea8bf6b47d65b2fde54a967090f65b
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/20/2020
ms.locfileid: "48626409"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="7e0fe-103">контекст (набор требований Preview для предварительного просмотра почтового ящика)</span><span class="sxs-lookup"><span data-stu-id="7e0fe-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="7e0fe-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="7e0fe-104">[Office](office.md).context</span></span>

<span data-ttu-id="7e0fe-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="7e0fe-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="7e0fe-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="7e0fe-107">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-107">Requirements</span></span>

|<span data-ttu-id="7e0fe-108">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-108">Requirement</span></span>| <span data-ttu-id="7e0fe-109">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7e0fe-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-111">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-111">1.1</span></span>|
|[<span data-ttu-id="7e0fe-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="7e0fe-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="7e0fe-114">Properties</span></span>

| <span data-ttu-id="7e0fe-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="7e0fe-115">Property</span></span> | <span data-ttu-id="7e0fe-116">Способов</span><span class="sxs-lookup"><span data-stu-id="7e0fe-116">Modes</span></span> | <span data-ttu-id="7e0fe-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="7e0fe-117">Return type</span></span> | <span data-ttu-id="7e0fe-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="7e0fe-118">Minimum</span></span><br><span data-ttu-id="7e0fe-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="7e0fe-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="7e0fe-120">auth</span><span class="sxs-lookup"><span data-stu-id="7e0fe-120">auth</span></span>](#auth-auth) | <span data-ttu-id="7e0fe-121">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-121">Compose</span></span><br><span data-ttu-id="7e0fe-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-122">Read</span></span> | [<span data-ttu-id="7e0fe-123">Auth</span><span class="sxs-lookup"><span data-stu-id="7e0fe-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="7e0fe-124">IdentityAPI 1,3</span><span class="sxs-lookup"><span data-stu-id="7e0fe-124">IdentityAPI 1.3</span></span>](../../requirement-sets/identity-api-requirement-sets.md) |
| [<span data-ttu-id="7e0fe-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="7e0fe-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="7e0fe-126">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-126">Compose</span></span><br><span data-ttu-id="7e0fe-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-127">Read</span></span> | <span data-ttu-id="7e0fe-128">String</span><span class="sxs-lookup"><span data-stu-id="7e0fe-128">String</span></span> | [<span data-ttu-id="7e0fe-129">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e0fe-130">диагностики</span><span class="sxs-lookup"><span data-stu-id="7e0fe-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="7e0fe-131">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-131">Compose</span></span><br><span data-ttu-id="7e0fe-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-132">Read</span></span> | [<span data-ttu-id="7e0fe-133">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="7e0fe-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="7e0fe-134">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e0fe-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="7e0fe-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="7e0fe-136">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-136">Compose</span></span><br><span data-ttu-id="7e0fe-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-137">Read</span></span> | <span data-ttu-id="7e0fe-138">String</span><span class="sxs-lookup"><span data-stu-id="7e0fe-138">String</span></span> | [<span data-ttu-id="7e0fe-139">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e0fe-140">размещать</span><span class="sxs-lookup"><span data-stu-id="7e0fe-140">host</span></span>](#host-hosttype) | <span data-ttu-id="7e0fe-141">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-141">Compose</span></span><br><span data-ttu-id="7e0fe-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-142">Read</span></span> | [<span data-ttu-id="7e0fe-143">HostType</span><span class="sxs-lookup"><span data-stu-id="7e0fe-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="7e0fe-144">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e0fe-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="7e0fe-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="7e0fe-146">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-146">Compose</span></span><br><span data-ttu-id="7e0fe-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-147">Read</span></span> | [<span data-ttu-id="7e0fe-148">Mailbox</span><span class="sxs-lookup"><span data-stu-id="7e0fe-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="7e0fe-149">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e0fe-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="7e0fe-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="7e0fe-151">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-151">Compose</span></span><br><span data-ttu-id="7e0fe-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-152">Read</span></span> | [<span data-ttu-id="7e0fe-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="7e0fe-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="7e0fe-154">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="7e0fe-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="7e0fe-155">платформа</span><span class="sxs-lookup"><span data-stu-id="7e0fe-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="7e0fe-156">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-156">Compose</span></span><br><span data-ttu-id="7e0fe-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-157">Read</span></span> | [<span data-ttu-id="7e0fe-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="7e0fe-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="7e0fe-159">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e0fe-160">потребность</span><span class="sxs-lookup"><span data-stu-id="7e0fe-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="7e0fe-161">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-161">Compose</span></span><br><span data-ttu-id="7e0fe-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-162">Read</span></span> | [<span data-ttu-id="7e0fe-163">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="7e0fe-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="7e0fe-164">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e0fe-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="7e0fe-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="7e0fe-166">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-166">Compose</span></span><br><span data-ttu-id="7e0fe-167">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-167">Read</span></span> | [<span data-ttu-id="7e0fe-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7e0fe-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="7e0fe-169">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="7e0fe-170">ui</span><span class="sxs-lookup"><span data-stu-id="7e0fe-170">ui</span></span>](#ui-ui) | <span data-ttu-id="7e0fe-171">Создание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-171">Compose</span></span><br><span data-ttu-id="7e0fe-172">Чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-172">Read</span></span> | [<span data-ttu-id="7e0fe-173">UI</span><span class="sxs-lookup"><span data-stu-id="7e0fe-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="7e0fe-174">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="7e0fe-175">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="7e0fe-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="7e0fe-176">Проверка подлинности: [AUTH](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="7e0fe-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="7e0fe-177">Поддерживает [единый вход (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , предоставляя метод, позволяющий приложению Office получать маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="7e0fe-178">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="7e0fe-179">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-179">Type</span></span>

*   [<span data-ttu-id="7e0fe-180">Auth</span><span class="sxs-lookup"><span data-stu-id="7e0fe-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="7e0fe-181">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-181">Requirements</span></span>

|<span data-ttu-id="7e0fe-182">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-182">Requirement</span></span>| <span data-ttu-id="7e0fe-183">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="7e0fe-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-185">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="7e0fe-185">Preview</span></span>|
|[<span data-ttu-id="7e0fe-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e0fe-188">Пример</span><span class="sxs-lookup"><span data-stu-id="7e0fe-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="7e0fe-189">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="7e0fe-189">contentLanguage: String</span></span>

<span data-ttu-id="7e0fe-190">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="7e0fe-191">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="7e0fe-192">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-192">Type</span></span>

*   <span data-ttu-id="7e0fe-193">String</span><span class="sxs-lookup"><span data-stu-id="7e0fe-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7e0fe-194">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-194">Requirements</span></span>

|<span data-ttu-id="7e0fe-195">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-195">Requirement</span></span>| <span data-ttu-id="7e0fe-196">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7e0fe-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-198">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-198">1.1</span></span>|
|[<span data-ttu-id="7e0fe-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e0fe-201">Пример</span><span class="sxs-lookup"><span data-stu-id="7e0fe-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="7e0fe-202">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="7e0fe-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="7e0fe-203">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="7e0fe-204">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-204">Type</span></span>

*   [<span data-ttu-id="7e0fe-205">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="7e0fe-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="7e0fe-206">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-206">Requirements</span></span>

|<span data-ttu-id="7e0fe-207">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-207">Requirement</span></span>| <span data-ttu-id="7e0fe-208">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-209">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7e0fe-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-210">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-210">1.1</span></span>|
|[<span data-ttu-id="7e0fe-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-212">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e0fe-213">Пример</span><span class="sxs-lookup"><span data-stu-id="7e0fe-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="7e0fe-214">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="7e0fe-214">displayLanguage: String</span></span>

<span data-ttu-id="7e0fe-215">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="7e0fe-216">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="7e0fe-217">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-217">Type</span></span>

*   <span data-ttu-id="7e0fe-218">String</span><span class="sxs-lookup"><span data-stu-id="7e0fe-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="7e0fe-219">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-219">Requirements</span></span>

|<span data-ttu-id="7e0fe-220">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-220">Requirement</span></span>| <span data-ttu-id="7e0fe-221">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-222">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7e0fe-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-223">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-223">1.1</span></span>|
|[<span data-ttu-id="7e0fe-224">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-225">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e0fe-226">Пример</span><span class="sxs-lookup"><span data-stu-id="7e0fe-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="7e0fe-227">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="7e0fe-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="7e0fe-228">Получает приложение Office, в котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-228">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="7e0fe-229">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-229">Type</span></span>

*   [<span data-ttu-id="7e0fe-230">HostType</span><span class="sxs-lookup"><span data-stu-id="7e0fe-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="7e0fe-231">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-231">Requirements</span></span>

|<span data-ttu-id="7e0fe-232">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-232">Requirement</span></span>| <span data-ttu-id="7e0fe-233">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-234">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7e0fe-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-235">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-235">1.1</span></span>|
|[<span data-ttu-id="7e0fe-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e0fe-238">Пример</span><span class="sxs-lookup"><span data-stu-id="7e0fe-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="7e0fe-239">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="7e0fe-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="7e0fe-240">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="7e0fe-241">Этот элемент поддерживается только в Outlook для Windows.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="7e0fe-242">Использование цветов тем Office позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с помощью **файла > учетной записи office > пользовательского интерфейса темы**Office, которая применяется ко всем клиентским приложениям Office.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="7e0fe-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="7e0fe-244">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-244">Type</span></span>

*   [<span data-ttu-id="7e0fe-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="7e0fe-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="7e0fe-246">Свойства:</span><span class="sxs-lookup"><span data-stu-id="7e0fe-246">Properties:</span></span>

|<span data-ttu-id="7e0fe-247">Имя</span><span class="sxs-lookup"><span data-stu-id="7e0fe-247">Name</span></span>| <span data-ttu-id="7e0fe-248">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-248">Type</span></span>| <span data-ttu-id="7e0fe-249">Описание</span><span class="sxs-lookup"><span data-stu-id="7e0fe-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="7e0fe-250">String</span><span class="sxs-lookup"><span data-stu-id="7e0fe-250">String</span></span>|<span data-ttu-id="7e0fe-251">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="7e0fe-252">String</span><span class="sxs-lookup"><span data-stu-id="7e0fe-252">String</span></span>|<span data-ttu-id="7e0fe-253">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="7e0fe-254">String</span><span class="sxs-lookup"><span data-stu-id="7e0fe-254">String</span></span>|<span data-ttu-id="7e0fe-255">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="7e0fe-256">String</span><span class="sxs-lookup"><span data-stu-id="7e0fe-256">String</span></span>|<span data-ttu-id="7e0fe-257">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="7e0fe-258">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-258">Requirements</span></span>

|<span data-ttu-id="7e0fe-259">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-259">Requirement</span></span>| <span data-ttu-id="7e0fe-260">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-261">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="7e0fe-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-262">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="7e0fe-262">Preview</span></span>|
|[<span data-ttu-id="7e0fe-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e0fe-265">Пример</span><span class="sxs-lookup"><span data-stu-id="7e0fe-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="7e0fe-266">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="7e0fe-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="7e0fe-267">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="7e0fe-268">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-268">Type</span></span>

*   [<span data-ttu-id="7e0fe-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="7e0fe-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="7e0fe-270">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-270">Requirements</span></span>

|<span data-ttu-id="7e0fe-271">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-271">Requirement</span></span>| <span data-ttu-id="7e0fe-272">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-273">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7e0fe-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-274">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-274">1.1</span></span>|
|[<span data-ttu-id="7e0fe-275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-276">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e0fe-277">Пример</span><span class="sxs-lookup"><span data-stu-id="7e0fe-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="7e0fe-278">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="7e0fe-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="7e0fe-279">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-279">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="7e0fe-280">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-280">Type</span></span>

*   [<span data-ttu-id="7e0fe-281">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="7e0fe-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="7e0fe-282">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-282">Requirements</span></span>

|<span data-ttu-id="7e0fe-283">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-283">Requirement</span></span>| <span data-ttu-id="7e0fe-284">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-285">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7e0fe-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-286">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-286">1.1</span></span>|
|[<span data-ttu-id="7e0fe-287">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-288">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="7e0fe-289">Пример</span><span class="sxs-lookup"><span data-stu-id="7e0fe-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="7e0fe-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="7e0fe-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="7e0fe-291">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="7e0fe-292">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="7e0fe-293">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-293">Type</span></span>

*   [<span data-ttu-id="7e0fe-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="7e0fe-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="7e0fe-295">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-295">Requirements</span></span>

|<span data-ttu-id="7e0fe-296">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-296">Requirement</span></span>| <span data-ttu-id="7e0fe-297">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-298">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7e0fe-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-299">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-299">1.1</span></span>|
|[<span data-ttu-id="7e0fe-300">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="7e0fe-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="7e0fe-301">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="7e0fe-301">Restricted</span></span>|
|[<span data-ttu-id="7e0fe-302">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-303">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="7e0fe-304">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="7e0fe-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="7e0fe-305">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="7e0fe-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="7e0fe-306">Тип</span><span class="sxs-lookup"><span data-stu-id="7e0fe-306">Type</span></span>

*   [<span data-ttu-id="7e0fe-307">UI</span><span class="sxs-lookup"><span data-stu-id="7e0fe-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="7e0fe-308">Требования</span><span class="sxs-lookup"><span data-stu-id="7e0fe-308">Requirements</span></span>

|<span data-ttu-id="7e0fe-309">Требование</span><span class="sxs-lookup"><span data-stu-id="7e0fe-309">Requirement</span></span>| <span data-ttu-id="7e0fe-310">Значение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="7e0fe-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="7e0fe-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="7e0fe-312">1.1</span><span class="sxs-lookup"><span data-stu-id="7e0fe-312">1.1</span></span>|
|[<span data-ttu-id="7e0fe-313">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="7e0fe-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="7e0fe-314">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="7e0fe-314">Compose or Read</span></span>|
