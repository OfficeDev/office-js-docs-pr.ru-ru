---
title: Office.context — предварительная версия набора обязательных элементов
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора обязательных элементов API почтового ящика.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: 64a96336ec181747fecf06c8cd2441b600ac8a10
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431117"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="0130e-103">контекст (набор требований Preview для предварительного просмотра почтового ящика)</span><span class="sxs-lookup"><span data-stu-id="0130e-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="0130e-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="0130e-104">[Office](office.md).context</span></span>

<span data-ttu-id="0130e-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="0130e-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="0130e-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span><span class="sxs-lookup"><span data-stu-id="0130e-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview&preserve-view=true).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0130e-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="0130e-107">Requirements</span></span>

|<span data-ttu-id="0130e-108">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-108">Requirement</span></span>| <span data-ttu-id="0130e-109">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0130e-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-111">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-111">1.1</span></span>|
|[<span data-ttu-id="0130e-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="0130e-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="0130e-114">Properties</span></span>

| <span data-ttu-id="0130e-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="0130e-115">Property</span></span> | <span data-ttu-id="0130e-116">Способов</span><span class="sxs-lookup"><span data-stu-id="0130e-116">Modes</span></span> | <span data-ttu-id="0130e-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="0130e-117">Return type</span></span> | <span data-ttu-id="0130e-118">Minimum</span><span class="sxs-lookup"><span data-stu-id="0130e-118">Minimum</span></span><br><span data-ttu-id="0130e-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="0130e-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0130e-120">auth</span><span class="sxs-lookup"><span data-stu-id="0130e-120">auth</span></span>](#auth-auth) | <span data-ttu-id="0130e-121">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-121">Compose</span></span><br><span data-ttu-id="0130e-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-122">Read</span></span> | [<span data-ttu-id="0130e-123">Auth</span><span class="sxs-lookup"><span data-stu-id="0130e-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="0130e-124">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0130e-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="0130e-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="0130e-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="0130e-126">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-126">Compose</span></span><br><span data-ttu-id="0130e-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-127">Read</span></span> | <span data-ttu-id="0130e-128">String</span><span class="sxs-lookup"><span data-stu-id="0130e-128">String</span></span> | [<span data-ttu-id="0130e-129">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0130e-130">диагностики</span><span class="sxs-lookup"><span data-stu-id="0130e-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="0130e-131">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-131">Compose</span></span><br><span data-ttu-id="0130e-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-132">Read</span></span> | [<span data-ttu-id="0130e-133">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="0130e-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="0130e-134">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0130e-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="0130e-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="0130e-136">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-136">Compose</span></span><br><span data-ttu-id="0130e-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-137">Read</span></span> | <span data-ttu-id="0130e-138">String</span><span class="sxs-lookup"><span data-stu-id="0130e-138">String</span></span> | [<span data-ttu-id="0130e-139">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0130e-140">размещать</span><span class="sxs-lookup"><span data-stu-id="0130e-140">host</span></span>](#host-hosttype) | <span data-ttu-id="0130e-141">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-141">Compose</span></span><br><span data-ttu-id="0130e-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-142">Read</span></span> | [<span data-ttu-id="0130e-143">HostType</span><span class="sxs-lookup"><span data-stu-id="0130e-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="0130e-144">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0130e-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="0130e-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="0130e-146">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-146">Compose</span></span><br><span data-ttu-id="0130e-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-147">Read</span></span> | [<span data-ttu-id="0130e-148">Mailbox</span><span class="sxs-lookup"><span data-stu-id="0130e-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="0130e-149">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0130e-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="0130e-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="0130e-151">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-151">Compose</span></span><br><span data-ttu-id="0130e-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-152">Read</span></span> | [<span data-ttu-id="0130e-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="0130e-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="0130e-154">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0130e-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="0130e-155">управляем</span><span class="sxs-lookup"><span data-stu-id="0130e-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="0130e-156">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-156">Compose</span></span><br><span data-ttu-id="0130e-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-157">Read</span></span> | [<span data-ttu-id="0130e-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="0130e-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="0130e-159">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0130e-160">потребность</span><span class="sxs-lookup"><span data-stu-id="0130e-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="0130e-161">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-161">Compose</span></span><br><span data-ttu-id="0130e-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-162">Read</span></span> | [<span data-ttu-id="0130e-163">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="0130e-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="0130e-164">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0130e-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="0130e-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="0130e-166">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-166">Compose</span></span><br><span data-ttu-id="0130e-167">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-167">Read</span></span> | [<span data-ttu-id="0130e-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="0130e-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="0130e-169">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0130e-170">ui</span><span class="sxs-lookup"><span data-stu-id="0130e-170">ui</span></span>](#ui-ui) | <span data-ttu-id="0130e-171">Создание</span><span class="sxs-lookup"><span data-stu-id="0130e-171">Compose</span></span><br><span data-ttu-id="0130e-172">Чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-172">Read</span></span> | [<span data-ttu-id="0130e-173">UI</span><span class="sxs-lookup"><span data-stu-id="0130e-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview&preserve-view=true) | [<span data-ttu-id="0130e-174">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="0130e-175">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="0130e-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="0130e-176">Проверка подлинности: [AUTH](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="0130e-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="0130e-177">Поддерживает [единый вход (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , предоставляя метод, позволяющий приложению Office получать маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="0130e-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office application to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="0130e-178">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="0130e-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="0130e-179">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-179">Type</span></span>

*   [<span data-ttu-id="0130e-180">Auth</span><span class="sxs-lookup"><span data-stu-id="0130e-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="0130e-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="0130e-181">Requirements</span></span>

|<span data-ttu-id="0130e-182">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-182">Requirement</span></span>| <span data-ttu-id="0130e-183">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0130e-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-185">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0130e-185">Preview</span></span>|
|[<span data-ttu-id="0130e-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0130e-188">Пример</span><span class="sxs-lookup"><span data-stu-id="0130e-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="0130e-189">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="0130e-189">contentLanguage: String</span></span>

<span data-ttu-id="0130e-190">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="0130e-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="0130e-191">`contentLanguage`Значение соответствует текущему **языковому** параметру редактирования, указанному в разделе **Параметры > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="0130e-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="0130e-192">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-192">Type</span></span>

*   <span data-ttu-id="0130e-193">String</span><span class="sxs-lookup"><span data-stu-id="0130e-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0130e-194">Требования</span><span class="sxs-lookup"><span data-stu-id="0130e-194">Requirements</span></span>

|<span data-ttu-id="0130e-195">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-195">Requirement</span></span>| <span data-ttu-id="0130e-196">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0130e-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-198">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-198">1.1</span></span>|
|[<span data-ttu-id="0130e-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0130e-201">Пример</span><span class="sxs-lookup"><span data-stu-id="0130e-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="0130e-202">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="0130e-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="0130e-203">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="0130e-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="0130e-204">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-204">Type</span></span>

*   [<span data-ttu-id="0130e-205">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="0130e-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="0130e-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="0130e-206">Requirements</span></span>

|<span data-ttu-id="0130e-207">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-207">Requirement</span></span>| <span data-ttu-id="0130e-208">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-209">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0130e-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-210">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-210">1.1</span></span>|
|[<span data-ttu-id="0130e-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-212">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0130e-213">Пример</span><span class="sxs-lookup"><span data-stu-id="0130e-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="0130e-214">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="0130e-214">displayLanguage: String</span></span>

<span data-ttu-id="0130e-215">Получает языковой стандарт (язык) в формате языковых тегов RFC 1766, указанный пользователем для пользовательского интерфейса клиентского приложения Office.</span><span class="sxs-lookup"><span data-stu-id="0130e-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office client application.</span></span>

<span data-ttu-id="0130e-216">`displayLanguage`Значение соответствует текущему параметру **языка отображения** , указанному с **параметрами > файлов > языке** в клиентском приложении Office.</span><span class="sxs-lookup"><span data-stu-id="0130e-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office client application.</span></span>

##### <a name="type"></a><span data-ttu-id="0130e-217">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-217">Type</span></span>

*   <span data-ttu-id="0130e-218">String</span><span class="sxs-lookup"><span data-stu-id="0130e-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0130e-219">Требования</span><span class="sxs-lookup"><span data-stu-id="0130e-219">Requirements</span></span>

|<span data-ttu-id="0130e-220">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-220">Requirement</span></span>| <span data-ttu-id="0130e-221">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-222">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0130e-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-223">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-223">1.1</span></span>|
|[<span data-ttu-id="0130e-224">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-225">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0130e-226">Пример</span><span class="sxs-lookup"><span data-stu-id="0130e-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="0130e-227">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="0130e-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="0130e-228">Получает приложение Office, в котором размещается надстройка.</span><span class="sxs-lookup"><span data-stu-id="0130e-228">Gets the Office application that is hosting the add-in.</span></span>

##### <a name="type"></a><span data-ttu-id="0130e-229">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-229">Type</span></span>

*   [<span data-ttu-id="0130e-230">HostType</span><span class="sxs-lookup"><span data-stu-id="0130e-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="0130e-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="0130e-231">Requirements</span></span>

|<span data-ttu-id="0130e-232">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-232">Requirement</span></span>| <span data-ttu-id="0130e-233">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-234">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0130e-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-235">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-235">1.1</span></span>|
|[<span data-ttu-id="0130e-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0130e-238">Пример</span><span class="sxs-lookup"><span data-stu-id="0130e-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="0130e-239">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="0130e-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="0130e-240">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="0130e-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="0130e-241">Этот элемент поддерживается только в Outlook для Windows.</span><span class="sxs-lookup"><span data-stu-id="0130e-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="0130e-242">Использование цветов тем Office позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с помощью **файла > учетной записи office > пользовательского интерфейса темы**Office, которая применяется ко всем клиентским приложениям Office.</span><span class="sxs-lookup"><span data-stu-id="0130e-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office client applications.</span></span> <span data-ttu-id="0130e-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="0130e-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="0130e-244">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-244">Type</span></span>

*   [<span data-ttu-id="0130e-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="0130e-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="0130e-246">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0130e-246">Properties:</span></span>

|<span data-ttu-id="0130e-247">Имя</span><span class="sxs-lookup"><span data-stu-id="0130e-247">Name</span></span>| <span data-ttu-id="0130e-248">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-248">Type</span></span>| <span data-ttu-id="0130e-249">Описание</span><span class="sxs-lookup"><span data-stu-id="0130e-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="0130e-250">String</span><span class="sxs-lookup"><span data-stu-id="0130e-250">String</span></span>|<span data-ttu-id="0130e-251">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0130e-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="0130e-252">String</span><span class="sxs-lookup"><span data-stu-id="0130e-252">String</span></span>|<span data-ttu-id="0130e-253">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0130e-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="0130e-254">String</span><span class="sxs-lookup"><span data-stu-id="0130e-254">String</span></span>|<span data-ttu-id="0130e-255">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0130e-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="0130e-256">String</span><span class="sxs-lookup"><span data-stu-id="0130e-256">String</span></span>|<span data-ttu-id="0130e-257">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0130e-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0130e-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="0130e-258">Requirements</span></span>

|<span data-ttu-id="0130e-259">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-259">Requirement</span></span>| <span data-ttu-id="0130e-260">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-261">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0130e-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-262">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0130e-262">Preview</span></span>|
|[<span data-ttu-id="0130e-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0130e-265">Пример</span><span class="sxs-lookup"><span data-stu-id="0130e-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="0130e-266">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="0130e-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="0130e-267">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="0130e-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="0130e-268">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-268">Type</span></span>

*   [<span data-ttu-id="0130e-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="0130e-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="0130e-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="0130e-270">Requirements</span></span>

|<span data-ttu-id="0130e-271">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-271">Requirement</span></span>| <span data-ttu-id="0130e-272">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-273">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0130e-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-274">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-274">1.1</span></span>|
|[<span data-ttu-id="0130e-275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-276">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0130e-277">Пример</span><span class="sxs-lookup"><span data-stu-id="0130e-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="0130e-278">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="0130e-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="0130e-279">Предоставляет метод для определения наборов требований, поддерживаемых в текущем приложении и платформе.</span><span class="sxs-lookup"><span data-stu-id="0130e-279">Provides a method for determining what requirement sets are supported on the current application and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="0130e-280">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-280">Type</span></span>

*   [<span data-ttu-id="0130e-281">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="0130e-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="0130e-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="0130e-282">Requirements</span></span>

|<span data-ttu-id="0130e-283">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-283">Requirement</span></span>| <span data-ttu-id="0130e-284">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-285">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0130e-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-286">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-286">1.1</span></span>|
|[<span data-ttu-id="0130e-287">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-288">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0130e-289">Пример</span><span class="sxs-lookup"><span data-stu-id="0130e-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="0130e-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="0130e-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="0130e-291">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="0130e-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="0130e-292">`RoamingSettings`Объект позволяет хранить и получать доступ к данным для почтовой надстройки, хранящейся в почтовом ящике пользователя, поэтому эта надстройка будет доступна для всех клиентов Outlook, используемых для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="0130e-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any Outlook client used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="0130e-293">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-293">Type</span></span>

*   [<span data-ttu-id="0130e-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="0130e-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="0130e-295">Requirements</span><span class="sxs-lookup"><span data-stu-id="0130e-295">Requirements</span></span>

|<span data-ttu-id="0130e-296">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-296">Requirement</span></span>| <span data-ttu-id="0130e-297">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-298">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0130e-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-299">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-299">1.1</span></span>|
|[<span data-ttu-id="0130e-300">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0130e-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="0130e-301">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="0130e-301">Restricted</span></span>|
|[<span data-ttu-id="0130e-302">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-303">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="0130e-304">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="0130e-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="0130e-305">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="0130e-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="0130e-306">Тип</span><span class="sxs-lookup"><span data-stu-id="0130e-306">Type</span></span>

*   [<span data-ttu-id="0130e-307">UI</span><span class="sxs-lookup"><span data-stu-id="0130e-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="0130e-308">Requirements</span><span class="sxs-lookup"><span data-stu-id="0130e-308">Requirements</span></span>

|<span data-ttu-id="0130e-309">Требование</span><span class="sxs-lookup"><span data-stu-id="0130e-309">Requirement</span></span>| <span data-ttu-id="0130e-310">Значение</span><span class="sxs-lookup"><span data-stu-id="0130e-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="0130e-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0130e-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0130e-312">1.1</span><span class="sxs-lookup"><span data-stu-id="0130e-312">1.1</span></span>|
|[<span data-ttu-id="0130e-313">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0130e-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0130e-314">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0130e-314">Compose or Read</span></span>|
