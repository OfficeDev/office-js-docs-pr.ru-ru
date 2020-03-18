---
title: Office.context — предварительная версия набора обязательных элементов
description: Объектная модель для объекта контекста Outlook в API надстроек Outlook (Предварительная версия API почтовых ящиков).
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 409f0a5b46eba667f79228f45081c160c3c3ce7f
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42717805"
---
# <a name="context"></a><span data-ttu-id="d723b-103">context</span><span class="sxs-lookup"><span data-stu-id="d723b-103">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="d723b-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="d723b-104">[Office](office.md).context</span></span>

<span data-ttu-id="d723b-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="d723b-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="d723b-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="d723b-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="d723b-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="d723b-107">Requirements</span></span>

|<span data-ttu-id="d723b-108">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-108">Requirement</span></span>| <span data-ttu-id="d723b-109">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d723b-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-111">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-111">1.1</span></span>|
|[<span data-ttu-id="d723b-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="d723b-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="d723b-114">Properties</span></span>

| <span data-ttu-id="d723b-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="d723b-115">Property</span></span> | <span data-ttu-id="d723b-116">Способов</span><span class="sxs-lookup"><span data-stu-id="d723b-116">Modes</span></span> | <span data-ttu-id="d723b-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="d723b-117">Return type</span></span> | <span data-ttu-id="d723b-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="d723b-118">Minimum</span></span><br><span data-ttu-id="d723b-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="d723b-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="d723b-120">auth</span><span class="sxs-lookup"><span data-stu-id="d723b-120">auth</span></span>](#auth-auth) | <span data-ttu-id="d723b-121">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-121">Compose</span></span><br><span data-ttu-id="d723b-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-122">Read</span></span> | [<span data-ttu-id="d723b-123">Auth</span><span class="sxs-lookup"><span data-stu-id="d723b-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="d723b-124">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="d723b-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="d723b-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="d723b-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="d723b-126">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-126">Compose</span></span><br><span data-ttu-id="d723b-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-127">Read</span></span> | <span data-ttu-id="d723b-128">String</span><span class="sxs-lookup"><span data-stu-id="d723b-128">String</span></span> | [<span data-ttu-id="d723b-129">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d723b-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="d723b-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="d723b-131">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-131">Compose</span></span><br><span data-ttu-id="d723b-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-132">Read</span></span> | [<span data-ttu-id="d723b-133">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="d723b-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="d723b-134">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d723b-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="d723b-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="d723b-136">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-136">Compose</span></span><br><span data-ttu-id="d723b-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-137">Read</span></span> | <span data-ttu-id="d723b-138">String</span><span class="sxs-lookup"><span data-stu-id="d723b-138">String</span></span> | [<span data-ttu-id="d723b-139">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d723b-140">размещать</span><span class="sxs-lookup"><span data-stu-id="d723b-140">host</span></span>](#host-hosttype) | <span data-ttu-id="d723b-141">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-141">Compose</span></span><br><span data-ttu-id="d723b-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-142">Read</span></span> | [<span data-ttu-id="d723b-143">HostType</span><span class="sxs-lookup"><span data-stu-id="d723b-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="d723b-144">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d723b-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="d723b-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="d723b-146">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-146">Compose</span></span><br><span data-ttu-id="d723b-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-147">Read</span></span> | [<span data-ttu-id="d723b-148">Mailbox</span><span class="sxs-lookup"><span data-stu-id="d723b-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="d723b-149">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d723b-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="d723b-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="d723b-151">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-151">Compose</span></span><br><span data-ttu-id="d723b-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-152">Read</span></span> | [<span data-ttu-id="d723b-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="d723b-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="d723b-154">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="d723b-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="d723b-155">управляем</span><span class="sxs-lookup"><span data-stu-id="d723b-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="d723b-156">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-156">Compose</span></span><br><span data-ttu-id="d723b-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-157">Read</span></span> | [<span data-ttu-id="d723b-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="d723b-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="d723b-159">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d723b-160">потребность</span><span class="sxs-lookup"><span data-stu-id="d723b-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="d723b-161">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-161">Compose</span></span><br><span data-ttu-id="d723b-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-162">Read</span></span> | [<span data-ttu-id="d723b-163">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="d723b-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="d723b-164">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d723b-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="d723b-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="d723b-166">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-166">Compose</span></span><br><span data-ttu-id="d723b-167">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-167">Read</span></span> | [<span data-ttu-id="d723b-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d723b-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="d723b-169">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="d723b-170">ui</span><span class="sxs-lookup"><span data-stu-id="d723b-170">ui</span></span>](#ui-ui) | <span data-ttu-id="d723b-171">Создание</span><span class="sxs-lookup"><span data-stu-id="d723b-171">Compose</span></span><br><span data-ttu-id="d723b-172">Чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-172">Read</span></span> | [<span data-ttu-id="d723b-173">UI</span><span class="sxs-lookup"><span data-stu-id="d723b-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="d723b-174">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="d723b-175">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="d723b-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="d723b-176">Проверка подлинности: [AUTH](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="d723b-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="d723b-177">Поддерживает [единый вход (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , предоставляя метод, позволяющий ведущему приложению Office получать маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="d723b-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="d723b-178">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="d723b-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="d723b-179">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-179">Type</span></span>

*   [<span data-ttu-id="d723b-180">Auth</span><span class="sxs-lookup"><span data-stu-id="d723b-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="d723b-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="d723b-181">Requirements</span></span>

|<span data-ttu-id="d723b-182">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-182">Requirement</span></span>| <span data-ttu-id="d723b-183">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d723b-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-185">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="d723b-185">Preview</span></span>|
|[<span data-ttu-id="d723b-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d723b-188">Пример</span><span class="sxs-lookup"><span data-stu-id="d723b-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="d723b-189">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="d723b-189">contentLanguage: String</span></span>

<span data-ttu-id="d723b-190">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="d723b-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="d723b-191">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="d723b-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="d723b-192">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-192">Type</span></span>

*   <span data-ttu-id="d723b-193">String</span><span class="sxs-lookup"><span data-stu-id="d723b-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d723b-194">Требования</span><span class="sxs-lookup"><span data-stu-id="d723b-194">Requirements</span></span>

|<span data-ttu-id="d723b-195">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-195">Requirement</span></span>| <span data-ttu-id="d723b-196">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d723b-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-198">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-198">1.1</span></span>|
|[<span data-ttu-id="d723b-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d723b-201">Пример</span><span class="sxs-lookup"><span data-stu-id="d723b-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="d723b-202">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="d723b-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="d723b-203">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="d723b-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="d723b-204">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-204">Type</span></span>

*   [<span data-ttu-id="d723b-205">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="d723b-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="d723b-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="d723b-206">Requirements</span></span>

|<span data-ttu-id="d723b-207">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-207">Requirement</span></span>| <span data-ttu-id="d723b-208">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-209">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d723b-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-210">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-210">1.1</span></span>|
|[<span data-ttu-id="d723b-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-212">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d723b-213">Пример</span><span class="sxs-lookup"><span data-stu-id="d723b-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="d723b-214">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="d723b-214">displayLanguage: String</span></span>

<span data-ttu-id="d723b-215">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="d723b-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="d723b-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="d723b-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="d723b-217">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-217">Type</span></span>

*   <span data-ttu-id="d723b-218">String</span><span class="sxs-lookup"><span data-stu-id="d723b-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d723b-219">Требования</span><span class="sxs-lookup"><span data-stu-id="d723b-219">Requirements</span></span>

|<span data-ttu-id="d723b-220">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-220">Requirement</span></span>| <span data-ttu-id="d723b-221">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-222">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d723b-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-223">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-223">1.1</span></span>|
|[<span data-ttu-id="d723b-224">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-225">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d723b-226">Пример</span><span class="sxs-lookup"><span data-stu-id="d723b-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="d723b-227">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="d723b-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="d723b-228">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="d723b-228">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="d723b-229">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-229">Type</span></span>

*   [<span data-ttu-id="d723b-230">HostType</span><span class="sxs-lookup"><span data-stu-id="d723b-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="d723b-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="d723b-231">Requirements</span></span>

|<span data-ttu-id="d723b-232">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-232">Requirement</span></span>| <span data-ttu-id="d723b-233">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-234">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d723b-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-235">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-235">1.1</span></span>|
|[<span data-ttu-id="d723b-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d723b-238">Пример</span><span class="sxs-lookup"><span data-stu-id="d723b-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="d723b-239">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="d723b-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="d723b-240">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="d723b-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="d723b-241">Этот элемент поддерживается только в Outlook для Windows.</span><span class="sxs-lookup"><span data-stu-id="d723b-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="d723b-242">Использование цветов тем Office позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с помощью **файла > учетной записи office > пользовательского интерфейса темы**Office, которая применяется ко всем ведущим приложениям Office.</span><span class="sxs-lookup"><span data-stu-id="d723b-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="d723b-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="d723b-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="d723b-244">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-244">Type</span></span>

*   [<span data-ttu-id="d723b-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="d723b-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="d723b-246">Свойства:</span><span class="sxs-lookup"><span data-stu-id="d723b-246">Properties:</span></span>

|<span data-ttu-id="d723b-247">Имя</span><span class="sxs-lookup"><span data-stu-id="d723b-247">Name</span></span>| <span data-ttu-id="d723b-248">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-248">Type</span></span>| <span data-ttu-id="d723b-249">Описание</span><span class="sxs-lookup"><span data-stu-id="d723b-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="d723b-250">String</span><span class="sxs-lookup"><span data-stu-id="d723b-250">String</span></span>|<span data-ttu-id="d723b-251">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="d723b-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="d723b-252">String</span><span class="sxs-lookup"><span data-stu-id="d723b-252">String</span></span>|<span data-ttu-id="d723b-253">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="d723b-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="d723b-254">String</span><span class="sxs-lookup"><span data-stu-id="d723b-254">String</span></span>|<span data-ttu-id="d723b-255">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="d723b-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="d723b-256">String</span><span class="sxs-lookup"><span data-stu-id="d723b-256">String</span></span>|<span data-ttu-id="d723b-257">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="d723b-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="d723b-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="d723b-258">Requirements</span></span>

|<span data-ttu-id="d723b-259">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-259">Requirement</span></span>| <span data-ttu-id="d723b-260">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-261">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="d723b-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-262">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="d723b-262">Preview</span></span>|
|[<span data-ttu-id="d723b-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d723b-265">Пример</span><span class="sxs-lookup"><span data-stu-id="d723b-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="d723b-266">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="d723b-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="d723b-267">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="d723b-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="d723b-268">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-268">Type</span></span>

*   [<span data-ttu-id="d723b-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="d723b-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="d723b-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="d723b-270">Requirements</span></span>

|<span data-ttu-id="d723b-271">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-271">Requirement</span></span>| <span data-ttu-id="d723b-272">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-273">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d723b-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-274">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-274">1.1</span></span>|
|[<span data-ttu-id="d723b-275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-276">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d723b-277">Пример</span><span class="sxs-lookup"><span data-stu-id="d723b-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="d723b-278">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="d723b-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="d723b-279">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="d723b-279">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="d723b-280">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-280">Type</span></span>

*   [<span data-ttu-id="d723b-281">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="d723b-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="d723b-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="d723b-282">Requirements</span></span>

|<span data-ttu-id="d723b-283">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-283">Requirement</span></span>| <span data-ttu-id="d723b-284">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-285">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d723b-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-286">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-286">1.1</span></span>|
|[<span data-ttu-id="d723b-287">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-288">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d723b-289">Пример</span><span class="sxs-lookup"><span data-stu-id="d723b-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="d723b-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="d723b-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="d723b-291">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="d723b-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="d723b-292">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="d723b-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="d723b-293">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-293">Type</span></span>

*   [<span data-ttu-id="d723b-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="d723b-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="d723b-295">Requirements</span><span class="sxs-lookup"><span data-stu-id="d723b-295">Requirements</span></span>

|<span data-ttu-id="d723b-296">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-296">Requirement</span></span>| <span data-ttu-id="d723b-297">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-298">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d723b-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-299">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-299">1.1</span></span>|
|[<span data-ttu-id="d723b-300">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d723b-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="d723b-301">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="d723b-301">Restricted</span></span>|
|[<span data-ttu-id="d723b-302">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-303">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="d723b-304">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="d723b-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="d723b-305">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="d723b-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="d723b-306">Тип</span><span class="sxs-lookup"><span data-stu-id="d723b-306">Type</span></span>

*   [<span data-ttu-id="d723b-307">UI</span><span class="sxs-lookup"><span data-stu-id="d723b-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="d723b-308">Requirements</span><span class="sxs-lookup"><span data-stu-id="d723b-308">Requirements</span></span>

|<span data-ttu-id="d723b-309">Требование</span><span class="sxs-lookup"><span data-stu-id="d723b-309">Requirement</span></span>| <span data-ttu-id="d723b-310">Значение</span><span class="sxs-lookup"><span data-stu-id="d723b-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="d723b-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d723b-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="d723b-312">1.1</span><span class="sxs-lookup"><span data-stu-id="d723b-312">1.1</span></span>|
|[<span data-ttu-id="d723b-313">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d723b-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="d723b-314">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d723b-314">Compose or Read</span></span>|
