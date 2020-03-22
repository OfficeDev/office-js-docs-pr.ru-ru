---
title: Office.context — предварительная версия набора обязательных элементов
description: Элементы объекта Office. Context, доступные для надстроек Outlook с помощью набора обязательных элементов API почтового ящика.
ms.date: 03/18/2020
localization_priority: Normal
ms.openlocfilehash: c61769cb1ae98097ffabb8b3ef19b2f82257c2b1
ms.sourcegitcommit: 6c381634c77d316f34747131860db0a0bced2529
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/21/2020
ms.locfileid: "42890867"
---
# <a name="context-mailbox-preview-requirement-set"></a><span data-ttu-id="f599f-103">контекст (набор требований Preview для предварительного просмотра почтового ящика)</span><span class="sxs-lookup"><span data-stu-id="f599f-103">context (Mailbox preview requirement set)</span></span>

### <a name="officecontext"></a><span data-ttu-id="f599f-104">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="f599f-104">[Office](office.md).context</span></span>

<span data-ttu-id="f599f-105">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="f599f-105">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="f599f-106">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="f599f-106">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="f599f-107">Requirements</span><span class="sxs-lookup"><span data-stu-id="f599f-107">Requirements</span></span>

|<span data-ttu-id="f599f-108">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-108">Requirement</span></span>| <span data-ttu-id="f599f-109">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-109">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-110">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f599f-110">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-111">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-111">1.1</span></span>|
|[<span data-ttu-id="f599f-112">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-112">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-113">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-113">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f599f-114">Свойства</span><span class="sxs-lookup"><span data-stu-id="f599f-114">Properties</span></span>

| <span data-ttu-id="f599f-115">Свойство</span><span class="sxs-lookup"><span data-stu-id="f599f-115">Property</span></span> | <span data-ttu-id="f599f-116">Способов</span><span class="sxs-lookup"><span data-stu-id="f599f-116">Modes</span></span> | <span data-ttu-id="f599f-117">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="f599f-117">Return type</span></span> | <span data-ttu-id="f599f-118">Минимальные</span><span class="sxs-lookup"><span data-stu-id="f599f-118">Minimum</span></span><br><span data-ttu-id="f599f-119">набор требований</span><span class="sxs-lookup"><span data-stu-id="f599f-119">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="f599f-120">auth</span><span class="sxs-lookup"><span data-stu-id="f599f-120">auth</span></span>](#auth-auth) | <span data-ttu-id="f599f-121">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-121">Compose</span></span><br><span data-ttu-id="f599f-122">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-122">Read</span></span> | [<span data-ttu-id="f599f-123">Auth</span><span class="sxs-lookup"><span data-stu-id="f599f-123">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="f599f-124">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="f599f-124">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="f599f-125">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="f599f-125">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="f599f-126">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-126">Compose</span></span><br><span data-ttu-id="f599f-127">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-127">Read</span></span> | <span data-ttu-id="f599f-128">String</span><span class="sxs-lookup"><span data-stu-id="f599f-128">String</span></span> | [<span data-ttu-id="f599f-129">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-129">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f599f-130">diagnostics</span><span class="sxs-lookup"><span data-stu-id="f599f-130">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="f599f-131">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-131">Compose</span></span><br><span data-ttu-id="f599f-132">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-132">Read</span></span> | [<span data-ttu-id="f599f-133">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="f599f-133">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="f599f-134">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-134">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f599f-135">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="f599f-135">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="f599f-136">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-136">Compose</span></span><br><span data-ttu-id="f599f-137">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-137">Read</span></span> | <span data-ttu-id="f599f-138">String</span><span class="sxs-lookup"><span data-stu-id="f599f-138">String</span></span> | [<span data-ttu-id="f599f-139">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-139">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f599f-140">размещать</span><span class="sxs-lookup"><span data-stu-id="f599f-140">host</span></span>](#host-hosttype) | <span data-ttu-id="f599f-141">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-141">Compose</span></span><br><span data-ttu-id="f599f-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-142">Read</span></span> | [<span data-ttu-id="f599f-143">HostType</span><span class="sxs-lookup"><span data-stu-id="f599f-143">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="f599f-144">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-144">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f599f-145">mailbox</span><span class="sxs-lookup"><span data-stu-id="f599f-145">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="f599f-146">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-146">Compose</span></span><br><span data-ttu-id="f599f-147">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-147">Read</span></span> | [<span data-ttu-id="f599f-148">Mailbox</span><span class="sxs-lookup"><span data-stu-id="f599f-148">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="f599f-149">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-149">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f599f-150">officeTheme</span><span class="sxs-lookup"><span data-stu-id="f599f-150">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="f599f-151">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-151">Compose</span></span><br><span data-ttu-id="f599f-152">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-152">Read</span></span> | [<span data-ttu-id="f599f-153">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="f599f-153">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="f599f-154">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="f599f-154">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="f599f-155">управляем</span><span class="sxs-lookup"><span data-stu-id="f599f-155">platform</span></span>](#platform-platformtype) | <span data-ttu-id="f599f-156">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-156">Compose</span></span><br><span data-ttu-id="f599f-157">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-157">Read</span></span> | [<span data-ttu-id="f599f-158">PlatformType</span><span class="sxs-lookup"><span data-stu-id="f599f-158">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="f599f-159">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-159">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f599f-160">потребность</span><span class="sxs-lookup"><span data-stu-id="f599f-160">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="f599f-161">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-161">Compose</span></span><br><span data-ttu-id="f599f-162">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-162">Read</span></span> | [<span data-ttu-id="f599f-163">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="f599f-163">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="f599f-164">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-164">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f599f-165">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="f599f-165">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="f599f-166">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-166">Compose</span></span><br><span data-ttu-id="f599f-167">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-167">Read</span></span> | [<span data-ttu-id="f599f-168">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="f599f-168">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="f599f-169">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-169">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="f599f-170">ui</span><span class="sxs-lookup"><span data-stu-id="f599f-170">ui</span></span>](#ui-ui) | <span data-ttu-id="f599f-171">Создание</span><span class="sxs-lookup"><span data-stu-id="f599f-171">Compose</span></span><br><span data-ttu-id="f599f-172">Чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-172">Read</span></span> | [<span data-ttu-id="f599f-173">UI</span><span class="sxs-lookup"><span data-stu-id="f599f-173">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="f599f-174">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-174">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="f599f-175">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="f599f-175">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="f599f-176">Проверка подлинности: [AUTH](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="f599f-176">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="f599f-177">Поддерживает [единый вход (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , предоставляя метод, позволяющий ведущему приложению Office получать маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="f599f-177">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="f599f-178">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="f599f-178">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="f599f-179">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-179">Type</span></span>

*   [<span data-ttu-id="f599f-180">Auth</span><span class="sxs-lookup"><span data-stu-id="f599f-180">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="f599f-181">Requirements</span><span class="sxs-lookup"><span data-stu-id="f599f-181">Requirements</span></span>

|<span data-ttu-id="f599f-182">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-182">Requirement</span></span>| <span data-ttu-id="f599f-183">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-183">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-184">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f599f-184">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-185">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="f599f-185">Preview</span></span>|
|[<span data-ttu-id="f599f-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-186">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f599f-188">Пример</span><span class="sxs-lookup"><span data-stu-id="f599f-188">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="f599f-189">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="f599f-189">contentLanguage: String</span></span>

<span data-ttu-id="f599f-190">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="f599f-190">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="f599f-191">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="f599f-191">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="f599f-192">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-192">Type</span></span>

*   <span data-ttu-id="f599f-193">String</span><span class="sxs-lookup"><span data-stu-id="f599f-193">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f599f-194">Требования</span><span class="sxs-lookup"><span data-stu-id="f599f-194">Requirements</span></span>

|<span data-ttu-id="f599f-195">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-195">Requirement</span></span>| <span data-ttu-id="f599f-196">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-196">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-197">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f599f-197">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-198">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-198">1.1</span></span>|
|[<span data-ttu-id="f599f-199">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-199">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-200">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-200">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f599f-201">Пример</span><span class="sxs-lookup"><span data-stu-id="f599f-201">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="f599f-202">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="f599f-202">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="f599f-203">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="f599f-203">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="f599f-204">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-204">Type</span></span>

*   [<span data-ttu-id="f599f-205">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="f599f-205">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="f599f-206">Requirements</span><span class="sxs-lookup"><span data-stu-id="f599f-206">Requirements</span></span>

|<span data-ttu-id="f599f-207">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-207">Requirement</span></span>| <span data-ttu-id="f599f-208">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-208">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-209">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f599f-209">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-210">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-210">1.1</span></span>|
|[<span data-ttu-id="f599f-211">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-211">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-212">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-212">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f599f-213">Пример</span><span class="sxs-lookup"><span data-stu-id="f599f-213">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="f599f-214">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="f599f-214">displayLanguage: String</span></span>

<span data-ttu-id="f599f-215">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="f599f-215">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="f599f-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="f599f-216">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="f599f-217">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-217">Type</span></span>

*   <span data-ttu-id="f599f-218">String</span><span class="sxs-lookup"><span data-stu-id="f599f-218">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f599f-219">Требования</span><span class="sxs-lookup"><span data-stu-id="f599f-219">Requirements</span></span>

|<span data-ttu-id="f599f-220">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-220">Requirement</span></span>| <span data-ttu-id="f599f-221">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-221">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-222">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f599f-222">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-223">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-223">1.1</span></span>|
|[<span data-ttu-id="f599f-224">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-224">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-225">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-225">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f599f-226">Пример</span><span class="sxs-lookup"><span data-stu-id="f599f-226">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="f599f-227">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="f599f-227">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="f599f-228">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="f599f-228">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="f599f-229">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-229">Type</span></span>

*   [<span data-ttu-id="f599f-230">HostType</span><span class="sxs-lookup"><span data-stu-id="f599f-230">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="f599f-231">Requirements</span><span class="sxs-lookup"><span data-stu-id="f599f-231">Requirements</span></span>

|<span data-ttu-id="f599f-232">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-232">Requirement</span></span>| <span data-ttu-id="f599f-233">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-233">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-234">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f599f-234">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-235">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-235">1.1</span></span>|
|[<span data-ttu-id="f599f-236">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-236">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-237">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-237">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f599f-238">Пример</span><span class="sxs-lookup"><span data-stu-id="f599f-238">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="f599f-239">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="f599f-239">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="f599f-240">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="f599f-240">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="f599f-241">Этот элемент поддерживается только в Outlook для Windows.</span><span class="sxs-lookup"><span data-stu-id="f599f-241">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="f599f-242">Использование цветов тем Office позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с помощью **файла > учетной записи office > пользовательского интерфейса темы**Office, которая применяется ко всем ведущим приложениям Office.</span><span class="sxs-lookup"><span data-stu-id="f599f-242">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="f599f-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="f599f-243">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="f599f-244">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-244">Type</span></span>

*   [<span data-ttu-id="f599f-245">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="f599f-245">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="f599f-246">Свойства:</span><span class="sxs-lookup"><span data-stu-id="f599f-246">Properties:</span></span>

|<span data-ttu-id="f599f-247">Имя</span><span class="sxs-lookup"><span data-stu-id="f599f-247">Name</span></span>| <span data-ttu-id="f599f-248">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-248">Type</span></span>| <span data-ttu-id="f599f-249">Описание</span><span class="sxs-lookup"><span data-stu-id="f599f-249">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="f599f-250">String</span><span class="sxs-lookup"><span data-stu-id="f599f-250">String</span></span>|<span data-ttu-id="f599f-251">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="f599f-251">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="f599f-252">String</span><span class="sxs-lookup"><span data-stu-id="f599f-252">String</span></span>|<span data-ttu-id="f599f-253">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="f599f-253">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="f599f-254">String</span><span class="sxs-lookup"><span data-stu-id="f599f-254">String</span></span>|<span data-ttu-id="f599f-255">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="f599f-255">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="f599f-256">String</span><span class="sxs-lookup"><span data-stu-id="f599f-256">String</span></span>|<span data-ttu-id="f599f-257">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="f599f-257">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="f599f-258">Requirements</span><span class="sxs-lookup"><span data-stu-id="f599f-258">Requirements</span></span>

|<span data-ttu-id="f599f-259">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-259">Requirement</span></span>| <span data-ttu-id="f599f-260">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-260">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-261">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f599f-261">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-262">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="f599f-262">Preview</span></span>|
|[<span data-ttu-id="f599f-263">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-263">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-264">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-264">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f599f-265">Пример</span><span class="sxs-lookup"><span data-stu-id="f599f-265">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="f599f-266">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="f599f-266">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="f599f-267">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="f599f-267">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="f599f-268">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-268">Type</span></span>

*   [<span data-ttu-id="f599f-269">PlatformType</span><span class="sxs-lookup"><span data-stu-id="f599f-269">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="f599f-270">Requirements</span><span class="sxs-lookup"><span data-stu-id="f599f-270">Requirements</span></span>

|<span data-ttu-id="f599f-271">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-271">Requirement</span></span>| <span data-ttu-id="f599f-272">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-272">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-273">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f599f-273">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-274">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-274">1.1</span></span>|
|[<span data-ttu-id="f599f-275">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-275">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-276">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-276">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f599f-277">Пример</span><span class="sxs-lookup"><span data-stu-id="f599f-277">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="f599f-278">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="f599f-278">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="f599f-279">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="f599f-279">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="f599f-280">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-280">Type</span></span>

*   [<span data-ttu-id="f599f-281">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="f599f-281">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="f599f-282">Requirements</span><span class="sxs-lookup"><span data-stu-id="f599f-282">Requirements</span></span>

|<span data-ttu-id="f599f-283">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-283">Requirement</span></span>| <span data-ttu-id="f599f-284">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-284">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-285">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f599f-285">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-286">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-286">1.1</span></span>|
|[<span data-ttu-id="f599f-287">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-287">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-288">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-288">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f599f-289">Пример</span><span class="sxs-lookup"><span data-stu-id="f599f-289">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="f599f-290">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="f599f-290">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="f599f-291">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="f599f-291">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="f599f-292">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="f599f-292">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="f599f-293">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-293">Type</span></span>

*   [<span data-ttu-id="f599f-294">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="f599f-294">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="f599f-295">Requirements</span><span class="sxs-lookup"><span data-stu-id="f599f-295">Requirements</span></span>

|<span data-ttu-id="f599f-296">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-296">Requirement</span></span>| <span data-ttu-id="f599f-297">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-297">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-298">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f599f-298">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-299">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-299">1.1</span></span>|
|[<span data-ttu-id="f599f-300">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f599f-300">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="f599f-301">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="f599f-301">Restricted</span></span>|
|[<span data-ttu-id="f599f-302">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-302">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-303">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-303">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="f599f-304">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="f599f-304">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="f599f-305">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="f599f-305">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="f599f-306">Тип</span><span class="sxs-lookup"><span data-stu-id="f599f-306">Type</span></span>

*   [<span data-ttu-id="f599f-307">UI</span><span class="sxs-lookup"><span data-stu-id="f599f-307">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="f599f-308">Requirements</span><span class="sxs-lookup"><span data-stu-id="f599f-308">Requirements</span></span>

|<span data-ttu-id="f599f-309">Требование</span><span class="sxs-lookup"><span data-stu-id="f599f-309">Requirement</span></span>| <span data-ttu-id="f599f-310">Значение</span><span class="sxs-lookup"><span data-stu-id="f599f-310">Value</span></span>|
|---|---|
|[<span data-ttu-id="f599f-311">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f599f-311">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="f599f-312">1.1</span><span class="sxs-lookup"><span data-stu-id="f599f-312">1.1</span></span>|
|[<span data-ttu-id="f599f-313">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f599f-313">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="f599f-314">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f599f-314">Compose or Read</span></span>|
