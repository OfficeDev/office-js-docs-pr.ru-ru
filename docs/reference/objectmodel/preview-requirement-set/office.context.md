---
title: Office.context — предварительная версия набора обязательных элементов
description: ''
ms.date: 12/16/2019
localization_priority: Normal
ms.openlocfilehash: 9c2c661ce870e2007bd891aee040c6b3564f7b9e
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42165519"
---
# <a name="context"></a><span data-ttu-id="0366a-102">context</span><span class="sxs-lookup"><span data-stu-id="0366a-102">context</span></span>

### <a name="officecontext"></a><span data-ttu-id="0366a-103">[Office](office.md).context</span><span class="sxs-lookup"><span data-stu-id="0366a-103">[Office](office.md).context</span></span>

<span data-ttu-id="0366a-104">Office. context предоставляет общие интерфейсы, которые используются надстройками во всех приложениях Office.</span><span class="sxs-lookup"><span data-stu-id="0366a-104">Office.context provides shared interfaces that are used by add-ins in all of the Office apps.</span></span> <span data-ttu-id="0366a-105">В этом листинге документируется только те интерфейсы, которые используются надстройками Outlook. Полный список пространств имен Office. Context представлен в статье [Справочник по Office. Context в общем API](/javascript/api/office/office.context?view=outlook-js-preview).</span><span class="sxs-lookup"><span data-stu-id="0366a-105">This listing documents only those interfaces that are used by Outlook add-ins. For a full listing of the Office.context namespace, see the [Office.context reference in the Common API](/javascript/api/office/office.context?view=outlook-js-preview).</span></span>

##### <a name="requirements"></a><span data-ttu-id="0366a-106">Requirements</span><span class="sxs-lookup"><span data-stu-id="0366a-106">Requirements</span></span>

|<span data-ttu-id="0366a-107">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-107">Requirement</span></span>| <span data-ttu-id="0366a-108">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-108">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-109">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0366a-109">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-110">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-110">1.1</span></span>|
|[<span data-ttu-id="0366a-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-111">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="0366a-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="0366a-113">Properties</span></span>

| <span data-ttu-id="0366a-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="0366a-114">Property</span></span> | <span data-ttu-id="0366a-115">Способов</span><span class="sxs-lookup"><span data-stu-id="0366a-115">Modes</span></span> | <span data-ttu-id="0366a-116">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="0366a-116">Return type</span></span> | <span data-ttu-id="0366a-117">Минимальные</span><span class="sxs-lookup"><span data-stu-id="0366a-117">Minimum</span></span><br><span data-ttu-id="0366a-118">набор требований</span><span class="sxs-lookup"><span data-stu-id="0366a-118">requirement set</span></span> |
|---|---|---|:---:|
| [<span data-ttu-id="0366a-119">auth</span><span class="sxs-lookup"><span data-stu-id="0366a-119">auth</span></span>](#auth-auth) | <span data-ttu-id="0366a-120">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-120">Compose</span></span><br><span data-ttu-id="0366a-121">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-121">Read</span></span> | [<span data-ttu-id="0366a-122">Auth</span><span class="sxs-lookup"><span data-stu-id="0366a-122">Auth</span></span>](/javascript/api/office/office.auth?view=outlook-js-preview) | [<span data-ttu-id="0366a-123">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0366a-123">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="0366a-124">contentLanguage</span><span class="sxs-lookup"><span data-stu-id="0366a-124">contentLanguage</span></span>](#contentlanguage-string) | <span data-ttu-id="0366a-125">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-125">Compose</span></span><br><span data-ttu-id="0366a-126">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-126">Read</span></span> | <span data-ttu-id="0366a-127">String</span><span class="sxs-lookup"><span data-stu-id="0366a-127">String</span></span> | [<span data-ttu-id="0366a-128">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-128">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0366a-129">diagnostics</span><span class="sxs-lookup"><span data-stu-id="0366a-129">diagnostics</span></span>](#diagnostics-contextinformation) | <span data-ttu-id="0366a-130">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-130">Compose</span></span><br><span data-ttu-id="0366a-131">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-131">Read</span></span> | [<span data-ttu-id="0366a-132">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="0366a-132">ContextInformation</span></span>](/javascript/api/office/office.contextinformation?view=outlook-js-preview) | [<span data-ttu-id="0366a-133">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-133">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0366a-134">displayLanguage</span><span class="sxs-lookup"><span data-stu-id="0366a-134">displayLanguage</span></span>](#displaylanguage-string) | <span data-ttu-id="0366a-135">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-135">Compose</span></span><br><span data-ttu-id="0366a-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-136">Read</span></span> | <span data-ttu-id="0366a-137">String</span><span class="sxs-lookup"><span data-stu-id="0366a-137">String</span></span> | [<span data-ttu-id="0366a-138">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-138">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0366a-139">размещать</span><span class="sxs-lookup"><span data-stu-id="0366a-139">host</span></span>](#host-hosttype) | <span data-ttu-id="0366a-140">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-140">Compose</span></span><br><span data-ttu-id="0366a-141">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-141">Read</span></span> | [<span data-ttu-id="0366a-142">HostType</span><span class="sxs-lookup"><span data-stu-id="0366a-142">HostType</span></span>](/javascript/api/office/office.hosttype?view=outlook-js-preview) | [<span data-ttu-id="0366a-143">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-143">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0366a-144">mailbox</span><span class="sxs-lookup"><span data-stu-id="0366a-144">mailbox</span></span>](office.context.mailbox.md) | <span data-ttu-id="0366a-145">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-145">Compose</span></span><br><span data-ttu-id="0366a-146">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-146">Read</span></span> | [<span data-ttu-id="0366a-147">Mailbox</span><span class="sxs-lookup"><span data-stu-id="0366a-147">Mailbox</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview) | [<span data-ttu-id="0366a-148">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-148">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0366a-149">officeTheme</span><span class="sxs-lookup"><span data-stu-id="0366a-149">officeTheme</span></span>](#officetheme-officetheme) | <span data-ttu-id="0366a-150">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-150">Compose</span></span><br><span data-ttu-id="0366a-151">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-151">Read</span></span> | [<span data-ttu-id="0366a-152">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="0366a-152">OfficeTheme</span></span>](/javascript/api/office/office.officetheme?view=outlook-js-preview) | [<span data-ttu-id="0366a-153">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0366a-153">Preview</span></span>](../preview-requirement-set/outlook-requirement-set-preview.md) |
| [<span data-ttu-id="0366a-154">управляем</span><span class="sxs-lookup"><span data-stu-id="0366a-154">platform</span></span>](#platform-platformtype) | <span data-ttu-id="0366a-155">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-155">Compose</span></span><br><span data-ttu-id="0366a-156">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-156">Read</span></span> | [<span data-ttu-id="0366a-157">PlatformType</span><span class="sxs-lookup"><span data-stu-id="0366a-157">PlatformType</span></span>](/javascript/api/office/office.platformtype?view=outlook-js-preview) | [<span data-ttu-id="0366a-158">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-158">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0366a-159">потребность</span><span class="sxs-lookup"><span data-stu-id="0366a-159">requirements</span></span>](#requirements-requirementsetsupport) | <span data-ttu-id="0366a-160">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-160">Compose</span></span><br><span data-ttu-id="0366a-161">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-161">Read</span></span> | [<span data-ttu-id="0366a-162">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="0366a-162">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport?view=outlook-js-preview) | [<span data-ttu-id="0366a-163">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-163">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0366a-164">roamingSettings</span><span class="sxs-lookup"><span data-stu-id="0366a-164">roamingSettings</span></span>](#roamingsettings-roamingsettings) | <span data-ttu-id="0366a-165">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-165">Compose</span></span><br><span data-ttu-id="0366a-166">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-166">Read</span></span> | [<span data-ttu-id="0366a-167">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="0366a-167">RoamingSettings</span></span>](/javascript/api/outlook/office.roamingsettings?view=outlook-js-preview) | [<span data-ttu-id="0366a-168">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-168">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |
| [<span data-ttu-id="0366a-169">ui</span><span class="sxs-lookup"><span data-stu-id="0366a-169">ui</span></span>](#ui-ui) | <span data-ttu-id="0366a-170">Создание</span><span class="sxs-lookup"><span data-stu-id="0366a-170">Compose</span></span><br><span data-ttu-id="0366a-171">Чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-171">Read</span></span> | [<span data-ttu-id="0366a-172">UI</span><span class="sxs-lookup"><span data-stu-id="0366a-172">UI</span></span>](/javascript/api/office/office.ui?view=outlook-js-preview) | [<span data-ttu-id="0366a-173">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-173">1.1</span></span>](../requirement-set-1.1/outlook-requirement-set-1.1.md) |

## <a name="property-details"></a><span data-ttu-id="0366a-174">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="0366a-174">Property details</span></span>

#### <a name="auth-auth"></a><span data-ttu-id="0366a-175">Проверка подлинности: [AUTH](/javascript/api/office/office.auth)</span><span class="sxs-lookup"><span data-stu-id="0366a-175">auth: [Auth](/javascript/api/office/office.auth)</span></span>

<span data-ttu-id="0366a-176">Поддерживает [единый вход (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) , предоставляя метод, позволяющий ведущему приложению Office получать маркер доступа к веб-приложению надстройки.</span><span class="sxs-lookup"><span data-stu-id="0366a-176">Supports [single sign-on (SSO)](../../../outlook/authenticate-a-user-with-an-sso-token.md) by providing a method that allows the Office host to obtain an access token to the add-in's web application.</span></span> <span data-ttu-id="0366a-177">Косвенно это также дает возможность надстройке получать доступ к данным Microsoft Graph пользователя, вошедшего в систему, не требуя от пользователя еще раз выполнить вход в систему.</span><span class="sxs-lookup"><span data-stu-id="0366a-177">Indirectly, this also enables the add-in to access the signed-in user's Microsoft Graph data without requiring the user to sign in a second time.</span></span>

##### <a name="type"></a><span data-ttu-id="0366a-178">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-178">Type</span></span>

*   [<span data-ttu-id="0366a-179">Auth</span><span class="sxs-lookup"><span data-stu-id="0366a-179">Auth</span></span>](/javascript/api/office/office.auth)

##### <a name="requirements"></a><span data-ttu-id="0366a-180">Requirements</span><span class="sxs-lookup"><span data-stu-id="0366a-180">Requirements</span></span>

|<span data-ttu-id="0366a-181">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-181">Requirement</span></span>| <span data-ttu-id="0366a-182">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-182">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-183">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0366a-183">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-184">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0366a-184">Preview</span></span>|
|[<span data-ttu-id="0366a-185">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-185">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-186">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-186">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0366a-187">Пример</span><span class="sxs-lookup"><span data-stu-id="0366a-187">Example</span></span>

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

#### <a name="contentlanguage-string"></a><span data-ttu-id="0366a-188">contentLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="0366a-188">contentLanguage: String</span></span>

<span data-ttu-id="0366a-189">Получает языковой стандарт (язык), указанный пользователем для редактирования элемента.</span><span class="sxs-lookup"><span data-stu-id="0366a-189">Gets the locale (language) specified by the user for editing the item.</span></span>

<span data-ttu-id="0366a-190">`contentLanguage` Значение соответствует текущему **языковому** параметру редактирования, указанному с **параметрами > файлов > языке** в ведущем приложении Office.</span><span class="sxs-lookup"><span data-stu-id="0366a-190">The `contentLanguage` value reflects the current **Editing Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="0366a-191">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-191">Type</span></span>

*   <span data-ttu-id="0366a-192">String</span><span class="sxs-lookup"><span data-stu-id="0366a-192">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0366a-193">Требования</span><span class="sxs-lookup"><span data-stu-id="0366a-193">Requirements</span></span>

|<span data-ttu-id="0366a-194">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-194">Requirement</span></span>| <span data-ttu-id="0366a-195">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-195">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-196">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0366a-196">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-197">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-197">1.1</span></span>|
|[<span data-ttu-id="0366a-198">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-198">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-199">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-199">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0366a-200">Пример</span><span class="sxs-lookup"><span data-stu-id="0366a-200">Example</span></span>

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

#### <a name="diagnostics-contextinformation"></a><span data-ttu-id="0366a-201">Диагностика: [контекстинформатион](/javascript/api/office/office.contextinformation)</span><span class="sxs-lookup"><span data-stu-id="0366a-201">diagnostics: [ContextInformation](/javascript/api/office/office.contextinformation)</span></span>

<span data-ttu-id="0366a-202">Получает сведения о среде, в которой выполняется надстройка.</span><span class="sxs-lookup"><span data-stu-id="0366a-202">Gets information about the environment in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="0366a-203">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-203">Type</span></span>

*   [<span data-ttu-id="0366a-204">контекстинформатион</span><span class="sxs-lookup"><span data-stu-id="0366a-204">ContextInformation</span></span>](/javascript/api/office/office.contextinformation)

##### <a name="requirements"></a><span data-ttu-id="0366a-205">Requirements</span><span class="sxs-lookup"><span data-stu-id="0366a-205">Requirements</span></span>

|<span data-ttu-id="0366a-206">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-206">Requirement</span></span>| <span data-ttu-id="0366a-207">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-207">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-208">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0366a-208">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-209">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-209">1.1</span></span>|
|[<span data-ttu-id="0366a-210">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-210">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-211">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-211">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0366a-212">Пример</span><span class="sxs-lookup"><span data-stu-id="0366a-212">Example</span></span>

```js
console.log(JSON.stringify(Office.context.diagnostics));
```

<br>

---
---

#### <a name="displaylanguage-string"></a><span data-ttu-id="0366a-213">displayLanguage: строка</span><span class="sxs-lookup"><span data-stu-id="0366a-213">displayLanguage: String</span></span>

<span data-ttu-id="0366a-214">Получает определенный пользователем языковой стандарт (язык) в формате обозначений языка RFC 1766 для пользовательского интерфейса ведущего приложения Office.</span><span class="sxs-lookup"><span data-stu-id="0366a-214">Gets the locale (language) in RFC 1766 Language tag format specified by the user for the UI of the Office host application.</span></span>

<span data-ttu-id="0366a-215">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span><span class="sxs-lookup"><span data-stu-id="0366a-215">The `displayLanguage` value reflects the current **Display Language** setting specified with **File > Options > Language** in the Office host application.</span></span>

##### <a name="type"></a><span data-ttu-id="0366a-216">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-216">Type</span></span>

*   <span data-ttu-id="0366a-217">String</span><span class="sxs-lookup"><span data-stu-id="0366a-217">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="0366a-218">Требования</span><span class="sxs-lookup"><span data-stu-id="0366a-218">Requirements</span></span>

|<span data-ttu-id="0366a-219">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-219">Requirement</span></span>| <span data-ttu-id="0366a-220">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-220">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-221">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0366a-221">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-222">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-222">1.1</span></span>|
|[<span data-ttu-id="0366a-223">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-223">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-224">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-224">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0366a-225">Пример</span><span class="sxs-lookup"><span data-stu-id="0366a-225">Example</span></span>

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

#### <a name="host-hosttype"></a><span data-ttu-id="0366a-226">узел: [HostType](/javascript/api/office/office.hosttype)</span><span class="sxs-lookup"><span data-stu-id="0366a-226">host: [HostType](/javascript/api/office/office.hosttype)</span></span>

<span data-ttu-id="0366a-227">Получает узел приложений Office, в котором работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="0366a-227">Gets the Office application host in which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="0366a-228">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-228">Type</span></span>

*   [<span data-ttu-id="0366a-229">HostType</span><span class="sxs-lookup"><span data-stu-id="0366a-229">HostType</span></span>](/javascript/api/office/office.hosttype)

##### <a name="requirements"></a><span data-ttu-id="0366a-230">Requirements</span><span class="sxs-lookup"><span data-stu-id="0366a-230">Requirements</span></span>

|<span data-ttu-id="0366a-231">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-231">Requirement</span></span>| <span data-ttu-id="0366a-232">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-232">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-233">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0366a-233">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-234">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-234">1.1</span></span>|
|[<span data-ttu-id="0366a-235">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-235">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-236">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-236">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0366a-237">Пример</span><span class="sxs-lookup"><span data-stu-id="0366a-237">Example</span></span>

```js
console.log(JSON.stringify(Office.context.host));
```

<br>

---
---

#### <a name="officetheme-officetheme"></a><span data-ttu-id="0366a-238">officeTheme: [officeTheme](/javascript/api/office/office.officetheme)</span><span class="sxs-lookup"><span data-stu-id="0366a-238">officeTheme: [OfficeTheme](/javascript/api/office/office.officetheme)</span></span>

<span data-ttu-id="0366a-239">Предоставляет доступ к свойствам цветов темы Office.</span><span class="sxs-lookup"><span data-stu-id="0366a-239">Provides access to the properties for Office theme colors.</span></span>

> [!NOTE]
> <span data-ttu-id="0366a-240">Этот элемент поддерживается только в Outlook для Windows.</span><span class="sxs-lookup"><span data-stu-id="0366a-240">This member is only supported in Outlook on Windows.</span></span>

<span data-ttu-id="0366a-241">Использование цветов тем Office позволяет координировать цветовую схему надстройки с текущей темой Office, выбранной пользователем с помощью **файла > учетной записи office > пользовательского интерфейса темы**Office, которая применяется ко всем ведущим приложениям Office.</span><span class="sxs-lookup"><span data-stu-id="0366a-241">Using Office theme colors lets you coordinate the color scheme of your add-in with the current Office theme selected by the user with **File > Office Account > Office Theme UI**, which is applied across all Office host applications.</span></span> <span data-ttu-id="0366a-242">Using Office theme colors is appropriate for mail and task pane add-ins.</span><span class="sxs-lookup"><span data-stu-id="0366a-242">Using Office theme colors is appropriate for mail and task pane add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="0366a-243">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-243">Type</span></span>

*   [<span data-ttu-id="0366a-244">OfficeTheme</span><span class="sxs-lookup"><span data-stu-id="0366a-244">OfficeTheme</span></span>](/javascript/api/office/office.officetheme)

##### <a name="properties"></a><span data-ttu-id="0366a-245">Свойства:</span><span class="sxs-lookup"><span data-stu-id="0366a-245">Properties:</span></span>

|<span data-ttu-id="0366a-246">Имя</span><span class="sxs-lookup"><span data-stu-id="0366a-246">Name</span></span>| <span data-ttu-id="0366a-247">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-247">Type</span></span>| <span data-ttu-id="0366a-248">Описание</span><span class="sxs-lookup"><span data-stu-id="0366a-248">Description</span></span>|
|---|---|---|
|`bodyBackgroundColor`| <span data-ttu-id="0366a-249">String</span><span class="sxs-lookup"><span data-stu-id="0366a-249">String</span></span>|<span data-ttu-id="0366a-250">Получает цвет фона текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0366a-250">Gets the Office theme body background color as a hexadecimal color triplet.</span></span>|
|`bodyForegroundColor`| <span data-ttu-id="0366a-251">String</span><span class="sxs-lookup"><span data-stu-id="0366a-251">String</span></span>|<span data-ttu-id="0366a-252">Получает цвет переднего плана текста сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0366a-252">Gets the Office theme body foreground color as a hexadecimal color triplet.</span></span>|
|`controlBackgroundColor`| <span data-ttu-id="0366a-253">String</span><span class="sxs-lookup"><span data-stu-id="0366a-253">String</span></span>|<span data-ttu-id="0366a-254">Получает цвет фона элемента управления для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0366a-254">Gets the Office theme control background color as a hexadecimal color triplet.</span></span>|
|`controlForegroundColor`| <span data-ttu-id="0366a-255">String</span><span class="sxs-lookup"><span data-stu-id="0366a-255">String</span></span>|<span data-ttu-id="0366a-256">Получает цвет элемента управления текстом сообщения для темы Office в виде шестнадцатеричной триады цветов.</span><span class="sxs-lookup"><span data-stu-id="0366a-256">Gets the Office theme body control color as a hexadecimal color triplet.</span></span>|

##### <a name="requirements"></a><span data-ttu-id="0366a-257">Requirements</span><span class="sxs-lookup"><span data-stu-id="0366a-257">Requirements</span></span>

|<span data-ttu-id="0366a-258">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-258">Requirement</span></span>| <span data-ttu-id="0366a-259">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-259">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-260">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="0366a-260">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-261">Предварительная версия</span><span class="sxs-lookup"><span data-stu-id="0366a-261">Preview</span></span>|
|[<span data-ttu-id="0366a-262">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-262">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-263">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-263">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0366a-264">Пример</span><span class="sxs-lookup"><span data-stu-id="0366a-264">Example</span></span>

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

#### <a name="platform-platformtype"></a><span data-ttu-id="0366a-265">Платформа: [PlatformType](/javascript/api/office/office.platformtype)</span><span class="sxs-lookup"><span data-stu-id="0366a-265">platform: [PlatformType](/javascript/api/office/office.platformtype)</span></span>

<span data-ttu-id="0366a-266">Предоставляет платформу, на которой работает надстройка.</span><span class="sxs-lookup"><span data-stu-id="0366a-266">Provides the platform on which the add-in is running.</span></span>

##### <a name="type"></a><span data-ttu-id="0366a-267">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-267">Type</span></span>

*   [<span data-ttu-id="0366a-268">PlatformType</span><span class="sxs-lookup"><span data-stu-id="0366a-268">PlatformType</span></span>](/javascript/api/office/office.platformtype)

##### <a name="requirements"></a><span data-ttu-id="0366a-269">Requirements</span><span class="sxs-lookup"><span data-stu-id="0366a-269">Requirements</span></span>

|<span data-ttu-id="0366a-270">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-270">Requirement</span></span>| <span data-ttu-id="0366a-271">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-271">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-272">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0366a-272">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-273">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-273">1.1</span></span>|
|[<span data-ttu-id="0366a-274">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-274">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-275">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-275">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0366a-276">Пример</span><span class="sxs-lookup"><span data-stu-id="0366a-276">Example</span></span>

```js
console.log(JSON.stringify(Office.context.platform));
```

<br>

---
---

#### <a name="requirements-requirementsetsupport"></a><span data-ttu-id="0366a-277">требования: [рекуирементсетсуппорт](/javascript/api/office/office.requirementsetsupport)</span><span class="sxs-lookup"><span data-stu-id="0366a-277">requirements: [RequirementSetSupport](/javascript/api/office/office.requirementsetsupport)</span></span>

<span data-ttu-id="0366a-278">Предоставляет метод для определения наборов требований, поддерживаемых на текущем узле и платформе.</span><span class="sxs-lookup"><span data-stu-id="0366a-278">Provides a method for determining what requirement sets are supported on the current host and platform.</span></span>

##### <a name="type"></a><span data-ttu-id="0366a-279">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-279">Type</span></span>

*   [<span data-ttu-id="0366a-280">рекуирементсетсуппорт</span><span class="sxs-lookup"><span data-stu-id="0366a-280">RequirementSetSupport</span></span>](/javascript/api/office/office.requirementsetsupport)

##### <a name="requirements"></a><span data-ttu-id="0366a-281">Requirements</span><span class="sxs-lookup"><span data-stu-id="0366a-281">Requirements</span></span>

|<span data-ttu-id="0366a-282">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-282">Requirement</span></span>| <span data-ttu-id="0366a-283">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-283">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-284">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0366a-284">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-285">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-285">1.1</span></span>|
|[<span data-ttu-id="0366a-286">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-286">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-287">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-287">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="0366a-288">Пример</span><span class="sxs-lookup"><span data-stu-id="0366a-288">Example</span></span>

```js
console.log(JSON.stringify(Office.context.requirements.isSetSupported("mailbox", "1.1")));
```

<br>

---
---

#### <a name="roamingsettings-roamingsettings"></a><span data-ttu-id="0366a-289">roamingSettings: [roamingSettings](/javascript/api/outlook/office.roamingsettings)</span><span class="sxs-lookup"><span data-stu-id="0366a-289">roamingSettings: [RoamingSettings](/javascript/api/outlook/office.roamingsettings)</span></span>

<span data-ttu-id="0366a-290">Получает объект, представляющий настраиваемые параметры или состояние надстройки почты, сохраненное в почтовом ящике пользователя.</span><span class="sxs-lookup"><span data-stu-id="0366a-290">Gets an object that represents the custom settings or state of a mail add-in saved to a user's mailbox.</span></span>

<span data-ttu-id="0366a-291">Объект `RoamingSettings` позволяет сохранять данные для надстройки почты, записанные в почтовом ящике пользователя, и получать к ним доступ, таким образом делая их доступными для этой надстройки, когда она запускается из любого клиентского ведущего приложения, используемого для доступа к этому почтовому ящику.</span><span class="sxs-lookup"><span data-stu-id="0366a-291">The `RoamingSettings` object lets you store and access data for a mail add-in that is stored in a user's mailbox, so that is available to that add-in when it is running from any host client application used to access that mailbox.</span></span>

##### <a name="type"></a><span data-ttu-id="0366a-292">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-292">Type</span></span>

*   [<span data-ttu-id="0366a-293">RoamingSettings</span><span class="sxs-lookup"><span data-stu-id="0366a-293">RoamingSettings</span></span>](/javascript/api/outlook/office.RoamingSettings)

##### <a name="requirements"></a><span data-ttu-id="0366a-294">Requirements</span><span class="sxs-lookup"><span data-stu-id="0366a-294">Requirements</span></span>

|<span data-ttu-id="0366a-295">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-295">Requirement</span></span>| <span data-ttu-id="0366a-296">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-296">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-297">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0366a-297">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-298">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-298">1.1</span></span>|
|[<span data-ttu-id="0366a-299">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="0366a-299">Minimum permission level</span></span>](../../../outlook/understanding-outlook-add-in-permissions.md)| <span data-ttu-id="0366a-300">С ограничениями</span><span class="sxs-lookup"><span data-stu-id="0366a-300">Restricted</span></span>|
|[<span data-ttu-id="0366a-301">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-301">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-302">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-302">Compose or Read</span></span>|

<br>

---
---

#### <a name="ui-ui"></a><span data-ttu-id="0366a-303">Пользовательский интерфейс: [Пользовательский интерфейс](/javascript/api/office/office.ui)</span><span class="sxs-lookup"><span data-stu-id="0366a-303">ui: [UI](/javascript/api/office/office.ui)</span></span>

<span data-ttu-id="0366a-304">Предоставляет объекты и методы, которые можно использовать для создания компонентов пользовательского интерфейса, например диалоговых окон, в надстройках Office и управления ими.</span><span class="sxs-lookup"><span data-stu-id="0366a-304">Provides objects and methods that you can use to create and manipulate UI components, such as dialog boxes, in your Office Add-ins.</span></span>

##### <a name="type"></a><span data-ttu-id="0366a-305">Тип</span><span class="sxs-lookup"><span data-stu-id="0366a-305">Type</span></span>

*   [<span data-ttu-id="0366a-306">UI</span><span class="sxs-lookup"><span data-stu-id="0366a-306">UI</span></span>](/javascript/api/office/office.ui)

##### <a name="requirements"></a><span data-ttu-id="0366a-307">Requirements</span><span class="sxs-lookup"><span data-stu-id="0366a-307">Requirements</span></span>

|<span data-ttu-id="0366a-308">Требование</span><span class="sxs-lookup"><span data-stu-id="0366a-308">Requirement</span></span>| <span data-ttu-id="0366a-309">Значение</span><span class="sxs-lookup"><span data-stu-id="0366a-309">Value</span></span>|
|---|---|
|[<span data-ttu-id="0366a-310">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="0366a-310">Minimum mailbox requirement set version</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)| <span data-ttu-id="0366a-311">1.1</span><span class="sxs-lookup"><span data-stu-id="0366a-311">1.1</span></span>|
|[<span data-ttu-id="0366a-312">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="0366a-312">Applicable Outlook mode</span></span>](../../../outlook/outlook-add-ins-overview.md#extension-points)| <span data-ttu-id="0366a-313">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="0366a-313">Compose or Read</span></span>|
