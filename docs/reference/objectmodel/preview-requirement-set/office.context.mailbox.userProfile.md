---
title: Office. Context. Mailbox. userProfile — Предварительная версия набора требований
description: ''
ms.date: 11/25/2019
localization_priority: Normal
ms.openlocfilehash: 4afc64f247155576ab3f0024d1929a29a0f7dc0c
ms.sourcegitcommit: 05a883a7fd89136301ce35aabc57638e9f563288
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/27/2019
ms.locfileid: "39629260"
---
# <a name="userprofile"></a><span data-ttu-id="f4a74-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="f4a74-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="f4a74-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="f4a74-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4a74-104">Requirements</span><span class="sxs-lookup"><span data-stu-id="f4a74-104">Requirements</span></span>

|<span data-ttu-id="f4a74-105">Требование</span><span class="sxs-lookup"><span data-stu-id="f4a74-105">Requirement</span></span>| <span data-ttu-id="f4a74-106">Значение</span><span class="sxs-lookup"><span data-stu-id="f4a74-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4a74-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f4a74-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4a74-108">1.0</span><span class="sxs-lookup"><span data-stu-id="f4a74-108">1.0</span></span>|
|[<span data-ttu-id="f4a74-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f4a74-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f4a74-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4a74-110">ReadItem</span></span>|
|[<span data-ttu-id="f4a74-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f4a74-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4a74-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f4a74-112">Compose or Read</span></span>|

##### <a name="properties"></a><span data-ttu-id="f4a74-113">Свойства</span><span class="sxs-lookup"><span data-stu-id="f4a74-113">Properties</span></span>

| <span data-ttu-id="f4a74-114">Свойство</span><span class="sxs-lookup"><span data-stu-id="f4a74-114">Property</span></span> | <span data-ttu-id="f4a74-115">Минимальные</span><span class="sxs-lookup"><span data-stu-id="f4a74-115">Minimum</span></span><br><span data-ttu-id="f4a74-116">уровень разрешения</span><span class="sxs-lookup"><span data-stu-id="f4a74-116">permission level</span></span> | <span data-ttu-id="f4a74-117">Способов</span><span class="sxs-lookup"><span data-stu-id="f4a74-117">Modes</span></span> | <span data-ttu-id="f4a74-118">Тип возвращаемых данных</span><span class="sxs-lookup"><span data-stu-id="f4a74-118">Return type</span></span> | <span data-ttu-id="f4a74-119">Минимальные</span><span class="sxs-lookup"><span data-stu-id="f4a74-119">Minimum</span></span><br><span data-ttu-id="f4a74-120">набор требований</span><span class="sxs-lookup"><span data-stu-id="f4a74-120">requirement set</span></span> |
|---|---|---|---|---|
| [<span data-ttu-id="f4a74-121">accountType</span><span class="sxs-lookup"><span data-stu-id="f4a74-121">accountType</span></span>](#accounttype-string) | <span data-ttu-id="f4a74-122">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4a74-122">ReadItem</span></span> | <span data-ttu-id="f4a74-123">Создание</span><span class="sxs-lookup"><span data-stu-id="f4a74-123">Compose</span></span><br><span data-ttu-id="f4a74-124">Чтение</span><span class="sxs-lookup"><span data-stu-id="f4a74-124">Read</span></span> | <span data-ttu-id="f4a74-125">String</span><span class="sxs-lookup"><span data-stu-id="f4a74-125">String</span></span> | <span data-ttu-id="f4a74-126">1.6</span><span class="sxs-lookup"><span data-stu-id="f4a74-126">1.6</span></span> |
| [<span data-ttu-id="f4a74-127">displayName</span><span class="sxs-lookup"><span data-stu-id="f4a74-127">displayName</span></span>](#displayname-string) | <span data-ttu-id="f4a74-128">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4a74-128">ReadItem</span></span> | <span data-ttu-id="f4a74-129">Создание</span><span class="sxs-lookup"><span data-stu-id="f4a74-129">Compose</span></span><br><span data-ttu-id="f4a74-130">Чтение</span><span class="sxs-lookup"><span data-stu-id="f4a74-130">Read</span></span> | <span data-ttu-id="f4a74-131">String</span><span class="sxs-lookup"><span data-stu-id="f4a74-131">String</span></span> | <span data-ttu-id="f4a74-132">1.0</span><span class="sxs-lookup"><span data-stu-id="f4a74-132">1.0</span></span> |
| [<span data-ttu-id="f4a74-133">emailAddress</span><span class="sxs-lookup"><span data-stu-id="f4a74-133">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="f4a74-134">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4a74-134">ReadItem</span></span> | <span data-ttu-id="f4a74-135">Создание</span><span class="sxs-lookup"><span data-stu-id="f4a74-135">Compose</span></span><br><span data-ttu-id="f4a74-136">Чтение</span><span class="sxs-lookup"><span data-stu-id="f4a74-136">Read</span></span> | <span data-ttu-id="f4a74-137">String</span><span class="sxs-lookup"><span data-stu-id="f4a74-137">String</span></span> | <span data-ttu-id="f4a74-138">1.0</span><span class="sxs-lookup"><span data-stu-id="f4a74-138">1.0</span></span> |
| [<span data-ttu-id="f4a74-139">timeZone</span><span class="sxs-lookup"><span data-stu-id="f4a74-139">timeZone</span></span>](#timezone-string) | <span data-ttu-id="f4a74-140">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4a74-140">ReadItem</span></span> | <span data-ttu-id="f4a74-141">Создание</span><span class="sxs-lookup"><span data-stu-id="f4a74-141">Compose</span></span><br><span data-ttu-id="f4a74-142">Чтение</span><span class="sxs-lookup"><span data-stu-id="f4a74-142">Read</span></span> | <span data-ttu-id="f4a74-143">String</span><span class="sxs-lookup"><span data-stu-id="f4a74-143">String</span></span> | <span data-ttu-id="f4a74-144">1.0</span><span class="sxs-lookup"><span data-stu-id="f4a74-144">1.0</span></span> |

## <a name="property-details"></a><span data-ttu-id="f4a74-145">Сведения о свойстве</span><span class="sxs-lookup"><span data-stu-id="f4a74-145">Property details</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="f4a74-146">accountType: строка</span><span class="sxs-lookup"><span data-stu-id="f4a74-146">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="f4a74-147">В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии в Mac (сборка 16.9.1212 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="f4a74-147">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="f4a74-148">Возвращает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="f4a74-148">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="f4a74-149">Возможные значения перечислены в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="f4a74-149">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="f4a74-150">Значение</span><span class="sxs-lookup"><span data-stu-id="f4a74-150">Value</span></span> | <span data-ttu-id="f4a74-151">Описание</span><span class="sxs-lookup"><span data-stu-id="f4a74-151">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="f4a74-152">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="f4a74-152">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="f4a74-153">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="f4a74-153">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="f4a74-154">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="f4a74-154">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="f4a74-155">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="f4a74-155">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="f4a74-156">Тип</span><span class="sxs-lookup"><span data-stu-id="f4a74-156">Type</span></span>

*   <span data-ttu-id="f4a74-157">String</span><span class="sxs-lookup"><span data-stu-id="f4a74-157">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4a74-158">Требования</span><span class="sxs-lookup"><span data-stu-id="f4a74-158">Requirements</span></span>

|<span data-ttu-id="f4a74-159">Требование</span><span class="sxs-lookup"><span data-stu-id="f4a74-159">Requirement</span></span>| <span data-ttu-id="f4a74-160">Значение</span><span class="sxs-lookup"><span data-stu-id="f4a74-160">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4a74-161">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="f4a74-161">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4a74-162">1.6</span><span class="sxs-lookup"><span data-stu-id="f4a74-162">1.6</span></span> |
|[<span data-ttu-id="f4a74-163">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f4a74-163">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f4a74-164">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4a74-164">ReadItem</span></span>|
|[<span data-ttu-id="f4a74-165">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f4a74-165">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4a74-166">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f4a74-166">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f4a74-167">Пример</span><span class="sxs-lookup"><span data-stu-id="f4a74-167">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="f4a74-168">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="f4a74-168">displayName: String</span></span>

<span data-ttu-id="f4a74-169">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="f4a74-169">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="f4a74-170">Тип</span><span class="sxs-lookup"><span data-stu-id="f4a74-170">Type</span></span>

*   <span data-ttu-id="f4a74-171">String</span><span class="sxs-lookup"><span data-stu-id="f4a74-171">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4a74-172">Требования</span><span class="sxs-lookup"><span data-stu-id="f4a74-172">Requirements</span></span>

|<span data-ttu-id="f4a74-173">Требование</span><span class="sxs-lookup"><span data-stu-id="f4a74-173">Requirement</span></span>| <span data-ttu-id="f4a74-174">Значение</span><span class="sxs-lookup"><span data-stu-id="f4a74-174">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4a74-175">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f4a74-175">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4a74-176">1.0</span><span class="sxs-lookup"><span data-stu-id="f4a74-176">1.0</span></span>|
|[<span data-ttu-id="f4a74-177">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f4a74-177">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f4a74-178">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4a74-178">ReadItem</span></span>|
|[<span data-ttu-id="f4a74-179">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f4a74-179">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4a74-180">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f4a74-180">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f4a74-181">Пример</span><span class="sxs-lookup"><span data-stu-id="f4a74-181">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="f4a74-182">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="f4a74-182">emailAddress: String</span></span>

<span data-ttu-id="f4a74-183">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="f4a74-183">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="f4a74-184">Тип</span><span class="sxs-lookup"><span data-stu-id="f4a74-184">Type</span></span>

*   <span data-ttu-id="f4a74-185">String</span><span class="sxs-lookup"><span data-stu-id="f4a74-185">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4a74-186">Требования</span><span class="sxs-lookup"><span data-stu-id="f4a74-186">Requirements</span></span>

|<span data-ttu-id="f4a74-187">Требование</span><span class="sxs-lookup"><span data-stu-id="f4a74-187">Requirement</span></span>| <span data-ttu-id="f4a74-188">Значение</span><span class="sxs-lookup"><span data-stu-id="f4a74-188">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4a74-189">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f4a74-189">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4a74-190">1.0</span><span class="sxs-lookup"><span data-stu-id="f4a74-190">1.0</span></span>|
|[<span data-ttu-id="f4a74-191">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f4a74-191">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f4a74-192">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4a74-192">ReadItem</span></span>|
|[<span data-ttu-id="f4a74-193">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f4a74-193">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4a74-194">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f4a74-194">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f4a74-195">Пример</span><span class="sxs-lookup"><span data-stu-id="f4a74-195">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="f4a74-196">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="f4a74-196">timeZone: String</span></span>

<span data-ttu-id="f4a74-197">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="f4a74-197">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="f4a74-198">Тип</span><span class="sxs-lookup"><span data-stu-id="f4a74-198">Type</span></span>

*   <span data-ttu-id="f4a74-199">String</span><span class="sxs-lookup"><span data-stu-id="f4a74-199">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="f4a74-200">Требования</span><span class="sxs-lookup"><span data-stu-id="f4a74-200">Requirements</span></span>

|<span data-ttu-id="f4a74-201">Требование</span><span class="sxs-lookup"><span data-stu-id="f4a74-201">Requirement</span></span>| <span data-ttu-id="f4a74-202">Значение</span><span class="sxs-lookup"><span data-stu-id="f4a74-202">Value</span></span>|
|---|---|
|[<span data-ttu-id="f4a74-203">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="f4a74-203">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="f4a74-204">1.0</span><span class="sxs-lookup"><span data-stu-id="f4a74-204">1.0</span></span>|
|[<span data-ttu-id="f4a74-205">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="f4a74-205">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="f4a74-206">ReadItem</span><span class="sxs-lookup"><span data-stu-id="f4a74-206">ReadItem</span></span>|
|[<span data-ttu-id="f4a74-207">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="f4a74-207">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="f4a74-208">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="f4a74-208">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="f4a74-209">Пример</span><span class="sxs-lookup"><span data-stu-id="f4a74-209">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
