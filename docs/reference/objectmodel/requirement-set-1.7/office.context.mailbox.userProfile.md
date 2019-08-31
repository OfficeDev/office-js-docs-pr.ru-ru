---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,7
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 8e33d40bec9b561c642ad6e0da73ae13a18378b6
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36695905"
---
# <a name="userprofile"></a><span data-ttu-id="cb9eb-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="cb9eb-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="cb9eb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="cb9eb-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb9eb-104">Требования</span><span class="sxs-lookup"><span data-stu-id="cb9eb-104">Requirements</span></span>

|<span data-ttu-id="cb9eb-105">Требование</span><span class="sxs-lookup"><span data-stu-id="cb9eb-105">Requirement</span></span>| <span data-ttu-id="cb9eb-106">Значение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb9eb-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cb9eb-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb9eb-108">1.0</span><span class="sxs-lookup"><span data-stu-id="cb9eb-108">1.0</span></span>|
|[<span data-ttu-id="cb9eb-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cb9eb-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb9eb-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb9eb-110">ReadItem</span></span>|
|[<span data-ttu-id="cb9eb-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cb9eb-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cb9eb-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="cb9eb-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="cb9eb-113">Members and methods</span></span>

| <span data-ttu-id="cb9eb-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="cb9eb-114">Member</span></span> | <span data-ttu-id="cb9eb-115">Тип</span><span class="sxs-lookup"><span data-stu-id="cb9eb-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="cb9eb-116">accountType</span><span class="sxs-lookup"><span data-stu-id="cb9eb-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="cb9eb-117">Member</span><span class="sxs-lookup"><span data-stu-id="cb9eb-117">Member</span></span> |
| [<span data-ttu-id="cb9eb-118">displayName</span><span class="sxs-lookup"><span data-stu-id="cb9eb-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="cb9eb-119">Member</span><span class="sxs-lookup"><span data-stu-id="cb9eb-119">Member</span></span> |
| [<span data-ttu-id="cb9eb-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="cb9eb-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="cb9eb-121">Member</span><span class="sxs-lookup"><span data-stu-id="cb9eb-121">Member</span></span> |
| [<span data-ttu-id="cb9eb-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="cb9eb-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="cb9eb-123">Member</span><span class="sxs-lookup"><span data-stu-id="cb9eb-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="cb9eb-124">Members</span><span class="sxs-lookup"><span data-stu-id="cb9eb-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="cb9eb-125">accountType: строка</span><span class="sxs-lookup"><span data-stu-id="cb9eb-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="cb9eb-126">В настоящее время этот элемент поддерживается только Outlook 2016 или более поздней версии в Mac (сборка 16.9.1212 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="cb9eb-126">This member is currently only supported by Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="cb9eb-127">Возвращает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="cb9eb-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="cb9eb-128">Возможные значения перечислены в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="cb9eb-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="cb9eb-129">Значение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-129">Value</span></span> | <span data-ttu-id="cb9eb-130">Описание</span><span class="sxs-lookup"><span data-stu-id="cb9eb-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="cb9eb-131">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="cb9eb-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="cb9eb-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="cb9eb-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="cb9eb-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="cb9eb-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="cb9eb-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="cb9eb-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="cb9eb-135">Тип</span><span class="sxs-lookup"><span data-stu-id="cb9eb-135">Type</span></span>

*   <span data-ttu-id="cb9eb-136">String</span><span class="sxs-lookup"><span data-stu-id="cb9eb-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb9eb-137">Требования</span><span class="sxs-lookup"><span data-stu-id="cb9eb-137">Requirements</span></span>

|<span data-ttu-id="cb9eb-138">Требование</span><span class="sxs-lookup"><span data-stu-id="cb9eb-138">Requirement</span></span>| <span data-ttu-id="cb9eb-139">Значение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb9eb-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="cb9eb-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb9eb-141">1.6</span><span class="sxs-lookup"><span data-stu-id="cb9eb-141">1.6</span></span> |
|[<span data-ttu-id="cb9eb-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cb9eb-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb9eb-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb9eb-143">ReadItem</span></span>|
|[<span data-ttu-id="cb9eb-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cb9eb-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cb9eb-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb9eb-146">Пример</span><span class="sxs-lookup"><span data-stu-id="cb9eb-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

<br>

---
---

#### <a name="displayname-string"></a><span data-ttu-id="cb9eb-147">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="cb9eb-147">displayName: String</span></span>

<span data-ttu-id="cb9eb-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="cb9eb-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="cb9eb-149">Тип</span><span class="sxs-lookup"><span data-stu-id="cb9eb-149">Type</span></span>

*   <span data-ttu-id="cb9eb-150">String</span><span class="sxs-lookup"><span data-stu-id="cb9eb-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb9eb-151">Требования</span><span class="sxs-lookup"><span data-stu-id="cb9eb-151">Requirements</span></span>

|<span data-ttu-id="cb9eb-152">Требование</span><span class="sxs-lookup"><span data-stu-id="cb9eb-152">Requirement</span></span>| <span data-ttu-id="cb9eb-153">Значение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb9eb-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cb9eb-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb9eb-155">1.0</span><span class="sxs-lookup"><span data-stu-id="cb9eb-155">1.0</span></span>|
|[<span data-ttu-id="cb9eb-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cb9eb-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb9eb-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb9eb-157">ReadItem</span></span>|
|[<span data-ttu-id="cb9eb-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cb9eb-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cb9eb-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb9eb-160">Пример</span><span class="sxs-lookup"><span data-stu-id="cb9eb-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="cb9eb-161">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="cb9eb-161">emailAddress: String</span></span>

<span data-ttu-id="cb9eb-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="cb9eb-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="cb9eb-163">Тип</span><span class="sxs-lookup"><span data-stu-id="cb9eb-163">Type</span></span>

*   <span data-ttu-id="cb9eb-164">String</span><span class="sxs-lookup"><span data-stu-id="cb9eb-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb9eb-165">Требования</span><span class="sxs-lookup"><span data-stu-id="cb9eb-165">Requirements</span></span>

|<span data-ttu-id="cb9eb-166">Требование</span><span class="sxs-lookup"><span data-stu-id="cb9eb-166">Requirement</span></span>| <span data-ttu-id="cb9eb-167">Значение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb9eb-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cb9eb-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb9eb-169">1.0</span><span class="sxs-lookup"><span data-stu-id="cb9eb-169">1.0</span></span>|
|[<span data-ttu-id="cb9eb-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cb9eb-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb9eb-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb9eb-171">ReadItem</span></span>|
|[<span data-ttu-id="cb9eb-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cb9eb-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cb9eb-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb9eb-174">Пример</span><span class="sxs-lookup"><span data-stu-id="cb9eb-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="cb9eb-175">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="cb9eb-175">timeZone: String</span></span>

<span data-ttu-id="cb9eb-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="cb9eb-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="cb9eb-177">Тип</span><span class="sxs-lookup"><span data-stu-id="cb9eb-177">Type</span></span>

*   <span data-ttu-id="cb9eb-178">String</span><span class="sxs-lookup"><span data-stu-id="cb9eb-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="cb9eb-179">Требования</span><span class="sxs-lookup"><span data-stu-id="cb9eb-179">Requirements</span></span>

|<span data-ttu-id="cb9eb-180">Требование</span><span class="sxs-lookup"><span data-stu-id="cb9eb-180">Requirement</span></span>| <span data-ttu-id="cb9eb-181">Значение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="cb9eb-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="cb9eb-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="cb9eb-183">1.0</span><span class="sxs-lookup"><span data-stu-id="cb9eb-183">1.0</span></span>|
|[<span data-ttu-id="cb9eb-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="cb9eb-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="cb9eb-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="cb9eb-185">ReadItem</span></span>|
|[<span data-ttu-id="cb9eb-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="cb9eb-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="cb9eb-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="cb9eb-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="cb9eb-188">Пример</span><span class="sxs-lookup"><span data-stu-id="cb9eb-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
