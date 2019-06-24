---
title: Office. Context. Mailbox. userProfile — Предварительная версия набора требований
description: ''
ms.date: 06/20/2019
localization_priority: Normal
ms.openlocfilehash: 051615f28db420ad1d198e21d9a9eeb1cad3fe7a
ms.sourcegitcommit: 382e2735a1295da914f2bfc38883e518070cec61
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/21/2019
ms.locfileid: "35127529"
---
# <a name="userprofile"></a><span data-ttu-id="8bf65-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="8bf65-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="8bf65-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="8bf65-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="8bf65-104">Требования</span><span class="sxs-lookup"><span data-stu-id="8bf65-104">Requirements</span></span>

|<span data-ttu-id="8bf65-105">Требование</span><span class="sxs-lookup"><span data-stu-id="8bf65-105">Requirement</span></span>| <span data-ttu-id="8bf65-106">Значение</span><span class="sxs-lookup"><span data-stu-id="8bf65-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bf65-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8bf65-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bf65-108">1.0</span><span class="sxs-lookup"><span data-stu-id="8bf65-108">1.0</span></span>|
|[<span data-ttu-id="8bf65-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8bf65-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bf65-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bf65-110">ReadItem</span></span>|
|[<span data-ttu-id="8bf65-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8bf65-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8bf65-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8bf65-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8bf65-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="8bf65-113">Members and methods</span></span>

| <span data-ttu-id="8bf65-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="8bf65-114">Member</span></span> | <span data-ttu-id="8bf65-115">Тип</span><span class="sxs-lookup"><span data-stu-id="8bf65-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8bf65-116">accountType</span><span class="sxs-lookup"><span data-stu-id="8bf65-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="8bf65-117">Member</span><span class="sxs-lookup"><span data-stu-id="8bf65-117">Member</span></span> |
| [<span data-ttu-id="8bf65-118">displayName</span><span class="sxs-lookup"><span data-stu-id="8bf65-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="8bf65-119">Member</span><span class="sxs-lookup"><span data-stu-id="8bf65-119">Member</span></span> |
| [<span data-ttu-id="8bf65-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="8bf65-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="8bf65-121">Member</span><span class="sxs-lookup"><span data-stu-id="8bf65-121">Member</span></span> |
| [<span data-ttu-id="8bf65-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="8bf65-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="8bf65-123">Member</span><span class="sxs-lookup"><span data-stu-id="8bf65-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="8bf65-124">Members</span><span class="sxs-lookup"><span data-stu-id="8bf65-124">Members</span></span>

#### <a name="accounttype-string"></a><span data-ttu-id="8bf65-125">accountType: строка</span><span class="sxs-lookup"><span data-stu-id="8bf65-125">accountType: String</span></span>

> [!NOTE]
> <span data-ttu-id="8bf65-126">В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии в Mac (сборка 16.9.1212 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="8bf65-126">This member is currently only supported in Outlook 2016 or later on Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="8bf65-127">Возвращает тип учетной записи пользователя, связанного с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="8bf65-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="8bf65-128">Возможные значения перечислены в приведенной ниже таблице.</span><span class="sxs-lookup"><span data-stu-id="8bf65-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="8bf65-129">Значение</span><span class="sxs-lookup"><span data-stu-id="8bf65-129">Value</span></span> | <span data-ttu-id="8bf65-130">Описание</span><span class="sxs-lookup"><span data-stu-id="8bf65-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="8bf65-131">Почтовый ящик находится на локальном сервере Exchange.</span><span class="sxs-lookup"><span data-stu-id="8bf65-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="8bf65-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="8bf65-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="8bf65-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="8bf65-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="8bf65-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="8bf65-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="8bf65-135">Тип</span><span class="sxs-lookup"><span data-stu-id="8bf65-135">Type</span></span>

*   <span data-ttu-id="8bf65-136">String</span><span class="sxs-lookup"><span data-stu-id="8bf65-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8bf65-137">Требования</span><span class="sxs-lookup"><span data-stu-id="8bf65-137">Requirements</span></span>

|<span data-ttu-id="8bf65-138">Требование</span><span class="sxs-lookup"><span data-stu-id="8bf65-138">Requirement</span></span>| <span data-ttu-id="8bf65-139">Значение</span><span class="sxs-lookup"><span data-stu-id="8bf65-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bf65-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8bf65-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bf65-141">1.6</span><span class="sxs-lookup"><span data-stu-id="8bf65-141">1.6</span></span> |
|[<span data-ttu-id="8bf65-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8bf65-142">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bf65-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bf65-143">ReadItem</span></span>|
|[<span data-ttu-id="8bf65-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8bf65-144">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8bf65-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8bf65-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bf65-146">Пример</span><span class="sxs-lookup"><span data-stu-id="8bf65-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

---
---

#### <a name="displayname-string"></a><span data-ttu-id="8bf65-147">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="8bf65-147">displayName: String</span></span>

<span data-ttu-id="8bf65-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="8bf65-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="8bf65-149">Тип</span><span class="sxs-lookup"><span data-stu-id="8bf65-149">Type</span></span>

*   <span data-ttu-id="8bf65-150">String</span><span class="sxs-lookup"><span data-stu-id="8bf65-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8bf65-151">Требования</span><span class="sxs-lookup"><span data-stu-id="8bf65-151">Requirements</span></span>

|<span data-ttu-id="8bf65-152">Требование</span><span class="sxs-lookup"><span data-stu-id="8bf65-152">Requirement</span></span>| <span data-ttu-id="8bf65-153">Значение</span><span class="sxs-lookup"><span data-stu-id="8bf65-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bf65-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8bf65-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bf65-155">1.0</span><span class="sxs-lookup"><span data-stu-id="8bf65-155">1.0</span></span>|
|[<span data-ttu-id="8bf65-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8bf65-156">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bf65-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bf65-157">ReadItem</span></span>|
|[<span data-ttu-id="8bf65-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8bf65-158">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8bf65-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8bf65-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bf65-160">Пример</span><span class="sxs-lookup"><span data-stu-id="8bf65-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="8bf65-161">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="8bf65-161">emailAddress: String</span></span>

<span data-ttu-id="8bf65-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="8bf65-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="8bf65-163">Тип</span><span class="sxs-lookup"><span data-stu-id="8bf65-163">Type</span></span>

*   <span data-ttu-id="8bf65-164">String</span><span class="sxs-lookup"><span data-stu-id="8bf65-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8bf65-165">Требования</span><span class="sxs-lookup"><span data-stu-id="8bf65-165">Requirements</span></span>

|<span data-ttu-id="8bf65-166">Требование</span><span class="sxs-lookup"><span data-stu-id="8bf65-166">Requirement</span></span>| <span data-ttu-id="8bf65-167">Значение</span><span class="sxs-lookup"><span data-stu-id="8bf65-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bf65-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8bf65-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bf65-169">1.0</span><span class="sxs-lookup"><span data-stu-id="8bf65-169">1.0</span></span>|
|[<span data-ttu-id="8bf65-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8bf65-170">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bf65-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bf65-171">ReadItem</span></span>|
|[<span data-ttu-id="8bf65-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8bf65-172">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8bf65-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8bf65-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bf65-174">Пример</span><span class="sxs-lookup"><span data-stu-id="8bf65-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

---
---

#### <a name="timezone-string"></a><span data-ttu-id="8bf65-175">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="8bf65-175">timeZone: String</span></span>

<span data-ttu-id="8bf65-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="8bf65-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="8bf65-177">Тип</span><span class="sxs-lookup"><span data-stu-id="8bf65-177">Type</span></span>

*   <span data-ttu-id="8bf65-178">String</span><span class="sxs-lookup"><span data-stu-id="8bf65-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8bf65-179">Требования</span><span class="sxs-lookup"><span data-stu-id="8bf65-179">Requirements</span></span>

|<span data-ttu-id="8bf65-180">Требование</span><span class="sxs-lookup"><span data-stu-id="8bf65-180">Requirement</span></span>| <span data-ttu-id="8bf65-181">Значение</span><span class="sxs-lookup"><span data-stu-id="8bf65-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="8bf65-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8bf65-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8bf65-183">1.0</span><span class="sxs-lookup"><span data-stu-id="8bf65-183">1.0</span></span>|
|[<span data-ttu-id="8bf65-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8bf65-184">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8bf65-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8bf65-185">ReadItem</span></span>|
|[<span data-ttu-id="8bf65-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8bf65-186">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8bf65-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8bf65-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8bf65-188">Пример</span><span class="sxs-lookup"><span data-stu-id="8bf65-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
