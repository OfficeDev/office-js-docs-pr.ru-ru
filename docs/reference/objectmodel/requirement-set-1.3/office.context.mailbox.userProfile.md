---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 20393d0ac650de34054b912d9e53a9ac167fddb2
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696255"
---
# <a name="userprofile"></a><span data-ttu-id="11c91-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="11c91-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="11c91-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="11c91-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="11c91-104">Требования</span><span class="sxs-lookup"><span data-stu-id="11c91-104">Requirements</span></span>

|<span data-ttu-id="11c91-105">Требование</span><span class="sxs-lookup"><span data-stu-id="11c91-105">Requirement</span></span>| <span data-ttu-id="11c91-106">Значение</span><span class="sxs-lookup"><span data-stu-id="11c91-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="11c91-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="11c91-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11c91-108">1.0</span><span class="sxs-lookup"><span data-stu-id="11c91-108">1.0</span></span>|
|[<span data-ttu-id="11c91-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="11c91-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11c91-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11c91-110">ReadItem</span></span>|
|[<span data-ttu-id="11c91-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="11c91-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11c91-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="11c91-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="11c91-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="11c91-113">Members and methods</span></span>

| <span data-ttu-id="11c91-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="11c91-114">Member</span></span> | <span data-ttu-id="11c91-115">Тип</span><span class="sxs-lookup"><span data-stu-id="11c91-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="11c91-116">displayName</span><span class="sxs-lookup"><span data-stu-id="11c91-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="11c91-117">Member</span><span class="sxs-lookup"><span data-stu-id="11c91-117">Member</span></span> |
| [<span data-ttu-id="11c91-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="11c91-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="11c91-119">Member</span><span class="sxs-lookup"><span data-stu-id="11c91-119">Member</span></span> |
| [<span data-ttu-id="11c91-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="11c91-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="11c91-121">Member</span><span class="sxs-lookup"><span data-stu-id="11c91-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="11c91-122">Members</span><span class="sxs-lookup"><span data-stu-id="11c91-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="11c91-123">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="11c91-123">displayName: String</span></span>

<span data-ttu-id="11c91-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="11c91-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="11c91-125">Тип</span><span class="sxs-lookup"><span data-stu-id="11c91-125">Type</span></span>

*   <span data-ttu-id="11c91-126">String</span><span class="sxs-lookup"><span data-stu-id="11c91-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="11c91-127">Требования</span><span class="sxs-lookup"><span data-stu-id="11c91-127">Requirements</span></span>

|<span data-ttu-id="11c91-128">Требование</span><span class="sxs-lookup"><span data-stu-id="11c91-128">Requirement</span></span>| <span data-ttu-id="11c91-129">Значение</span><span class="sxs-lookup"><span data-stu-id="11c91-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="11c91-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="11c91-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11c91-131">1.0</span><span class="sxs-lookup"><span data-stu-id="11c91-131">1.0</span></span>|
|[<span data-ttu-id="11c91-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="11c91-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11c91-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11c91-133">ReadItem</span></span>|
|[<span data-ttu-id="11c91-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="11c91-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11c91-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="11c91-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="11c91-136">Пример</span><span class="sxs-lookup"><span data-stu-id="11c91-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="11c91-137">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="11c91-137">emailAddress: String</span></span>

<span data-ttu-id="11c91-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="11c91-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="11c91-139">Тип</span><span class="sxs-lookup"><span data-stu-id="11c91-139">Type</span></span>

*   <span data-ttu-id="11c91-140">String</span><span class="sxs-lookup"><span data-stu-id="11c91-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="11c91-141">Требования</span><span class="sxs-lookup"><span data-stu-id="11c91-141">Requirements</span></span>

|<span data-ttu-id="11c91-142">Требование</span><span class="sxs-lookup"><span data-stu-id="11c91-142">Requirement</span></span>| <span data-ttu-id="11c91-143">Значение</span><span class="sxs-lookup"><span data-stu-id="11c91-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="11c91-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="11c91-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11c91-145">1.0</span><span class="sxs-lookup"><span data-stu-id="11c91-145">1.0</span></span>|
|[<span data-ttu-id="11c91-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="11c91-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11c91-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11c91-147">ReadItem</span></span>|
|[<span data-ttu-id="11c91-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="11c91-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11c91-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="11c91-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="11c91-150">Пример</span><span class="sxs-lookup"><span data-stu-id="11c91-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="11c91-151">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="11c91-151">timeZone: String</span></span>

<span data-ttu-id="11c91-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="11c91-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="11c91-153">Тип</span><span class="sxs-lookup"><span data-stu-id="11c91-153">Type</span></span>

*   <span data-ttu-id="11c91-154">String</span><span class="sxs-lookup"><span data-stu-id="11c91-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="11c91-155">Требования</span><span class="sxs-lookup"><span data-stu-id="11c91-155">Requirements</span></span>

|<span data-ttu-id="11c91-156">Требование</span><span class="sxs-lookup"><span data-stu-id="11c91-156">Requirement</span></span>| <span data-ttu-id="11c91-157">Значение</span><span class="sxs-lookup"><span data-stu-id="11c91-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="11c91-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="11c91-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="11c91-159">1.0</span><span class="sxs-lookup"><span data-stu-id="11c91-159">1.0</span></span>|
|[<span data-ttu-id="11c91-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="11c91-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="11c91-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="11c91-161">ReadItem</span></span>|
|[<span data-ttu-id="11c91-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="11c91-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="11c91-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="11c91-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="11c91-164">Пример</span><span class="sxs-lookup"><span data-stu-id="11c91-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
