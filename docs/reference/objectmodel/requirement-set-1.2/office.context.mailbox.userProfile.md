---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 8ba2a21b16c51c827155d793241b80c5c510dd5a
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696339"
---
# <a name="userprofile"></a><span data-ttu-id="d463d-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="d463d-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="d463d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="d463d-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="d463d-104">Требования</span><span class="sxs-lookup"><span data-stu-id="d463d-104">Requirements</span></span>

|<span data-ttu-id="d463d-105">Требование</span><span class="sxs-lookup"><span data-stu-id="d463d-105">Requirement</span></span>| <span data-ttu-id="d463d-106">Значение</span><span class="sxs-lookup"><span data-stu-id="d463d-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="d463d-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d463d-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d463d-108">1.0</span><span class="sxs-lookup"><span data-stu-id="d463d-108">1.0</span></span>|
|[<span data-ttu-id="d463d-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d463d-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d463d-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d463d-110">ReadItem</span></span>|
|[<span data-ttu-id="d463d-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d463d-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d463d-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d463d-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="d463d-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="d463d-113">Members and methods</span></span>

| <span data-ttu-id="d463d-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="d463d-114">Member</span></span> | <span data-ttu-id="d463d-115">Тип</span><span class="sxs-lookup"><span data-stu-id="d463d-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="d463d-116">displayName</span><span class="sxs-lookup"><span data-stu-id="d463d-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="d463d-117">Member</span><span class="sxs-lookup"><span data-stu-id="d463d-117">Member</span></span> |
| [<span data-ttu-id="d463d-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="d463d-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="d463d-119">Member</span><span class="sxs-lookup"><span data-stu-id="d463d-119">Member</span></span> |
| [<span data-ttu-id="d463d-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="d463d-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="d463d-121">Member</span><span class="sxs-lookup"><span data-stu-id="d463d-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="d463d-122">Members</span><span class="sxs-lookup"><span data-stu-id="d463d-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="d463d-123">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="d463d-123">displayName: String</span></span>

<span data-ttu-id="d463d-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="d463d-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="d463d-125">Тип</span><span class="sxs-lookup"><span data-stu-id="d463d-125">Type</span></span>

*   <span data-ttu-id="d463d-126">String</span><span class="sxs-lookup"><span data-stu-id="d463d-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d463d-127">Требования</span><span class="sxs-lookup"><span data-stu-id="d463d-127">Requirements</span></span>

|<span data-ttu-id="d463d-128">Требование</span><span class="sxs-lookup"><span data-stu-id="d463d-128">Requirement</span></span>| <span data-ttu-id="d463d-129">Значение</span><span class="sxs-lookup"><span data-stu-id="d463d-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="d463d-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d463d-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d463d-131">1.0</span><span class="sxs-lookup"><span data-stu-id="d463d-131">1.0</span></span>|
|[<span data-ttu-id="d463d-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d463d-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d463d-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d463d-133">ReadItem</span></span>|
|[<span data-ttu-id="d463d-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d463d-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d463d-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d463d-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d463d-136">Пример</span><span class="sxs-lookup"><span data-stu-id="d463d-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="d463d-137">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="d463d-137">emailAddress: String</span></span>

<span data-ttu-id="d463d-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="d463d-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="d463d-139">Тип</span><span class="sxs-lookup"><span data-stu-id="d463d-139">Type</span></span>

*   <span data-ttu-id="d463d-140">String</span><span class="sxs-lookup"><span data-stu-id="d463d-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d463d-141">Требования</span><span class="sxs-lookup"><span data-stu-id="d463d-141">Requirements</span></span>

|<span data-ttu-id="d463d-142">Требование</span><span class="sxs-lookup"><span data-stu-id="d463d-142">Requirement</span></span>| <span data-ttu-id="d463d-143">Значение</span><span class="sxs-lookup"><span data-stu-id="d463d-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="d463d-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d463d-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d463d-145">1.0</span><span class="sxs-lookup"><span data-stu-id="d463d-145">1.0</span></span>|
|[<span data-ttu-id="d463d-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d463d-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d463d-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d463d-147">ReadItem</span></span>|
|[<span data-ttu-id="d463d-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d463d-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d463d-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d463d-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d463d-150">Пример</span><span class="sxs-lookup"><span data-stu-id="d463d-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="d463d-151">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="d463d-151">timeZone: String</span></span>

<span data-ttu-id="d463d-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="d463d-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="d463d-153">Тип</span><span class="sxs-lookup"><span data-stu-id="d463d-153">Type</span></span>

*   <span data-ttu-id="d463d-154">String</span><span class="sxs-lookup"><span data-stu-id="d463d-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="d463d-155">Требования</span><span class="sxs-lookup"><span data-stu-id="d463d-155">Requirements</span></span>

|<span data-ttu-id="d463d-156">Требование</span><span class="sxs-lookup"><span data-stu-id="d463d-156">Requirement</span></span>| <span data-ttu-id="d463d-157">Значение</span><span class="sxs-lookup"><span data-stu-id="d463d-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="d463d-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="d463d-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="d463d-159">1.0</span><span class="sxs-lookup"><span data-stu-id="d463d-159">1.0</span></span>|
|[<span data-ttu-id="d463d-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="d463d-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="d463d-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="d463d-161">ReadItem</span></span>|
|[<span data-ttu-id="d463d-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="d463d-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="d463d-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="d463d-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="d463d-164">Пример</span><span class="sxs-lookup"><span data-stu-id="d463d-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
