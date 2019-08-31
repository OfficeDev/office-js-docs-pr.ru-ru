---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 06492623e0b9ab16792d6b23dfaeb27d99125ff1
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696402"
---
# <a name="userprofile"></a><span data-ttu-id="88fe7-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="88fe7-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="88fe7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="88fe7-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="88fe7-104">Требования</span><span class="sxs-lookup"><span data-stu-id="88fe7-104">Requirements</span></span>

|<span data-ttu-id="88fe7-105">Требование</span><span class="sxs-lookup"><span data-stu-id="88fe7-105">Requirement</span></span>| <span data-ttu-id="88fe7-106">Значение</span><span class="sxs-lookup"><span data-stu-id="88fe7-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="88fe7-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="88fe7-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88fe7-108">1.0</span><span class="sxs-lookup"><span data-stu-id="88fe7-108">1.0</span></span>|
|[<span data-ttu-id="88fe7-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="88fe7-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88fe7-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88fe7-110">ReadItem</span></span>|
|[<span data-ttu-id="88fe7-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="88fe7-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="88fe7-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="88fe7-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="88fe7-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="88fe7-113">Members and methods</span></span>

| <span data-ttu-id="88fe7-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="88fe7-114">Member</span></span> | <span data-ttu-id="88fe7-115">Тип</span><span class="sxs-lookup"><span data-stu-id="88fe7-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="88fe7-116">displayName</span><span class="sxs-lookup"><span data-stu-id="88fe7-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="88fe7-117">Member</span><span class="sxs-lookup"><span data-stu-id="88fe7-117">Member</span></span> |
| [<span data-ttu-id="88fe7-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="88fe7-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="88fe7-119">Member</span><span class="sxs-lookup"><span data-stu-id="88fe7-119">Member</span></span> |
| [<span data-ttu-id="88fe7-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="88fe7-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="88fe7-121">Member</span><span class="sxs-lookup"><span data-stu-id="88fe7-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="88fe7-122">Members</span><span class="sxs-lookup"><span data-stu-id="88fe7-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="88fe7-123">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="88fe7-123">displayName: String</span></span>

<span data-ttu-id="88fe7-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="88fe7-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="88fe7-125">Тип</span><span class="sxs-lookup"><span data-stu-id="88fe7-125">Type</span></span>

*   <span data-ttu-id="88fe7-126">String</span><span class="sxs-lookup"><span data-stu-id="88fe7-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="88fe7-127">Требования</span><span class="sxs-lookup"><span data-stu-id="88fe7-127">Requirements</span></span>

|<span data-ttu-id="88fe7-128">Требование</span><span class="sxs-lookup"><span data-stu-id="88fe7-128">Requirement</span></span>| <span data-ttu-id="88fe7-129">Значение</span><span class="sxs-lookup"><span data-stu-id="88fe7-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="88fe7-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="88fe7-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88fe7-131">1.0</span><span class="sxs-lookup"><span data-stu-id="88fe7-131">1.0</span></span>|
|[<span data-ttu-id="88fe7-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="88fe7-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88fe7-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88fe7-133">ReadItem</span></span>|
|[<span data-ttu-id="88fe7-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="88fe7-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="88fe7-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="88fe7-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88fe7-136">Пример</span><span class="sxs-lookup"><span data-stu-id="88fe7-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="88fe7-137">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="88fe7-137">emailAddress: String</span></span>

<span data-ttu-id="88fe7-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="88fe7-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="88fe7-139">Тип</span><span class="sxs-lookup"><span data-stu-id="88fe7-139">Type</span></span>

*   <span data-ttu-id="88fe7-140">String</span><span class="sxs-lookup"><span data-stu-id="88fe7-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="88fe7-141">Требования</span><span class="sxs-lookup"><span data-stu-id="88fe7-141">Requirements</span></span>

|<span data-ttu-id="88fe7-142">Требование</span><span class="sxs-lookup"><span data-stu-id="88fe7-142">Requirement</span></span>| <span data-ttu-id="88fe7-143">Значение</span><span class="sxs-lookup"><span data-stu-id="88fe7-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="88fe7-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="88fe7-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88fe7-145">1.0</span><span class="sxs-lookup"><span data-stu-id="88fe7-145">1.0</span></span>|
|[<span data-ttu-id="88fe7-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="88fe7-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88fe7-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88fe7-147">ReadItem</span></span>|
|[<span data-ttu-id="88fe7-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="88fe7-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="88fe7-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="88fe7-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88fe7-150">Пример</span><span class="sxs-lookup"><span data-stu-id="88fe7-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="88fe7-151">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="88fe7-151">timeZone: String</span></span>

<span data-ttu-id="88fe7-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="88fe7-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="88fe7-153">Тип</span><span class="sxs-lookup"><span data-stu-id="88fe7-153">Type</span></span>

*   <span data-ttu-id="88fe7-154">String</span><span class="sxs-lookup"><span data-stu-id="88fe7-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="88fe7-155">Требования</span><span class="sxs-lookup"><span data-stu-id="88fe7-155">Requirements</span></span>

|<span data-ttu-id="88fe7-156">Требование</span><span class="sxs-lookup"><span data-stu-id="88fe7-156">Requirement</span></span>| <span data-ttu-id="88fe7-157">Значение</span><span class="sxs-lookup"><span data-stu-id="88fe7-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="88fe7-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="88fe7-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="88fe7-159">1.0</span><span class="sxs-lookup"><span data-stu-id="88fe7-159">1.0</span></span>|
|[<span data-ttu-id="88fe7-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="88fe7-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="88fe7-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="88fe7-161">ReadItem</span></span>|
|[<span data-ttu-id="88fe7-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="88fe7-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="88fe7-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="88fe7-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="88fe7-164">Пример</span><span class="sxs-lookup"><span data-stu-id="88fe7-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
