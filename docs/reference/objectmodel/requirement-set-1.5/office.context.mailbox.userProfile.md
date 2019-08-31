---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,5
description: ''
ms.date: 04/24/2019
localization_priority: Normal
ms.openlocfilehash: 993fad674fcc616483ac927619e7ca64d81b7326
ms.sourcegitcommit: 1fb99b1b4e63868a0e81a928c69a34c42bf7e209
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/30/2019
ms.locfileid: "36696094"
---
# <a name="userprofile"></a><span data-ttu-id="8a558-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="8a558-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="8a558-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="8a558-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a558-104">Требования</span><span class="sxs-lookup"><span data-stu-id="8a558-104">Requirements</span></span>

|<span data-ttu-id="8a558-105">Требование</span><span class="sxs-lookup"><span data-stu-id="8a558-105">Requirement</span></span>| <span data-ttu-id="8a558-106">Значение</span><span class="sxs-lookup"><span data-stu-id="8a558-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a558-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8a558-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a558-108">1.0</span><span class="sxs-lookup"><span data-stu-id="8a558-108">1.0</span></span>|
|[<span data-ttu-id="8a558-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8a558-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a558-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a558-110">ReadItem</span></span>|
|[<span data-ttu-id="8a558-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8a558-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a558-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8a558-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8a558-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="8a558-113">Members and methods</span></span>

| <span data-ttu-id="8a558-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="8a558-114">Member</span></span> | <span data-ttu-id="8a558-115">Тип</span><span class="sxs-lookup"><span data-stu-id="8a558-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8a558-116">displayName</span><span class="sxs-lookup"><span data-stu-id="8a558-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="8a558-117">Member</span><span class="sxs-lookup"><span data-stu-id="8a558-117">Member</span></span> |
| [<span data-ttu-id="8a558-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="8a558-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="8a558-119">Member</span><span class="sxs-lookup"><span data-stu-id="8a558-119">Member</span></span> |
| [<span data-ttu-id="8a558-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="8a558-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="8a558-121">Member</span><span class="sxs-lookup"><span data-stu-id="8a558-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="8a558-122">Members</span><span class="sxs-lookup"><span data-stu-id="8a558-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="8a558-123">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="8a558-123">displayName: String</span></span>

<span data-ttu-id="8a558-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="8a558-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="8a558-125">Тип</span><span class="sxs-lookup"><span data-stu-id="8a558-125">Type</span></span>

*   <span data-ttu-id="8a558-126">String</span><span class="sxs-lookup"><span data-stu-id="8a558-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a558-127">Требования</span><span class="sxs-lookup"><span data-stu-id="8a558-127">Requirements</span></span>

|<span data-ttu-id="8a558-128">Требование</span><span class="sxs-lookup"><span data-stu-id="8a558-128">Requirement</span></span>| <span data-ttu-id="8a558-129">Значение</span><span class="sxs-lookup"><span data-stu-id="8a558-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a558-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8a558-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a558-131">1.0</span><span class="sxs-lookup"><span data-stu-id="8a558-131">1.0</span></span>|
|[<span data-ttu-id="8a558-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8a558-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a558-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a558-133">ReadItem</span></span>|
|[<span data-ttu-id="8a558-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8a558-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a558-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8a558-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a558-136">Пример</span><span class="sxs-lookup"><span data-stu-id="8a558-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

<br>

---
---

#### <a name="emailaddress-string"></a><span data-ttu-id="8a558-137">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="8a558-137">emailAddress: String</span></span>

<span data-ttu-id="8a558-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="8a558-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="8a558-139">Тип</span><span class="sxs-lookup"><span data-stu-id="8a558-139">Type</span></span>

*   <span data-ttu-id="8a558-140">String</span><span class="sxs-lookup"><span data-stu-id="8a558-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a558-141">Требования</span><span class="sxs-lookup"><span data-stu-id="8a558-141">Requirements</span></span>

|<span data-ttu-id="8a558-142">Требование</span><span class="sxs-lookup"><span data-stu-id="8a558-142">Requirement</span></span>| <span data-ttu-id="8a558-143">Значение</span><span class="sxs-lookup"><span data-stu-id="8a558-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a558-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8a558-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a558-145">1.0</span><span class="sxs-lookup"><span data-stu-id="8a558-145">1.0</span></span>|
|[<span data-ttu-id="8a558-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8a558-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a558-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a558-147">ReadItem</span></span>|
|[<span data-ttu-id="8a558-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8a558-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a558-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8a558-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a558-150">Пример</span><span class="sxs-lookup"><span data-stu-id="8a558-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

<br>

---
---

#### <a name="timezone-string"></a><span data-ttu-id="8a558-151">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="8a558-151">timeZone: String</span></span>

<span data-ttu-id="8a558-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="8a558-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="8a558-153">Тип</span><span class="sxs-lookup"><span data-stu-id="8a558-153">Type</span></span>

*   <span data-ttu-id="8a558-154">String</span><span class="sxs-lookup"><span data-stu-id="8a558-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8a558-155">Требования</span><span class="sxs-lookup"><span data-stu-id="8a558-155">Requirements</span></span>

|<span data-ttu-id="8a558-156">Требование</span><span class="sxs-lookup"><span data-stu-id="8a558-156">Requirement</span></span>| <span data-ttu-id="8a558-157">Значение</span><span class="sxs-lookup"><span data-stu-id="8a558-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="8a558-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8a558-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8a558-159">1.0</span><span class="sxs-lookup"><span data-stu-id="8a558-159">1.0</span></span>|
|[<span data-ttu-id="8a558-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8a558-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8a558-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8a558-161">ReadItem</span></span>|
|[<span data-ttu-id="8a558-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8a558-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="8a558-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8a558-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8a558-164">Пример</span><span class="sxs-lookup"><span data-stu-id="8a558-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
