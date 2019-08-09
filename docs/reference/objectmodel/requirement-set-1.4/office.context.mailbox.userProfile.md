---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.4
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7a728ebbec0136e0b2eddfb4402e45abe3f02ad4
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268637"
---
# <a name="userprofile"></a><span data-ttu-id="ce918-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ce918-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ce918-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ce918-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce918-104">Требования</span><span class="sxs-lookup"><span data-stu-id="ce918-104">Requirements</span></span>

|<span data-ttu-id="ce918-105">Требование</span><span class="sxs-lookup"><span data-stu-id="ce918-105">Requirement</span></span>| <span data-ttu-id="ce918-106">Значение</span><span class="sxs-lookup"><span data-stu-id="ce918-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce918-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce918-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce918-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ce918-108">1.0</span></span>|
|[<span data-ttu-id="ce918-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce918-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce918-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce918-110">ReadItem</span></span>|
|[<span data-ttu-id="ce918-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce918-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce918-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce918-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ce918-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="ce918-113">Members and methods</span></span>

| <span data-ttu-id="ce918-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="ce918-114">Member</span></span> | <span data-ttu-id="ce918-115">Тип</span><span class="sxs-lookup"><span data-stu-id="ce918-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ce918-116">displayName</span><span class="sxs-lookup"><span data-stu-id="ce918-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="ce918-117">Member</span><span class="sxs-lookup"><span data-stu-id="ce918-117">Member</span></span> |
| [<span data-ttu-id="ce918-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="ce918-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="ce918-119">Member</span><span class="sxs-lookup"><span data-stu-id="ce918-119">Member</span></span> |
| [<span data-ttu-id="ce918-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="ce918-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="ce918-121">Member</span><span class="sxs-lookup"><span data-stu-id="ce918-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ce918-122">Members</span><span class="sxs-lookup"><span data-stu-id="ce918-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="ce918-123">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="ce918-123">displayName: String</span></span>

<span data-ttu-id="ce918-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="ce918-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ce918-125">Тип</span><span class="sxs-lookup"><span data-stu-id="ce918-125">Type</span></span>

*   <span data-ttu-id="ce918-126">String</span><span class="sxs-lookup"><span data-stu-id="ce918-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce918-127">Требования</span><span class="sxs-lookup"><span data-stu-id="ce918-127">Requirements</span></span>

|<span data-ttu-id="ce918-128">Требование</span><span class="sxs-lookup"><span data-stu-id="ce918-128">Requirement</span></span>| <span data-ttu-id="ce918-129">Значение</span><span class="sxs-lookup"><span data-stu-id="ce918-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce918-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce918-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce918-131">1.0</span><span class="sxs-lookup"><span data-stu-id="ce918-131">1.0</span></span>|
|[<span data-ttu-id="ce918-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce918-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce918-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce918-133">ReadItem</span></span>|
|[<span data-ttu-id="ce918-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce918-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce918-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce918-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce918-136">Пример</span><span class="sxs-lookup"><span data-stu-id="ce918-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="ce918-137">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="ce918-137">emailAddress: String</span></span>

<span data-ttu-id="ce918-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="ce918-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ce918-139">Тип</span><span class="sxs-lookup"><span data-stu-id="ce918-139">Type</span></span>

*   <span data-ttu-id="ce918-140">String</span><span class="sxs-lookup"><span data-stu-id="ce918-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce918-141">Требования</span><span class="sxs-lookup"><span data-stu-id="ce918-141">Requirements</span></span>

|<span data-ttu-id="ce918-142">Требование</span><span class="sxs-lookup"><span data-stu-id="ce918-142">Requirement</span></span>| <span data-ttu-id="ce918-143">Значение</span><span class="sxs-lookup"><span data-stu-id="ce918-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce918-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce918-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce918-145">1.0</span><span class="sxs-lookup"><span data-stu-id="ce918-145">1.0</span></span>|
|[<span data-ttu-id="ce918-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce918-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce918-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce918-147">ReadItem</span></span>|
|[<span data-ttu-id="ce918-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce918-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce918-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce918-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce918-150">Пример</span><span class="sxs-lookup"><span data-stu-id="ce918-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="ce918-151">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="ce918-151">timeZone: String</span></span>

<span data-ttu-id="ce918-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ce918-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ce918-153">Тип</span><span class="sxs-lookup"><span data-stu-id="ce918-153">Type</span></span>

*   <span data-ttu-id="ce918-154">String</span><span class="sxs-lookup"><span data-stu-id="ce918-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ce918-155">Требования</span><span class="sxs-lookup"><span data-stu-id="ce918-155">Requirements</span></span>

|<span data-ttu-id="ce918-156">Требование</span><span class="sxs-lookup"><span data-stu-id="ce918-156">Requirement</span></span>| <span data-ttu-id="ce918-157">Значение</span><span class="sxs-lookup"><span data-stu-id="ce918-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="ce918-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ce918-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ce918-159">1.0</span><span class="sxs-lookup"><span data-stu-id="ce918-159">1.0</span></span>|
|[<span data-ttu-id="ce918-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ce918-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ce918-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ce918-161">ReadItem</span></span>|
|[<span data-ttu-id="ce918-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ce918-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ce918-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ce918-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ce918-164">Пример</span><span class="sxs-lookup"><span data-stu-id="ce918-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
