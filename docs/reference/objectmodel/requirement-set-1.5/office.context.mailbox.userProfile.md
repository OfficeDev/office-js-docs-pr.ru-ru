---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: fc20497cc8df8d091ba0195f7dca9b283ff4d1c2
ms.sourcegitcommit: a2950492a2337de3180b713f5693fe82dbdd6a17
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/27/2019
ms.locfileid: "30871019"
---
# <a name="userprofile"></a><span data-ttu-id="ff49a-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ff49a-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ff49a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ff49a-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff49a-104">Требования</span><span class="sxs-lookup"><span data-stu-id="ff49a-104">Requirements</span></span>

|<span data-ttu-id="ff49a-105">Требование</span><span class="sxs-lookup"><span data-stu-id="ff49a-105">Requirement</span></span>| <span data-ttu-id="ff49a-106">Значение</span><span class="sxs-lookup"><span data-stu-id="ff49a-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff49a-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ff49a-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff49a-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ff49a-108">1.0</span></span>|
|[<span data-ttu-id="ff49a-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ff49a-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff49a-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff49a-110">ReadItem</span></span>|
|[<span data-ttu-id="ff49a-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ff49a-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff49a-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ff49a-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ff49a-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="ff49a-113">Members and methods</span></span>

| <span data-ttu-id="ff49a-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="ff49a-114">Member</span></span> | <span data-ttu-id="ff49a-115">Тип</span><span class="sxs-lookup"><span data-stu-id="ff49a-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ff49a-116">displayName</span><span class="sxs-lookup"><span data-stu-id="ff49a-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="ff49a-117">Member</span><span class="sxs-lookup"><span data-stu-id="ff49a-117">Member</span></span> |
| [<span data-ttu-id="ff49a-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="ff49a-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="ff49a-119">Member</span><span class="sxs-lookup"><span data-stu-id="ff49a-119">Member</span></span> |
| [<span data-ttu-id="ff49a-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="ff49a-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="ff49a-121">Member</span><span class="sxs-lookup"><span data-stu-id="ff49a-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ff49a-122">Элементы</span><span class="sxs-lookup"><span data-stu-id="ff49a-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="ff49a-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ff49a-123">displayName :String</span></span>

<span data-ttu-id="ff49a-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="ff49a-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ff49a-125">Тип</span><span class="sxs-lookup"><span data-stu-id="ff49a-125">Type</span></span>

*   <span data-ttu-id="ff49a-126">String</span><span class="sxs-lookup"><span data-stu-id="ff49a-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff49a-127">Требования</span><span class="sxs-lookup"><span data-stu-id="ff49a-127">Requirements</span></span>

|<span data-ttu-id="ff49a-128">Требование</span><span class="sxs-lookup"><span data-stu-id="ff49a-128">Requirement</span></span>| <span data-ttu-id="ff49a-129">Значение</span><span class="sxs-lookup"><span data-stu-id="ff49a-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff49a-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ff49a-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff49a-131">1.0</span><span class="sxs-lookup"><span data-stu-id="ff49a-131">1.0</span></span>|
|[<span data-ttu-id="ff49a-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ff49a-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff49a-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff49a-133">ReadItem</span></span>|
|[<span data-ttu-id="ff49a-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ff49a-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff49a-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ff49a-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff49a-136">Пример</span><span class="sxs-lookup"><span data-stu-id="ff49a-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ff49a-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ff49a-137">emailAddress :String</span></span>

<span data-ttu-id="ff49a-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="ff49a-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ff49a-139">Тип</span><span class="sxs-lookup"><span data-stu-id="ff49a-139">Type</span></span>

*   <span data-ttu-id="ff49a-140">String</span><span class="sxs-lookup"><span data-stu-id="ff49a-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff49a-141">Требования</span><span class="sxs-lookup"><span data-stu-id="ff49a-141">Requirements</span></span>

|<span data-ttu-id="ff49a-142">Требование</span><span class="sxs-lookup"><span data-stu-id="ff49a-142">Requirement</span></span>| <span data-ttu-id="ff49a-143">Значение</span><span class="sxs-lookup"><span data-stu-id="ff49a-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff49a-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ff49a-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff49a-145">1.0</span><span class="sxs-lookup"><span data-stu-id="ff49a-145">1.0</span></span>|
|[<span data-ttu-id="ff49a-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ff49a-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff49a-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff49a-147">ReadItem</span></span>|
|[<span data-ttu-id="ff49a-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ff49a-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff49a-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ff49a-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff49a-150">Пример</span><span class="sxs-lookup"><span data-stu-id="ff49a-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ff49a-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ff49a-151">timeZone :String</span></span>

<span data-ttu-id="ff49a-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ff49a-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ff49a-153">Тип</span><span class="sxs-lookup"><span data-stu-id="ff49a-153">Type</span></span>

*   <span data-ttu-id="ff49a-154">String</span><span class="sxs-lookup"><span data-stu-id="ff49a-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ff49a-155">Требования</span><span class="sxs-lookup"><span data-stu-id="ff49a-155">Requirements</span></span>

|<span data-ttu-id="ff49a-156">Требование</span><span class="sxs-lookup"><span data-stu-id="ff49a-156">Requirement</span></span>| <span data-ttu-id="ff49a-157">Значение</span><span class="sxs-lookup"><span data-stu-id="ff49a-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="ff49a-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ff49a-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ff49a-159">1.0</span><span class="sxs-lookup"><span data-stu-id="ff49a-159">1.0</span></span>|
|[<span data-ttu-id="ff49a-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ff49a-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ff49a-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ff49a-161">ReadItem</span></span>|
|[<span data-ttu-id="ff49a-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ff49a-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="ff49a-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ff49a-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ff49a-164">Пример</span><span class="sxs-lookup"><span data-stu-id="ff49a-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
