---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,5
description: ''
ms.date: 03/19/2019
localization_priority: Normal
ms.openlocfilehash: fc20497cc8df8d091ba0195f7dca9b283ff4d1c2
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32451824"
---
# <a name="userprofile"></a><span data-ttu-id="bdec1-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="bdec1-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="bdec1-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="bdec1-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="bdec1-104">Требования</span><span class="sxs-lookup"><span data-stu-id="bdec1-104">Requirements</span></span>

|<span data-ttu-id="bdec1-105">Требование</span><span class="sxs-lookup"><span data-stu-id="bdec1-105">Requirement</span></span>| <span data-ttu-id="bdec1-106">Значение</span><span class="sxs-lookup"><span data-stu-id="bdec1-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="bdec1-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bdec1-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bdec1-108">1.0</span><span class="sxs-lookup"><span data-stu-id="bdec1-108">1.0</span></span>|
|[<span data-ttu-id="bdec1-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bdec1-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bdec1-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bdec1-110">ReadItem</span></span>|
|[<span data-ttu-id="bdec1-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bdec1-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bdec1-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bdec1-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="bdec1-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="bdec1-113">Members and methods</span></span>

| <span data-ttu-id="bdec1-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="bdec1-114">Member</span></span> | <span data-ttu-id="bdec1-115">Тип</span><span class="sxs-lookup"><span data-stu-id="bdec1-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="bdec1-116">displayName</span><span class="sxs-lookup"><span data-stu-id="bdec1-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="bdec1-117">Member</span><span class="sxs-lookup"><span data-stu-id="bdec1-117">Member</span></span> |
| [<span data-ttu-id="bdec1-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="bdec1-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="bdec1-119">Member</span><span class="sxs-lookup"><span data-stu-id="bdec1-119">Member</span></span> |
| [<span data-ttu-id="bdec1-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="bdec1-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="bdec1-121">Member</span><span class="sxs-lookup"><span data-stu-id="bdec1-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="bdec1-122">Элементы</span><span class="sxs-lookup"><span data-stu-id="bdec1-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="bdec1-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="bdec1-123">displayName :String</span></span>

<span data-ttu-id="bdec1-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="bdec1-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="bdec1-125">Тип</span><span class="sxs-lookup"><span data-stu-id="bdec1-125">Type</span></span>

*   <span data-ttu-id="bdec1-126">String</span><span class="sxs-lookup"><span data-stu-id="bdec1-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bdec1-127">Требования</span><span class="sxs-lookup"><span data-stu-id="bdec1-127">Requirements</span></span>

|<span data-ttu-id="bdec1-128">Требование</span><span class="sxs-lookup"><span data-stu-id="bdec1-128">Requirement</span></span>| <span data-ttu-id="bdec1-129">Значение</span><span class="sxs-lookup"><span data-stu-id="bdec1-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="bdec1-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bdec1-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bdec1-131">1.0</span><span class="sxs-lookup"><span data-stu-id="bdec1-131">1.0</span></span>|
|[<span data-ttu-id="bdec1-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bdec1-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bdec1-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bdec1-133">ReadItem</span></span>|
|[<span data-ttu-id="bdec1-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bdec1-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bdec1-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bdec1-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bdec1-136">Пример</span><span class="sxs-lookup"><span data-stu-id="bdec1-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="bdec1-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="bdec1-137">emailAddress :String</span></span>

<span data-ttu-id="bdec1-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="bdec1-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="bdec1-139">Тип</span><span class="sxs-lookup"><span data-stu-id="bdec1-139">Type</span></span>

*   <span data-ttu-id="bdec1-140">String</span><span class="sxs-lookup"><span data-stu-id="bdec1-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bdec1-141">Требования</span><span class="sxs-lookup"><span data-stu-id="bdec1-141">Requirements</span></span>

|<span data-ttu-id="bdec1-142">Требование</span><span class="sxs-lookup"><span data-stu-id="bdec1-142">Requirement</span></span>| <span data-ttu-id="bdec1-143">Значение</span><span class="sxs-lookup"><span data-stu-id="bdec1-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="bdec1-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bdec1-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bdec1-145">1.0</span><span class="sxs-lookup"><span data-stu-id="bdec1-145">1.0</span></span>|
|[<span data-ttu-id="bdec1-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bdec1-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bdec1-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bdec1-147">ReadItem</span></span>|
|[<span data-ttu-id="bdec1-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bdec1-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bdec1-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bdec1-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bdec1-150">Пример</span><span class="sxs-lookup"><span data-stu-id="bdec1-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="bdec1-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="bdec1-151">timeZone :String</span></span>

<span data-ttu-id="bdec1-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="bdec1-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="bdec1-153">Тип</span><span class="sxs-lookup"><span data-stu-id="bdec1-153">Type</span></span>

*   <span data-ttu-id="bdec1-154">String</span><span class="sxs-lookup"><span data-stu-id="bdec1-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="bdec1-155">Требования</span><span class="sxs-lookup"><span data-stu-id="bdec1-155">Requirements</span></span>

|<span data-ttu-id="bdec1-156">Требование</span><span class="sxs-lookup"><span data-stu-id="bdec1-156">Requirement</span></span>| <span data-ttu-id="bdec1-157">Значение</span><span class="sxs-lookup"><span data-stu-id="bdec1-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="bdec1-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="bdec1-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="bdec1-159">1.0</span><span class="sxs-lookup"><span data-stu-id="bdec1-159">1.0</span></span>|
|[<span data-ttu-id="bdec1-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="bdec1-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="bdec1-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="bdec1-161">ReadItem</span></span>|
|[<span data-ttu-id="bdec1-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="bdec1-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="bdec1-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="bdec1-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="bdec1-164">Пример</span><span class="sxs-lookup"><span data-stu-id="bdec1-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
