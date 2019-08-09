---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,3
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 8924d8b0dfa5bb43be8867cbd0e83ee01ff788cb
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268392"
---
# <a name="userprofile"></a><span data-ttu-id="91c16-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="91c16-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="91c16-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="91c16-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="91c16-104">Требования</span><span class="sxs-lookup"><span data-stu-id="91c16-104">Requirements</span></span>

|<span data-ttu-id="91c16-105">Требование</span><span class="sxs-lookup"><span data-stu-id="91c16-105">Requirement</span></span>| <span data-ttu-id="91c16-106">Значение</span><span class="sxs-lookup"><span data-stu-id="91c16-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="91c16-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="91c16-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="91c16-108">1.0</span><span class="sxs-lookup"><span data-stu-id="91c16-108">1.0</span></span>|
|[<span data-ttu-id="91c16-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91c16-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="91c16-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91c16-110">ReadItem</span></span>|
|[<span data-ttu-id="91c16-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91c16-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="91c16-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91c16-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="91c16-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="91c16-113">Members and methods</span></span>

| <span data-ttu-id="91c16-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="91c16-114">Member</span></span> | <span data-ttu-id="91c16-115">Тип</span><span class="sxs-lookup"><span data-stu-id="91c16-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="91c16-116">displayName</span><span class="sxs-lookup"><span data-stu-id="91c16-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="91c16-117">Member</span><span class="sxs-lookup"><span data-stu-id="91c16-117">Member</span></span> |
| [<span data-ttu-id="91c16-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="91c16-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="91c16-119">Member</span><span class="sxs-lookup"><span data-stu-id="91c16-119">Member</span></span> |
| [<span data-ttu-id="91c16-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="91c16-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="91c16-121">Member</span><span class="sxs-lookup"><span data-stu-id="91c16-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="91c16-122">Members</span><span class="sxs-lookup"><span data-stu-id="91c16-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="91c16-123">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="91c16-123">displayName: String</span></span>

<span data-ttu-id="91c16-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="91c16-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="91c16-125">Тип</span><span class="sxs-lookup"><span data-stu-id="91c16-125">Type</span></span>

*   <span data-ttu-id="91c16-126">String</span><span class="sxs-lookup"><span data-stu-id="91c16-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91c16-127">Требования</span><span class="sxs-lookup"><span data-stu-id="91c16-127">Requirements</span></span>

|<span data-ttu-id="91c16-128">Требование</span><span class="sxs-lookup"><span data-stu-id="91c16-128">Requirement</span></span>| <span data-ttu-id="91c16-129">Значение</span><span class="sxs-lookup"><span data-stu-id="91c16-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="91c16-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="91c16-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="91c16-131">1.0</span><span class="sxs-lookup"><span data-stu-id="91c16-131">1.0</span></span>|
|[<span data-ttu-id="91c16-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91c16-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="91c16-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91c16-133">ReadItem</span></span>|
|[<span data-ttu-id="91c16-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91c16-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="91c16-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91c16-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91c16-136">Пример</span><span class="sxs-lookup"><span data-stu-id="91c16-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="91c16-137">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="91c16-137">emailAddress: String</span></span>

<span data-ttu-id="91c16-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="91c16-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="91c16-139">Тип</span><span class="sxs-lookup"><span data-stu-id="91c16-139">Type</span></span>

*   <span data-ttu-id="91c16-140">String</span><span class="sxs-lookup"><span data-stu-id="91c16-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91c16-141">Требования</span><span class="sxs-lookup"><span data-stu-id="91c16-141">Requirements</span></span>

|<span data-ttu-id="91c16-142">Требование</span><span class="sxs-lookup"><span data-stu-id="91c16-142">Requirement</span></span>| <span data-ttu-id="91c16-143">Значение</span><span class="sxs-lookup"><span data-stu-id="91c16-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="91c16-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="91c16-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="91c16-145">1.0</span><span class="sxs-lookup"><span data-stu-id="91c16-145">1.0</span></span>|
|[<span data-ttu-id="91c16-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91c16-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="91c16-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91c16-147">ReadItem</span></span>|
|[<span data-ttu-id="91c16-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91c16-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="91c16-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91c16-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91c16-150">Пример</span><span class="sxs-lookup"><span data-stu-id="91c16-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="91c16-151">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="91c16-151">timeZone: String</span></span>

<span data-ttu-id="91c16-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="91c16-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="91c16-153">Тип</span><span class="sxs-lookup"><span data-stu-id="91c16-153">Type</span></span>

*   <span data-ttu-id="91c16-154">String</span><span class="sxs-lookup"><span data-stu-id="91c16-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="91c16-155">Требования</span><span class="sxs-lookup"><span data-stu-id="91c16-155">Requirements</span></span>

|<span data-ttu-id="91c16-156">Требование</span><span class="sxs-lookup"><span data-stu-id="91c16-156">Requirement</span></span>| <span data-ttu-id="91c16-157">Значение</span><span class="sxs-lookup"><span data-stu-id="91c16-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="91c16-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="91c16-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="91c16-159">1.0</span><span class="sxs-lookup"><span data-stu-id="91c16-159">1.0</span></span>|
|[<span data-ttu-id="91c16-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="91c16-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="91c16-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="91c16-161">ReadItem</span></span>|
|[<span data-ttu-id="91c16-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="91c16-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="91c16-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="91c16-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="91c16-164">Пример</span><span class="sxs-lookup"><span data-stu-id="91c16-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
