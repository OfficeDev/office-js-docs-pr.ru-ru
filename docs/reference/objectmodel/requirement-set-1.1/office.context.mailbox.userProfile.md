---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,1
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: af9a7f790f56124a86af08567690452b7f497408
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268490"
---
# <a name="userprofile"></a><span data-ttu-id="da947-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="da947-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="da947-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="da947-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="da947-104">Требования</span><span class="sxs-lookup"><span data-stu-id="da947-104">Requirements</span></span>

|<span data-ttu-id="da947-105">Требование</span><span class="sxs-lookup"><span data-stu-id="da947-105">Requirement</span></span>| <span data-ttu-id="da947-106">Значение</span><span class="sxs-lookup"><span data-stu-id="da947-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="da947-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="da947-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da947-108">1.0</span><span class="sxs-lookup"><span data-stu-id="da947-108">1.0</span></span>|
|[<span data-ttu-id="da947-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="da947-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da947-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da947-110">ReadItem</span></span>|
|[<span data-ttu-id="da947-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="da947-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da947-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="da947-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="da947-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="da947-113">Members and methods</span></span>

| <span data-ttu-id="da947-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="da947-114">Member</span></span> | <span data-ttu-id="da947-115">Тип</span><span class="sxs-lookup"><span data-stu-id="da947-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="da947-116">displayName</span><span class="sxs-lookup"><span data-stu-id="da947-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="da947-117">Member</span><span class="sxs-lookup"><span data-stu-id="da947-117">Member</span></span> |
| [<span data-ttu-id="da947-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="da947-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="da947-119">Member</span><span class="sxs-lookup"><span data-stu-id="da947-119">Member</span></span> |
| [<span data-ttu-id="da947-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="da947-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="da947-121">Member</span><span class="sxs-lookup"><span data-stu-id="da947-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="da947-122">Members</span><span class="sxs-lookup"><span data-stu-id="da947-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="da947-123">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="da947-123">displayName: String</span></span>

<span data-ttu-id="da947-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="da947-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="da947-125">Тип</span><span class="sxs-lookup"><span data-stu-id="da947-125">Type</span></span>

*   <span data-ttu-id="da947-126">String</span><span class="sxs-lookup"><span data-stu-id="da947-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da947-127">Требования</span><span class="sxs-lookup"><span data-stu-id="da947-127">Requirements</span></span>

|<span data-ttu-id="da947-128">Требование</span><span class="sxs-lookup"><span data-stu-id="da947-128">Requirement</span></span>| <span data-ttu-id="da947-129">Значение</span><span class="sxs-lookup"><span data-stu-id="da947-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="da947-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="da947-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da947-131">1.0</span><span class="sxs-lookup"><span data-stu-id="da947-131">1.0</span></span>|
|[<span data-ttu-id="da947-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="da947-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da947-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da947-133">ReadItem</span></span>|
|[<span data-ttu-id="da947-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="da947-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da947-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="da947-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da947-136">Пример</span><span class="sxs-lookup"><span data-stu-id="da947-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="da947-137">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="da947-137">emailAddress: String</span></span>

<span data-ttu-id="da947-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="da947-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="da947-139">Тип</span><span class="sxs-lookup"><span data-stu-id="da947-139">Type</span></span>

*   <span data-ttu-id="da947-140">String</span><span class="sxs-lookup"><span data-stu-id="da947-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da947-141">Требования</span><span class="sxs-lookup"><span data-stu-id="da947-141">Requirements</span></span>

|<span data-ttu-id="da947-142">Требование</span><span class="sxs-lookup"><span data-stu-id="da947-142">Requirement</span></span>| <span data-ttu-id="da947-143">Значение</span><span class="sxs-lookup"><span data-stu-id="da947-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="da947-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="da947-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da947-145">1.0</span><span class="sxs-lookup"><span data-stu-id="da947-145">1.0</span></span>|
|[<span data-ttu-id="da947-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="da947-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da947-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da947-147">ReadItem</span></span>|
|[<span data-ttu-id="da947-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="da947-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da947-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="da947-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da947-150">Пример</span><span class="sxs-lookup"><span data-stu-id="da947-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="da947-151">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="da947-151">timeZone: String</span></span>

<span data-ttu-id="da947-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="da947-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="da947-153">Тип</span><span class="sxs-lookup"><span data-stu-id="da947-153">Type</span></span>

*   <span data-ttu-id="da947-154">String</span><span class="sxs-lookup"><span data-stu-id="da947-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="da947-155">Требования</span><span class="sxs-lookup"><span data-stu-id="da947-155">Requirements</span></span>

|<span data-ttu-id="da947-156">Требование</span><span class="sxs-lookup"><span data-stu-id="da947-156">Requirement</span></span>| <span data-ttu-id="da947-157">Значение</span><span class="sxs-lookup"><span data-stu-id="da947-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="da947-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="da947-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="da947-159">1.0</span><span class="sxs-lookup"><span data-stu-id="da947-159">1.0</span></span>|
|[<span data-ttu-id="da947-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="da947-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="da947-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="da947-161">ReadItem</span></span>|
|[<span data-ttu-id="da947-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="da947-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="da947-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="da947-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="da947-164">Пример</span><span class="sxs-lookup"><span data-stu-id="da947-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
