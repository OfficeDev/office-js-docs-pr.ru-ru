---
title: Office. Context. Mailbox. userProfile — набор обязательных элементов 1,2
description: ''
ms.date: 08/08/2019
localization_priority: Normal
ms.openlocfilehash: 7258195e7ec0ef2432723d0f32f3d9ef1a3acf2b
ms.sourcegitcommit: 654ac1a0c477413662b48cffc0faee5cb65fc25f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/09/2019
ms.locfileid: "36268686"
---
# <a name="userprofile"></a><span data-ttu-id="26e5f-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="26e5f-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="26e5f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="26e5f-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="26e5f-104">Требования</span><span class="sxs-lookup"><span data-stu-id="26e5f-104">Requirements</span></span>

|<span data-ttu-id="26e5f-105">Требование</span><span class="sxs-lookup"><span data-stu-id="26e5f-105">Requirement</span></span>| <span data-ttu-id="26e5f-106">Значение</span><span class="sxs-lookup"><span data-stu-id="26e5f-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="26e5f-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26e5f-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26e5f-108">1.0</span><span class="sxs-lookup"><span data-stu-id="26e5f-108">1.0</span></span>|
|[<span data-ttu-id="26e5f-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26e5f-109">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26e5f-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26e5f-110">ReadItem</span></span>|
|[<span data-ttu-id="26e5f-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26e5f-111">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26e5f-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26e5f-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="26e5f-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="26e5f-113">Members and methods</span></span>

| <span data-ttu-id="26e5f-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="26e5f-114">Member</span></span> | <span data-ttu-id="26e5f-115">Тип</span><span class="sxs-lookup"><span data-stu-id="26e5f-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="26e5f-116">displayName</span><span class="sxs-lookup"><span data-stu-id="26e5f-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="26e5f-117">Member</span><span class="sxs-lookup"><span data-stu-id="26e5f-117">Member</span></span> |
| [<span data-ttu-id="26e5f-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="26e5f-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="26e5f-119">Member</span><span class="sxs-lookup"><span data-stu-id="26e5f-119">Member</span></span> |
| [<span data-ttu-id="26e5f-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="26e5f-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="26e5f-121">Member</span><span class="sxs-lookup"><span data-stu-id="26e5f-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="26e5f-122">Members</span><span class="sxs-lookup"><span data-stu-id="26e5f-122">Members</span></span>

#### <a name="displayname-string"></a><span data-ttu-id="26e5f-123">displayName: строка</span><span class="sxs-lookup"><span data-stu-id="26e5f-123">displayName: String</span></span>

<span data-ttu-id="26e5f-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="26e5f-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="26e5f-125">Тип</span><span class="sxs-lookup"><span data-stu-id="26e5f-125">Type</span></span>

*   <span data-ttu-id="26e5f-126">String</span><span class="sxs-lookup"><span data-stu-id="26e5f-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="26e5f-127">Требования</span><span class="sxs-lookup"><span data-stu-id="26e5f-127">Requirements</span></span>

|<span data-ttu-id="26e5f-128">Требование</span><span class="sxs-lookup"><span data-stu-id="26e5f-128">Requirement</span></span>| <span data-ttu-id="26e5f-129">Значение</span><span class="sxs-lookup"><span data-stu-id="26e5f-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="26e5f-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26e5f-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26e5f-131">1.0</span><span class="sxs-lookup"><span data-stu-id="26e5f-131">1.0</span></span>|
|[<span data-ttu-id="26e5f-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26e5f-132">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26e5f-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26e5f-133">ReadItem</span></span>|
|[<span data-ttu-id="26e5f-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26e5f-134">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26e5f-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26e5f-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="26e5f-136">Пример</span><span class="sxs-lookup"><span data-stu-id="26e5f-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

#### <a name="emailaddress-string"></a><span data-ttu-id="26e5f-137">emailAddress: строка</span><span class="sxs-lookup"><span data-stu-id="26e5f-137">emailAddress: String</span></span>

<span data-ttu-id="26e5f-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="26e5f-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="26e5f-139">Тип</span><span class="sxs-lookup"><span data-stu-id="26e5f-139">Type</span></span>

*   <span data-ttu-id="26e5f-140">String</span><span class="sxs-lookup"><span data-stu-id="26e5f-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="26e5f-141">Требования</span><span class="sxs-lookup"><span data-stu-id="26e5f-141">Requirements</span></span>

|<span data-ttu-id="26e5f-142">Требование</span><span class="sxs-lookup"><span data-stu-id="26e5f-142">Requirement</span></span>| <span data-ttu-id="26e5f-143">Значение</span><span class="sxs-lookup"><span data-stu-id="26e5f-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="26e5f-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26e5f-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26e5f-145">1.0</span><span class="sxs-lookup"><span data-stu-id="26e5f-145">1.0</span></span>|
|[<span data-ttu-id="26e5f-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26e5f-146">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26e5f-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26e5f-147">ReadItem</span></span>|
|[<span data-ttu-id="26e5f-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26e5f-148">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26e5f-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26e5f-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="26e5f-150">Пример</span><span class="sxs-lookup"><span data-stu-id="26e5f-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

#### <a name="timezone-string"></a><span data-ttu-id="26e5f-151">Часовой пояс: строка</span><span class="sxs-lookup"><span data-stu-id="26e5f-151">timeZone: String</span></span>

<span data-ttu-id="26e5f-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="26e5f-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="26e5f-153">Тип</span><span class="sxs-lookup"><span data-stu-id="26e5f-153">Type</span></span>

*   <span data-ttu-id="26e5f-154">String</span><span class="sxs-lookup"><span data-stu-id="26e5f-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="26e5f-155">Требования</span><span class="sxs-lookup"><span data-stu-id="26e5f-155">Requirements</span></span>

|<span data-ttu-id="26e5f-156">Требование</span><span class="sxs-lookup"><span data-stu-id="26e5f-156">Requirement</span></span>| <span data-ttu-id="26e5f-157">Значение</span><span class="sxs-lookup"><span data-stu-id="26e5f-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="26e5f-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="26e5f-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="26e5f-159">1.0</span><span class="sxs-lookup"><span data-stu-id="26e5f-159">1.0</span></span>|
|[<span data-ttu-id="26e5f-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="26e5f-160">Minimum permission level</span></span>](/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="26e5f-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="26e5f-161">ReadItem</span></span>|
|[<span data-ttu-id="26e5f-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="26e5f-162">Applicable Outlook mode</span></span>](/outlook/add-ins/#extension-points)| <span data-ttu-id="26e5f-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="26e5f-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="26e5f-164">Пример</span><span class="sxs-lookup"><span data-stu-id="26e5f-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
