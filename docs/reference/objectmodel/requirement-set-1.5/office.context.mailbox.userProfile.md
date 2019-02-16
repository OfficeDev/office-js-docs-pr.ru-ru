---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.5
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: e98e88cde184db121e69fdd267dff4e39d887b1f
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067833"
---
# <a name="userprofile"></a><span data-ttu-id="12f26-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="12f26-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="12f26-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="12f26-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="12f26-104">Требования</span><span class="sxs-lookup"><span data-stu-id="12f26-104">Requirements</span></span>

|<span data-ttu-id="12f26-105">Требование</span><span class="sxs-lookup"><span data-stu-id="12f26-105">Requirement</span></span>| <span data-ttu-id="12f26-106">Значение</span><span class="sxs-lookup"><span data-stu-id="12f26-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f26-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="12f26-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12f26-108">1.0</span><span class="sxs-lookup"><span data-stu-id="12f26-108">1.0</span></span>|
|[<span data-ttu-id="12f26-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="12f26-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12f26-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12f26-110">ReadItem</span></span>|
|[<span data-ttu-id="12f26-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="12f26-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="12f26-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="12f26-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="12f26-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="12f26-113">Members and methods</span></span>

| <span data-ttu-id="12f26-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="12f26-114">Member</span></span> | <span data-ttu-id="12f26-115">Тип</span><span class="sxs-lookup"><span data-stu-id="12f26-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="12f26-116">displayName</span><span class="sxs-lookup"><span data-stu-id="12f26-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="12f26-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="12f26-117">Member</span></span> |
| [<span data-ttu-id="12f26-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="12f26-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="12f26-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="12f26-119">Member</span></span> |
| [<span data-ttu-id="12f26-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="12f26-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="12f26-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="12f26-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="12f26-122">Элементы</span><span class="sxs-lookup"><span data-stu-id="12f26-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="12f26-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="12f26-123">displayName :String</span></span>

<span data-ttu-id="12f26-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="12f26-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="12f26-125">Тип</span><span class="sxs-lookup"><span data-stu-id="12f26-125">Type</span></span>

*   <span data-ttu-id="12f26-126">String</span><span class="sxs-lookup"><span data-stu-id="12f26-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="12f26-127">Требования</span><span class="sxs-lookup"><span data-stu-id="12f26-127">Requirements</span></span>

|<span data-ttu-id="12f26-128">Требование</span><span class="sxs-lookup"><span data-stu-id="12f26-128">Requirement</span></span>| <span data-ttu-id="12f26-129">Значение</span><span class="sxs-lookup"><span data-stu-id="12f26-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f26-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="12f26-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12f26-131">1.0</span><span class="sxs-lookup"><span data-stu-id="12f26-131">1.0</span></span>|
|[<span data-ttu-id="12f26-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="12f26-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12f26-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12f26-133">ReadItem</span></span>|
|[<span data-ttu-id="12f26-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="12f26-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="12f26-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="12f26-135">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f26-136">Пример</span><span class="sxs-lookup"><span data-stu-id="12f26-136">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="12f26-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="12f26-137">emailAddress :String</span></span>

<span data-ttu-id="12f26-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="12f26-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="12f26-139">Тип</span><span class="sxs-lookup"><span data-stu-id="12f26-139">Type</span></span>

*   <span data-ttu-id="12f26-140">String</span><span class="sxs-lookup"><span data-stu-id="12f26-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="12f26-141">Требования</span><span class="sxs-lookup"><span data-stu-id="12f26-141">Requirements</span></span>

|<span data-ttu-id="12f26-142">Требование</span><span class="sxs-lookup"><span data-stu-id="12f26-142">Requirement</span></span>| <span data-ttu-id="12f26-143">Значение</span><span class="sxs-lookup"><span data-stu-id="12f26-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f26-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="12f26-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12f26-145">1.0</span><span class="sxs-lookup"><span data-stu-id="12f26-145">1.0</span></span>|
|[<span data-ttu-id="12f26-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="12f26-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12f26-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12f26-147">ReadItem</span></span>|
|[<span data-ttu-id="12f26-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="12f26-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="12f26-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="12f26-149">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f26-150">Пример</span><span class="sxs-lookup"><span data-stu-id="12f26-150">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="12f26-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="12f26-151">timeZone :String</span></span>

<span data-ttu-id="12f26-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="12f26-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="12f26-153">Тип</span><span class="sxs-lookup"><span data-stu-id="12f26-153">Type</span></span>

*   <span data-ttu-id="12f26-154">String</span><span class="sxs-lookup"><span data-stu-id="12f26-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="12f26-155">Требования</span><span class="sxs-lookup"><span data-stu-id="12f26-155">Requirements</span></span>

|<span data-ttu-id="12f26-156">Требование</span><span class="sxs-lookup"><span data-stu-id="12f26-156">Requirement</span></span>| <span data-ttu-id="12f26-157">Значение</span><span class="sxs-lookup"><span data-stu-id="12f26-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="12f26-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="12f26-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="12f26-159">1.0</span><span class="sxs-lookup"><span data-stu-id="12f26-159">1.0</span></span>|
|[<span data-ttu-id="12f26-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="12f26-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="12f26-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="12f26-161">ReadItem</span></span>|
|[<span data-ttu-id="12f26-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="12f26-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="12f26-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="12f26-163">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="12f26-164">Пример</span><span class="sxs-lookup"><span data-stu-id="12f26-164">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
