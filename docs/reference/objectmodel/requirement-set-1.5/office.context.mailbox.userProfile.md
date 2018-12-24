---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.5
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 748daf4d14aae1d14560d29e1d76eeea09830573
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432720"
---
# <a name="userprofile"></a><span data-ttu-id="25011-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="25011-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="25011-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="25011-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="25011-104">Требования</span><span class="sxs-lookup"><span data-stu-id="25011-104">Requirements</span></span>

|<span data-ttu-id="25011-105">Требование</span><span class="sxs-lookup"><span data-stu-id="25011-105">Requirement</span></span>| <span data-ttu-id="25011-106">Значение</span><span class="sxs-lookup"><span data-stu-id="25011-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="25011-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="25011-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="25011-108">1.0</span><span class="sxs-lookup"><span data-stu-id="25011-108">1.0</span></span>|
|[<span data-ttu-id="25011-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="25011-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="25011-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="25011-110">ReadItem</span></span>|
|[<span data-ttu-id="25011-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="25011-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="25011-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="25011-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="25011-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="25011-113">Members and methods</span></span>

| <span data-ttu-id="25011-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="25011-114">Member</span></span> | <span data-ttu-id="25011-115">Тип</span><span class="sxs-lookup"><span data-stu-id="25011-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="25011-116">displayName</span><span class="sxs-lookup"><span data-stu-id="25011-116">displayName</span></span>](#displayname-string) | <span data-ttu-id="25011-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="25011-117">Member</span></span> |
| [<span data-ttu-id="25011-118">emailAddress</span><span class="sxs-lookup"><span data-stu-id="25011-118">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="25011-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="25011-119">Member</span></span> |
| [<span data-ttu-id="25011-120">timeZone</span><span class="sxs-lookup"><span data-stu-id="25011-120">timeZone</span></span>](#timezone-string) | <span data-ttu-id="25011-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="25011-121">Member</span></span> |

### <a name="members"></a><span data-ttu-id="25011-122">Элементы</span><span class="sxs-lookup"><span data-stu-id="25011-122">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="25011-123">displayName :String</span><span class="sxs-lookup"><span data-stu-id="25011-123">displayName :String</span></span>

<span data-ttu-id="25011-124">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="25011-124">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="25011-125">Тип:</span><span class="sxs-lookup"><span data-stu-id="25011-125">Type:</span></span>

*   <span data-ttu-id="25011-126">String</span><span class="sxs-lookup"><span data-stu-id="25011-126">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="25011-127">Требования</span><span class="sxs-lookup"><span data-stu-id="25011-127">Requirements</span></span>

|<span data-ttu-id="25011-128">Требование</span><span class="sxs-lookup"><span data-stu-id="25011-128">Requirement</span></span>| <span data-ttu-id="25011-129">Значение</span><span class="sxs-lookup"><span data-stu-id="25011-129">Value</span></span>|
|---|---|
|[<span data-ttu-id="25011-130">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="25011-130">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="25011-131">1.0</span><span class="sxs-lookup"><span data-stu-id="25011-131">1.0</span></span>|
|[<span data-ttu-id="25011-132">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="25011-132">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="25011-133">ReadItem</span><span class="sxs-lookup"><span data-stu-id="25011-133">ReadItem</span></span>|
|[<span data-ttu-id="25011-134">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="25011-134">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="25011-135">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="25011-135">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="25011-136">Пример</span><span class="sxs-lookup"><span data-stu-id="25011-136">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="25011-137">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="25011-137">emailAddress :String</span></span>

<span data-ttu-id="25011-138">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="25011-138">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="25011-139">Тип:</span><span class="sxs-lookup"><span data-stu-id="25011-139">Type:</span></span>

*   <span data-ttu-id="25011-140">String</span><span class="sxs-lookup"><span data-stu-id="25011-140">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="25011-141">Требования</span><span class="sxs-lookup"><span data-stu-id="25011-141">Requirements</span></span>

|<span data-ttu-id="25011-142">Требование</span><span class="sxs-lookup"><span data-stu-id="25011-142">Requirement</span></span>| <span data-ttu-id="25011-143">Значение</span><span class="sxs-lookup"><span data-stu-id="25011-143">Value</span></span>|
|---|---|
|[<span data-ttu-id="25011-144">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="25011-144">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="25011-145">1.0</span><span class="sxs-lookup"><span data-stu-id="25011-145">1.0</span></span>|
|[<span data-ttu-id="25011-146">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="25011-146">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="25011-147">ReadItem</span><span class="sxs-lookup"><span data-stu-id="25011-147">ReadItem</span></span>|
|[<span data-ttu-id="25011-148">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="25011-148">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="25011-149">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="25011-149">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="25011-150">Пример</span><span class="sxs-lookup"><span data-stu-id="25011-150">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="25011-151">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="25011-151">timeZone :String</span></span>

<span data-ttu-id="25011-152">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="25011-152">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="25011-153">Тип:</span><span class="sxs-lookup"><span data-stu-id="25011-153">Type:</span></span>

*   <span data-ttu-id="25011-154">String</span><span class="sxs-lookup"><span data-stu-id="25011-154">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="25011-155">Требования</span><span class="sxs-lookup"><span data-stu-id="25011-155">Requirements</span></span>

|<span data-ttu-id="25011-156">Требование</span><span class="sxs-lookup"><span data-stu-id="25011-156">Requirement</span></span>| <span data-ttu-id="25011-157">Значение</span><span class="sxs-lookup"><span data-stu-id="25011-157">Value</span></span>|
|---|---|
|[<span data-ttu-id="25011-158">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="25011-158">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="25011-159">1.0</span><span class="sxs-lookup"><span data-stu-id="25011-159">1.0</span></span>|
|[<span data-ttu-id="25011-160">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="25011-160">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="25011-161">ReadItem</span><span class="sxs-lookup"><span data-stu-id="25011-161">ReadItem</span></span>|
|[<span data-ttu-id="25011-162">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="25011-162">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="25011-163">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="25011-163">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="25011-164">Пример</span><span class="sxs-lookup"><span data-stu-id="25011-164">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```