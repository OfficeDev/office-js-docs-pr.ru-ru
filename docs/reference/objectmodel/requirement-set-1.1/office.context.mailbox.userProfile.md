---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.1
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 312cba4d5aace980b7c9b205899fac51d3da3de5
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27433175"
---
# <a name="userprofile"></a><span data-ttu-id="4a6d9-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="4a6d9-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="4a6d9-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="4a6d9-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a6d9-104">Требования</span><span class="sxs-lookup"><span data-stu-id="4a6d9-104">Requirements</span></span>

|<span data-ttu-id="4a6d9-105">Требование</span><span class="sxs-lookup"><span data-stu-id="4a6d9-105">Requirement</span></span>| <span data-ttu-id="4a6d9-106">Значение</span><span class="sxs-lookup"><span data-stu-id="4a6d9-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a6d9-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4a6d9-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a6d9-108">1.0</span><span class="sxs-lookup"><span data-stu-id="4a6d9-108">1.0</span></span>|
|[<span data-ttu-id="4a6d9-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4a6d9-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a6d9-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a6d9-110">ReadItem</span></span>|
|[<span data-ttu-id="4a6d9-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4a6d9-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a6d9-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4a6d9-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="4a6d9-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="4a6d9-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="4a6d9-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="4a6d9-114">displayName :String</span></span>

<span data-ttu-id="4a6d9-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="4a6d9-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="4a6d9-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="4a6d9-116">Type:</span></span>

*   <span data-ttu-id="4a6d9-117">String</span><span class="sxs-lookup"><span data-stu-id="4a6d9-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a6d9-118">Требования</span><span class="sxs-lookup"><span data-stu-id="4a6d9-118">Requirements</span></span>

|<span data-ttu-id="4a6d9-119">Требование</span><span class="sxs-lookup"><span data-stu-id="4a6d9-119">Requirement</span></span>| <span data-ttu-id="4a6d9-120">Значение</span><span class="sxs-lookup"><span data-stu-id="4a6d9-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a6d9-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4a6d9-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a6d9-122">1.0</span><span class="sxs-lookup"><span data-stu-id="4a6d9-122">1.0</span></span>|
|[<span data-ttu-id="4a6d9-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4a6d9-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a6d9-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a6d9-124">ReadItem</span></span>|
|[<span data-ttu-id="4a6d9-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4a6d9-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a6d9-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4a6d9-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a6d9-127">Пример</span><span class="sxs-lookup"><span data-stu-id="4a6d9-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="4a6d9-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="4a6d9-128">emailAddress :String</span></span>

<span data-ttu-id="4a6d9-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="4a6d9-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="4a6d9-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="4a6d9-130">Type:</span></span>

*   <span data-ttu-id="4a6d9-131">String</span><span class="sxs-lookup"><span data-stu-id="4a6d9-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a6d9-132">Требования</span><span class="sxs-lookup"><span data-stu-id="4a6d9-132">Requirements</span></span>

|<span data-ttu-id="4a6d9-133">Требование</span><span class="sxs-lookup"><span data-stu-id="4a6d9-133">Requirement</span></span>| <span data-ttu-id="4a6d9-134">Значение</span><span class="sxs-lookup"><span data-stu-id="4a6d9-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a6d9-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4a6d9-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a6d9-136">1.0</span><span class="sxs-lookup"><span data-stu-id="4a6d9-136">1.0</span></span>|
|[<span data-ttu-id="4a6d9-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4a6d9-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a6d9-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a6d9-138">ReadItem</span></span>|
|[<span data-ttu-id="4a6d9-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4a6d9-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a6d9-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4a6d9-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a6d9-141">Пример</span><span class="sxs-lookup"><span data-stu-id="4a6d9-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="4a6d9-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="4a6d9-142">timeZone :String</span></span>

<span data-ttu-id="4a6d9-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="4a6d9-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="4a6d9-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="4a6d9-144">Type:</span></span>

*   <span data-ttu-id="4a6d9-145">String</span><span class="sxs-lookup"><span data-stu-id="4a6d9-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="4a6d9-146">Требования</span><span class="sxs-lookup"><span data-stu-id="4a6d9-146">Requirements</span></span>

|<span data-ttu-id="4a6d9-147">Требование</span><span class="sxs-lookup"><span data-stu-id="4a6d9-147">Requirement</span></span>| <span data-ttu-id="4a6d9-148">Значение</span><span class="sxs-lookup"><span data-stu-id="4a6d9-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="4a6d9-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="4a6d9-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="4a6d9-150">1.0</span><span class="sxs-lookup"><span data-stu-id="4a6d9-150">1.0</span></span>|
|[<span data-ttu-id="4a6d9-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="4a6d9-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="4a6d9-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="4a6d9-152">ReadItem</span></span>|
|[<span data-ttu-id="4a6d9-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="4a6d9-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="4a6d9-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="4a6d9-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="4a6d9-155">Пример</span><span class="sxs-lookup"><span data-stu-id="4a6d9-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```