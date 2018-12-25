---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.4
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 55d0a789c8e46fd3f6ee69f39cf33f7e7d94c322
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432650"
---
# <a name="userprofile"></a><span data-ttu-id="ee3a2-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ee3a2-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ee3a2-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ee3a2-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee3a2-104">Требования</span><span class="sxs-lookup"><span data-stu-id="ee3a2-104">Requirements</span></span>

|<span data-ttu-id="ee3a2-105">Требование</span><span class="sxs-lookup"><span data-stu-id="ee3a2-105">Requirement</span></span>| <span data-ttu-id="ee3a2-106">Значение</span><span class="sxs-lookup"><span data-stu-id="ee3a2-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee3a2-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ee3a2-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee3a2-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ee3a2-108">1.0</span></span>|
|[<span data-ttu-id="ee3a2-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ee3a2-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee3a2-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee3a2-110">ReadItem</span></span>|
|[<span data-ttu-id="ee3a2-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ee3a2-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee3a2-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ee3a2-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="ee3a2-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="ee3a2-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="ee3a2-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ee3a2-114">displayName :String</span></span>

<span data-ttu-id="ee3a2-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="ee3a2-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ee3a2-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="ee3a2-116">Type:</span></span>

*   <span data-ttu-id="ee3a2-117">String</span><span class="sxs-lookup"><span data-stu-id="ee3a2-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee3a2-118">Требования</span><span class="sxs-lookup"><span data-stu-id="ee3a2-118">Requirements</span></span>

|<span data-ttu-id="ee3a2-119">Требование</span><span class="sxs-lookup"><span data-stu-id="ee3a2-119">Requirement</span></span>| <span data-ttu-id="ee3a2-120">Значение</span><span class="sxs-lookup"><span data-stu-id="ee3a2-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee3a2-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ee3a2-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee3a2-122">1.0</span><span class="sxs-lookup"><span data-stu-id="ee3a2-122">1.0</span></span>|
|[<span data-ttu-id="ee3a2-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ee3a2-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee3a2-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee3a2-124">ReadItem</span></span>|
|[<span data-ttu-id="ee3a2-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ee3a2-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee3a2-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ee3a2-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee3a2-127">Пример</span><span class="sxs-lookup"><span data-stu-id="ee3a2-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ee3a2-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ee3a2-128">emailAddress :String</span></span>

<span data-ttu-id="ee3a2-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="ee3a2-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ee3a2-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="ee3a2-130">Type:</span></span>

*   <span data-ttu-id="ee3a2-131">String</span><span class="sxs-lookup"><span data-stu-id="ee3a2-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee3a2-132">Требования</span><span class="sxs-lookup"><span data-stu-id="ee3a2-132">Requirements</span></span>

|<span data-ttu-id="ee3a2-133">Требование</span><span class="sxs-lookup"><span data-stu-id="ee3a2-133">Requirement</span></span>| <span data-ttu-id="ee3a2-134">Значение</span><span class="sxs-lookup"><span data-stu-id="ee3a2-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee3a2-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ee3a2-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee3a2-136">1.0</span><span class="sxs-lookup"><span data-stu-id="ee3a2-136">1.0</span></span>|
|[<span data-ttu-id="ee3a2-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ee3a2-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee3a2-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee3a2-138">ReadItem</span></span>|
|[<span data-ttu-id="ee3a2-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ee3a2-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee3a2-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ee3a2-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee3a2-141">Пример</span><span class="sxs-lookup"><span data-stu-id="ee3a2-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ee3a2-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ee3a2-142">timeZone :String</span></span>

<span data-ttu-id="ee3a2-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ee3a2-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ee3a2-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="ee3a2-144">Type:</span></span>

*   <span data-ttu-id="ee3a2-145">String</span><span class="sxs-lookup"><span data-stu-id="ee3a2-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ee3a2-146">Требования</span><span class="sxs-lookup"><span data-stu-id="ee3a2-146">Requirements</span></span>

|<span data-ttu-id="ee3a2-147">Требование</span><span class="sxs-lookup"><span data-stu-id="ee3a2-147">Requirement</span></span>| <span data-ttu-id="ee3a2-148">Значение</span><span class="sxs-lookup"><span data-stu-id="ee3a2-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="ee3a2-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ee3a2-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ee3a2-150">1.0</span><span class="sxs-lookup"><span data-stu-id="ee3a2-150">1.0</span></span>|
|[<span data-ttu-id="ee3a2-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ee3a2-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ee3a2-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ee3a2-152">ReadItem</span></span>|
|[<span data-ttu-id="ee3a2-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ee3a2-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ee3a2-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ee3a2-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="ee3a2-155">Пример</span><span class="sxs-lookup"><span data-stu-id="ee3a2-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```