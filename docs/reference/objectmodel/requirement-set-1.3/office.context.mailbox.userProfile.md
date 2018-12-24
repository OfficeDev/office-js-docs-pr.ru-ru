---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.3
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 9f36b5f1d31ad6709cf2c43ce7dcb3f91a35bd00
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432223"
---
# <a name="userprofile"></a><span data-ttu-id="20ddd-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="20ddd-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="20ddd-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="20ddd-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="20ddd-104">Требования</span><span class="sxs-lookup"><span data-stu-id="20ddd-104">Requirements</span></span>

|<span data-ttu-id="20ddd-105">Требование</span><span class="sxs-lookup"><span data-stu-id="20ddd-105">Requirement</span></span>| <span data-ttu-id="20ddd-106">Значение</span><span class="sxs-lookup"><span data-stu-id="20ddd-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="20ddd-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="20ddd-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20ddd-108">1.0</span><span class="sxs-lookup"><span data-stu-id="20ddd-108">1.0</span></span>|
|[<span data-ttu-id="20ddd-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="20ddd-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20ddd-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20ddd-110">ReadItem</span></span>|
|[<span data-ttu-id="20ddd-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="20ddd-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20ddd-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="20ddd-112">Compose or read</span></span>|

### <a name="members"></a><span data-ttu-id="20ddd-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="20ddd-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="20ddd-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="20ddd-114">displayName :String</span></span>

<span data-ttu-id="20ddd-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="20ddd-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="20ddd-116">Тип:</span><span class="sxs-lookup"><span data-stu-id="20ddd-116">Type:</span></span>

*   <span data-ttu-id="20ddd-117">String</span><span class="sxs-lookup"><span data-stu-id="20ddd-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="20ddd-118">Требования</span><span class="sxs-lookup"><span data-stu-id="20ddd-118">Requirements</span></span>

|<span data-ttu-id="20ddd-119">Требование</span><span class="sxs-lookup"><span data-stu-id="20ddd-119">Requirement</span></span>| <span data-ttu-id="20ddd-120">Значение</span><span class="sxs-lookup"><span data-stu-id="20ddd-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="20ddd-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="20ddd-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20ddd-122">1.0</span><span class="sxs-lookup"><span data-stu-id="20ddd-122">1.0</span></span>|
|[<span data-ttu-id="20ddd-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="20ddd-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20ddd-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20ddd-124">ReadItem</span></span>|
|[<span data-ttu-id="20ddd-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="20ddd-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20ddd-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="20ddd-126">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="20ddd-127">Пример</span><span class="sxs-lookup"><span data-stu-id="20ddd-127">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="20ddd-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="20ddd-128">emailAddress :String</span></span>

<span data-ttu-id="20ddd-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="20ddd-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="20ddd-130">Тип:</span><span class="sxs-lookup"><span data-stu-id="20ddd-130">Type:</span></span>

*   <span data-ttu-id="20ddd-131">String</span><span class="sxs-lookup"><span data-stu-id="20ddd-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="20ddd-132">Требования</span><span class="sxs-lookup"><span data-stu-id="20ddd-132">Requirements</span></span>

|<span data-ttu-id="20ddd-133">Требование</span><span class="sxs-lookup"><span data-stu-id="20ddd-133">Requirement</span></span>| <span data-ttu-id="20ddd-134">Значение</span><span class="sxs-lookup"><span data-stu-id="20ddd-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="20ddd-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="20ddd-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20ddd-136">1.0</span><span class="sxs-lookup"><span data-stu-id="20ddd-136">1.0</span></span>|
|[<span data-ttu-id="20ddd-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="20ddd-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20ddd-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20ddd-138">ReadItem</span></span>|
|[<span data-ttu-id="20ddd-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="20ddd-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20ddd-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="20ddd-140">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="20ddd-141">Пример</span><span class="sxs-lookup"><span data-stu-id="20ddd-141">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="20ddd-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="20ddd-142">timeZone :String</span></span>

<span data-ttu-id="20ddd-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="20ddd-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="20ddd-144">Тип:</span><span class="sxs-lookup"><span data-stu-id="20ddd-144">Type:</span></span>

*   <span data-ttu-id="20ddd-145">String</span><span class="sxs-lookup"><span data-stu-id="20ddd-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="20ddd-146">Требования</span><span class="sxs-lookup"><span data-stu-id="20ddd-146">Requirements</span></span>

|<span data-ttu-id="20ddd-147">Требование</span><span class="sxs-lookup"><span data-stu-id="20ddd-147">Requirement</span></span>| <span data-ttu-id="20ddd-148">Значение</span><span class="sxs-lookup"><span data-stu-id="20ddd-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="20ddd-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="20ddd-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="20ddd-150">1.0</span><span class="sxs-lookup"><span data-stu-id="20ddd-150">1.0</span></span>|
|[<span data-ttu-id="20ddd-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="20ddd-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="20ddd-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="20ddd-152">ReadItem</span></span>|
|[<span data-ttu-id="20ddd-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="20ddd-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="20ddd-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="20ddd-154">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="20ddd-155">Пример</span><span class="sxs-lookup"><span data-stu-id="20ddd-155">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```