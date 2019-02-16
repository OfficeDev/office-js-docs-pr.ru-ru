---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.3
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: e29cf90d1c5d4c288417ef98f6e9d22eaf908b67
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067932"
---
# <a name="userprofile"></a><span data-ttu-id="628aa-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="628aa-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="628aa-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="628aa-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="628aa-104">Требования</span><span class="sxs-lookup"><span data-stu-id="628aa-104">Requirements</span></span>

|<span data-ttu-id="628aa-105">Требование</span><span class="sxs-lookup"><span data-stu-id="628aa-105">Requirement</span></span>| <span data-ttu-id="628aa-106">Значение</span><span class="sxs-lookup"><span data-stu-id="628aa-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="628aa-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="628aa-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="628aa-108">1.0</span><span class="sxs-lookup"><span data-stu-id="628aa-108">1.0</span></span>|
|[<span data-ttu-id="628aa-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="628aa-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="628aa-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="628aa-110">ReadItem</span></span>|
|[<span data-ttu-id="628aa-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="628aa-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="628aa-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="628aa-112">Compose or Read</span></span>|

### <a name="members"></a><span data-ttu-id="628aa-113">Элементы</span><span class="sxs-lookup"><span data-stu-id="628aa-113">Members</span></span>

####  <a name="displayname-string"></a><span data-ttu-id="628aa-114">displayName :String</span><span class="sxs-lookup"><span data-stu-id="628aa-114">displayName :String</span></span>

<span data-ttu-id="628aa-115">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="628aa-115">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="628aa-116">Тип</span><span class="sxs-lookup"><span data-stu-id="628aa-116">Type</span></span>

*   <span data-ttu-id="628aa-117">String</span><span class="sxs-lookup"><span data-stu-id="628aa-117">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="628aa-118">Требования</span><span class="sxs-lookup"><span data-stu-id="628aa-118">Requirements</span></span>

|<span data-ttu-id="628aa-119">Требование</span><span class="sxs-lookup"><span data-stu-id="628aa-119">Requirement</span></span>| <span data-ttu-id="628aa-120">Значение</span><span class="sxs-lookup"><span data-stu-id="628aa-120">Value</span></span>|
|---|---|
|[<span data-ttu-id="628aa-121">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="628aa-121">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="628aa-122">1.0</span><span class="sxs-lookup"><span data-stu-id="628aa-122">1.0</span></span>|
|[<span data-ttu-id="628aa-123">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="628aa-123">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="628aa-124">ReadItem</span><span class="sxs-lookup"><span data-stu-id="628aa-124">ReadItem</span></span>|
|[<span data-ttu-id="628aa-125">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="628aa-125">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="628aa-126">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="628aa-126">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="628aa-127">Пример</span><span class="sxs-lookup"><span data-stu-id="628aa-127">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="628aa-128">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="628aa-128">emailAddress :String</span></span>

<span data-ttu-id="628aa-129">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="628aa-129">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="628aa-130">Тип</span><span class="sxs-lookup"><span data-stu-id="628aa-130">Type</span></span>

*   <span data-ttu-id="628aa-131">String</span><span class="sxs-lookup"><span data-stu-id="628aa-131">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="628aa-132">Требования</span><span class="sxs-lookup"><span data-stu-id="628aa-132">Requirements</span></span>

|<span data-ttu-id="628aa-133">Требование</span><span class="sxs-lookup"><span data-stu-id="628aa-133">Requirement</span></span>| <span data-ttu-id="628aa-134">Значение</span><span class="sxs-lookup"><span data-stu-id="628aa-134">Value</span></span>|
|---|---|
|[<span data-ttu-id="628aa-135">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="628aa-135">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="628aa-136">1.0</span><span class="sxs-lookup"><span data-stu-id="628aa-136">1.0</span></span>|
|[<span data-ttu-id="628aa-137">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="628aa-137">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="628aa-138">ReadItem</span><span class="sxs-lookup"><span data-stu-id="628aa-138">ReadItem</span></span>|
|[<span data-ttu-id="628aa-139">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="628aa-139">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="628aa-140">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="628aa-140">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="628aa-141">Пример</span><span class="sxs-lookup"><span data-stu-id="628aa-141">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="628aa-142">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="628aa-142">timeZone :String</span></span>

<span data-ttu-id="628aa-143">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="628aa-143">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="628aa-144">Тип</span><span class="sxs-lookup"><span data-stu-id="628aa-144">Type</span></span>

*   <span data-ttu-id="628aa-145">String</span><span class="sxs-lookup"><span data-stu-id="628aa-145">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="628aa-146">Требования</span><span class="sxs-lookup"><span data-stu-id="628aa-146">Requirements</span></span>

|<span data-ttu-id="628aa-147">Требование</span><span class="sxs-lookup"><span data-stu-id="628aa-147">Requirement</span></span>| <span data-ttu-id="628aa-148">Значение</span><span class="sxs-lookup"><span data-stu-id="628aa-148">Value</span></span>|
|---|---|
|[<span data-ttu-id="628aa-149">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="628aa-149">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="628aa-150">1.0</span><span class="sxs-lookup"><span data-stu-id="628aa-150">1.0</span></span>|
|[<span data-ttu-id="628aa-151">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="628aa-151">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="628aa-152">ReadItem</span><span class="sxs-lookup"><span data-stu-id="628aa-152">ReadItem</span></span>|
|[<span data-ttu-id="628aa-153">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="628aa-153">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="628aa-154">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="628aa-154">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="628aa-155">Пример</span><span class="sxs-lookup"><span data-stu-id="628aa-155">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
