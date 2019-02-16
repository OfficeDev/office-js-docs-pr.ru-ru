---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.6
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: 09457a41fe68ae03e035d3d3f4b80b139be348e0
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067876"
---
# <a name="userprofile"></a><span data-ttu-id="8be21-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="8be21-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="8be21-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="8be21-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="8be21-104">Требования</span><span class="sxs-lookup"><span data-stu-id="8be21-104">Requirements</span></span>

|<span data-ttu-id="8be21-105">Требование</span><span class="sxs-lookup"><span data-stu-id="8be21-105">Requirement</span></span>| <span data-ttu-id="8be21-106">Значение</span><span class="sxs-lookup"><span data-stu-id="8be21-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be21-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="8be21-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be21-108">1.0</span><span class="sxs-lookup"><span data-stu-id="8be21-108">1.0</span></span>|
|[<span data-ttu-id="8be21-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8be21-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be21-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be21-110">ReadItem</span></span>|
|[<span data-ttu-id="8be21-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8be21-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be21-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8be21-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="8be21-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="8be21-113">Members and methods</span></span>

| <span data-ttu-id="8be21-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="8be21-114">Member</span></span> | <span data-ttu-id="8be21-115">Тип</span><span class="sxs-lookup"><span data-stu-id="8be21-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="8be21-116">accountType</span><span class="sxs-lookup"><span data-stu-id="8be21-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="8be21-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="8be21-117">Member</span></span> |
| [<span data-ttu-id="8be21-118">displayName</span><span class="sxs-lookup"><span data-stu-id="8be21-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="8be21-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="8be21-119">Member</span></span> |
| [<span data-ttu-id="8be21-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="8be21-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="8be21-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="8be21-121">Member</span></span> |
| [<span data-ttu-id="8be21-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="8be21-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="8be21-123">Член</span><span class="sxs-lookup"><span data-stu-id="8be21-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="8be21-124">Элементы</span><span class="sxs-lookup"><span data-stu-id="8be21-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="8be21-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="8be21-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="8be21-126">В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии для Mac (сборка 16.9.1212 или более поздняя версия).</span><span class="sxs-lookup"><span data-stu-id="8be21-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="8be21-127">Возвращает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="8be21-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="8be21-128">Возможные значения перечислены в таблице ниже.</span><span class="sxs-lookup"><span data-stu-id="8be21-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="8be21-129">Значение</span><span class="sxs-lookup"><span data-stu-id="8be21-129">Value</span></span> | <span data-ttu-id="8be21-130">Описание</span><span class="sxs-lookup"><span data-stu-id="8be21-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="8be21-131">Почтовый ящик размещен на локальном сервере Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="8be21-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="8be21-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="8be21-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="8be21-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="8be21-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="8be21-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="8be21-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="8be21-135">Тип</span><span class="sxs-lookup"><span data-stu-id="8be21-135">Type</span></span>

*   <span data-ttu-id="8be21-136">String</span><span class="sxs-lookup"><span data-stu-id="8be21-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8be21-137">Требования</span><span class="sxs-lookup"><span data-stu-id="8be21-137">Requirements</span></span>

|<span data-ttu-id="8be21-138">Требование</span><span class="sxs-lookup"><span data-stu-id="8be21-138">Requirement</span></span>| <span data-ttu-id="8be21-139">Значение</span><span class="sxs-lookup"><span data-stu-id="8be21-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be21-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8be21-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be21-141">1.6</span><span class="sxs-lookup"><span data-stu-id="8be21-141">1.6</span></span> |
|[<span data-ttu-id="8be21-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8be21-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be21-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be21-143">ReadItem</span></span>|
|[<span data-ttu-id="8be21-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8be21-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be21-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8be21-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8be21-146">Пример</span><span class="sxs-lookup"><span data-stu-id="8be21-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="8be21-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="8be21-147">displayName :String</span></span>

<span data-ttu-id="8be21-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="8be21-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="8be21-149">Тип</span><span class="sxs-lookup"><span data-stu-id="8be21-149">Type</span></span>

*   <span data-ttu-id="8be21-150">String</span><span class="sxs-lookup"><span data-stu-id="8be21-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8be21-151">Требования</span><span class="sxs-lookup"><span data-stu-id="8be21-151">Requirements</span></span>

|<span data-ttu-id="8be21-152">Требование</span><span class="sxs-lookup"><span data-stu-id="8be21-152">Requirement</span></span>| <span data-ttu-id="8be21-153">Значение</span><span class="sxs-lookup"><span data-stu-id="8be21-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be21-154">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8be21-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be21-155">1.0</span><span class="sxs-lookup"><span data-stu-id="8be21-155">1.0</span></span>|
|[<span data-ttu-id="8be21-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8be21-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be21-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be21-157">ReadItem</span></span>|
|[<span data-ttu-id="8be21-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8be21-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be21-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8be21-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8be21-160">Пример</span><span class="sxs-lookup"><span data-stu-id="8be21-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="8be21-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="8be21-161">emailAddress :String</span></span>

<span data-ttu-id="8be21-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="8be21-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="8be21-163">Тип</span><span class="sxs-lookup"><span data-stu-id="8be21-163">Type</span></span>

*   <span data-ttu-id="8be21-164">String</span><span class="sxs-lookup"><span data-stu-id="8be21-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8be21-165">Требования</span><span class="sxs-lookup"><span data-stu-id="8be21-165">Requirements</span></span>

|<span data-ttu-id="8be21-166">Требование</span><span class="sxs-lookup"><span data-stu-id="8be21-166">Requirement</span></span>| <span data-ttu-id="8be21-167">Значение</span><span class="sxs-lookup"><span data-stu-id="8be21-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be21-168">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8be21-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be21-169">1.0</span><span class="sxs-lookup"><span data-stu-id="8be21-169">1.0</span></span>|
|[<span data-ttu-id="8be21-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8be21-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be21-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be21-171">ReadItem</span></span>|
|[<span data-ttu-id="8be21-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8be21-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be21-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8be21-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8be21-174">Пример</span><span class="sxs-lookup"><span data-stu-id="8be21-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="8be21-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="8be21-175">timeZone :String</span></span>

<span data-ttu-id="8be21-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="8be21-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="8be21-177">Тип</span><span class="sxs-lookup"><span data-stu-id="8be21-177">Type</span></span>

*   <span data-ttu-id="8be21-178">String</span><span class="sxs-lookup"><span data-stu-id="8be21-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="8be21-179">Требования</span><span class="sxs-lookup"><span data-stu-id="8be21-179">Requirements</span></span>

|<span data-ttu-id="8be21-180">Требование</span><span class="sxs-lookup"><span data-stu-id="8be21-180">Requirement</span></span>| <span data-ttu-id="8be21-181">Значение</span><span class="sxs-lookup"><span data-stu-id="8be21-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="8be21-182">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="8be21-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="8be21-183">1.0</span><span class="sxs-lookup"><span data-stu-id="8be21-183">1.0</span></span>|
|[<span data-ttu-id="8be21-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="8be21-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="8be21-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="8be21-185">ReadItem</span></span>|
|[<span data-ttu-id="8be21-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="8be21-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="8be21-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="8be21-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="8be21-188">Пример</span><span class="sxs-lookup"><span data-stu-id="8be21-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
