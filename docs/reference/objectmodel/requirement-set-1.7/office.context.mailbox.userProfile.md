---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.7
description: ''
ms.date: 02/15/2019
localization_priority: Normal
ms.openlocfilehash: fb55d11fd46a9957dab124514ef3bfe5a7c138eb
ms.sourcegitcommit: f26778b596b6b022814c39601485ff676ed4e2fa
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/16/2019
ms.locfileid: "30067869"
---
# <a name="userprofile"></a><span data-ttu-id="ddafe-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="ddafe-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="ddafe-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="ddafe-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="ddafe-104">Требования</span><span class="sxs-lookup"><span data-stu-id="ddafe-104">Requirements</span></span>

|<span data-ttu-id="ddafe-105">Требование</span><span class="sxs-lookup"><span data-stu-id="ddafe-105">Requirement</span></span>| <span data-ttu-id="ddafe-106">Значение</span><span class="sxs-lookup"><span data-stu-id="ddafe-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="ddafe-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="ddafe-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ddafe-108">1.0</span><span class="sxs-lookup"><span data-stu-id="ddafe-108">1.0</span></span>|
|[<span data-ttu-id="ddafe-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ddafe-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ddafe-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ddafe-110">ReadItem</span></span>|
|[<span data-ttu-id="ddafe-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ddafe-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ddafe-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ddafe-112">Compose or Read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="ddafe-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="ddafe-113">Members and methods</span></span>

| <span data-ttu-id="ddafe-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="ddafe-114">Member</span></span> | <span data-ttu-id="ddafe-115">Тип</span><span class="sxs-lookup"><span data-stu-id="ddafe-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="ddafe-116">accountType</span><span class="sxs-lookup"><span data-stu-id="ddafe-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="ddafe-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="ddafe-117">Member</span></span> |
| [<span data-ttu-id="ddafe-118">displayName</span><span class="sxs-lookup"><span data-stu-id="ddafe-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="ddafe-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="ddafe-119">Member</span></span> |
| [<span data-ttu-id="ddafe-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="ddafe-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="ddafe-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="ddafe-121">Member</span></span> |
| [<span data-ttu-id="ddafe-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="ddafe-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="ddafe-123">Член</span><span class="sxs-lookup"><span data-stu-id="ddafe-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="ddafe-124">Элементы</span><span class="sxs-lookup"><span data-stu-id="ddafe-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="ddafe-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="ddafe-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="ddafe-126">В настоящее время этот элемент поддерживается только Outlook 2016 для Mac (сборка 16.9.1212 или более поздняя).</span><span class="sxs-lookup"><span data-stu-id="ddafe-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="ddafe-127">Возвращает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="ddafe-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="ddafe-128">Возможные значения перечислены в таблице ниже.</span><span class="sxs-lookup"><span data-stu-id="ddafe-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="ddafe-129">Значение</span><span class="sxs-lookup"><span data-stu-id="ddafe-129">Value</span></span> | <span data-ttu-id="ddafe-130">Описание</span><span class="sxs-lookup"><span data-stu-id="ddafe-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="ddafe-131">Почтовый ящик размещен на локальном сервере Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="ddafe-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="ddafe-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="ddafe-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="ddafe-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="ddafe-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="ddafe-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="ddafe-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="ddafe-135">Тип</span><span class="sxs-lookup"><span data-stu-id="ddafe-135">Type</span></span>

*   <span data-ttu-id="ddafe-136">String</span><span class="sxs-lookup"><span data-stu-id="ddafe-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ddafe-137">Требования</span><span class="sxs-lookup"><span data-stu-id="ddafe-137">Requirements</span></span>

|<span data-ttu-id="ddafe-138">Требование</span><span class="sxs-lookup"><span data-stu-id="ddafe-138">Requirement</span></span>| <span data-ttu-id="ddafe-139">Значение</span><span class="sxs-lookup"><span data-stu-id="ddafe-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="ddafe-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ddafe-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ddafe-141">1.6</span><span class="sxs-lookup"><span data-stu-id="ddafe-141">1.6</span></span> |
|[<span data-ttu-id="ddafe-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ddafe-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ddafe-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ddafe-143">ReadItem</span></span>|
|[<span data-ttu-id="ddafe-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ddafe-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ddafe-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ddafe-145">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ddafe-146">Пример</span><span class="sxs-lookup"><span data-stu-id="ddafe-146">Example</span></span>

```javascript
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="ddafe-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="ddafe-147">displayName :String</span></span>

<span data-ttu-id="ddafe-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="ddafe-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="ddafe-149">Тип</span><span class="sxs-lookup"><span data-stu-id="ddafe-149">Type</span></span>

*   <span data-ttu-id="ddafe-150">String</span><span class="sxs-lookup"><span data-stu-id="ddafe-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ddafe-151">Требования</span><span class="sxs-lookup"><span data-stu-id="ddafe-151">Requirements</span></span>

|<span data-ttu-id="ddafe-152">Требование</span><span class="sxs-lookup"><span data-stu-id="ddafe-152">Requirement</span></span>| <span data-ttu-id="ddafe-153">Значение</span><span class="sxs-lookup"><span data-stu-id="ddafe-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="ddafe-154">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ddafe-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ddafe-155">1.0</span><span class="sxs-lookup"><span data-stu-id="ddafe-155">1.0</span></span>|
|[<span data-ttu-id="ddafe-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ddafe-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ddafe-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ddafe-157">ReadItem</span></span>|
|[<span data-ttu-id="ddafe-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ddafe-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ddafe-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ddafe-159">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ddafe-160">Пример</span><span class="sxs-lookup"><span data-stu-id="ddafe-160">Example</span></span>

```javascript
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="ddafe-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="ddafe-161">emailAddress :String</span></span>

<span data-ttu-id="ddafe-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="ddafe-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="ddafe-163">Тип</span><span class="sxs-lookup"><span data-stu-id="ddafe-163">Type</span></span>

*   <span data-ttu-id="ddafe-164">String</span><span class="sxs-lookup"><span data-stu-id="ddafe-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ddafe-165">Требования</span><span class="sxs-lookup"><span data-stu-id="ddafe-165">Requirements</span></span>

|<span data-ttu-id="ddafe-166">Требование</span><span class="sxs-lookup"><span data-stu-id="ddafe-166">Requirement</span></span>| <span data-ttu-id="ddafe-167">Значение</span><span class="sxs-lookup"><span data-stu-id="ddafe-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="ddafe-168">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ddafe-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ddafe-169">1.0</span><span class="sxs-lookup"><span data-stu-id="ddafe-169">1.0</span></span>|
|[<span data-ttu-id="ddafe-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ddafe-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ddafe-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ddafe-171">ReadItem</span></span>|
|[<span data-ttu-id="ddafe-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ddafe-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ddafe-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ddafe-173">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ddafe-174">Пример</span><span class="sxs-lookup"><span data-stu-id="ddafe-174">Example</span></span>

```javascript
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="ddafe-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="ddafe-175">timeZone :String</span></span>

<span data-ttu-id="ddafe-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="ddafe-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="ddafe-177">Тип</span><span class="sxs-lookup"><span data-stu-id="ddafe-177">Type</span></span>

*   <span data-ttu-id="ddafe-178">String</span><span class="sxs-lookup"><span data-stu-id="ddafe-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="ddafe-179">Требования</span><span class="sxs-lookup"><span data-stu-id="ddafe-179">Requirements</span></span>

|<span data-ttu-id="ddafe-180">Требование</span><span class="sxs-lookup"><span data-stu-id="ddafe-180">Requirement</span></span>| <span data-ttu-id="ddafe-181">Значение</span><span class="sxs-lookup"><span data-stu-id="ddafe-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="ddafe-182">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="ddafe-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="ddafe-183">1.0</span><span class="sxs-lookup"><span data-stu-id="ddafe-183">1.0</span></span>|
|[<span data-ttu-id="ddafe-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="ddafe-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="ddafe-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="ddafe-185">ReadItem</span></span>|
|[<span data-ttu-id="ddafe-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="ddafe-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="ddafe-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="ddafe-187">Compose or Read</span></span>|

##### <a name="example"></a><span data-ttu-id="ddafe-188">Пример</span><span class="sxs-lookup"><span data-stu-id="ddafe-188">Example</span></span>

```javascript
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
