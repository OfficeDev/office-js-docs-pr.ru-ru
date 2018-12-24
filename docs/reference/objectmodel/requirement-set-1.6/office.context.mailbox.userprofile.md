---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.6
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: fe30a390583dc646e9c8792710c580d02c373a1a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432900"
---
# <a name="userprofile"></a><span data-ttu-id="2b218-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="2b218-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="2b218-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="2b218-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b218-104">Требования</span><span class="sxs-lookup"><span data-stu-id="2b218-104">Requirements</span></span>

|<span data-ttu-id="2b218-105">Требование</span><span class="sxs-lookup"><span data-stu-id="2b218-105">Requirement</span></span>| <span data-ttu-id="2b218-106">Значение</span><span class="sxs-lookup"><span data-stu-id="2b218-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b218-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2b218-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b218-108">1.0</span><span class="sxs-lookup"><span data-stu-id="2b218-108">1.0</span></span>|
|[<span data-ttu-id="2b218-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2b218-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2b218-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2b218-110">ReadItem</span></span>|
|[<span data-ttu-id="2b218-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2b218-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2b218-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2b218-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="2b218-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="2b218-113">Members and methods</span></span>

| <span data-ttu-id="2b218-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="2b218-114">Member</span></span> | <span data-ttu-id="2b218-115">Тип</span><span class="sxs-lookup"><span data-stu-id="2b218-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="2b218-116">accountType</span><span class="sxs-lookup"><span data-stu-id="2b218-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="2b218-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="2b218-117">Member</span></span> |
| [<span data-ttu-id="2b218-118">displayName</span><span class="sxs-lookup"><span data-stu-id="2b218-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="2b218-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="2b218-119">Member</span></span> |
| [<span data-ttu-id="2b218-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="2b218-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="2b218-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="2b218-121">Member</span></span> |
| [<span data-ttu-id="2b218-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="2b218-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="2b218-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="2b218-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="2b218-124">Элементы</span><span class="sxs-lookup"><span data-stu-id="2b218-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="2b218-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="2b218-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="2b218-126">В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии для Mac (сборка 16.9.1212 или более поздняя версия).</span><span class="sxs-lookup"><span data-stu-id="2b218-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="2b218-127">Возвращает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="2b218-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="2b218-128">Возможные значения перечислены в таблице ниже.</span><span class="sxs-lookup"><span data-stu-id="2b218-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="2b218-129">Значение</span><span class="sxs-lookup"><span data-stu-id="2b218-129">Value</span></span> | <span data-ttu-id="2b218-130">Описание</span><span class="sxs-lookup"><span data-stu-id="2b218-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="2b218-131">Почтовый ящик размещен на локальном сервере Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="2b218-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="2b218-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="2b218-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="2b218-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="2b218-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="2b218-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="2b218-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="2b218-135">Тип:</span><span class="sxs-lookup"><span data-stu-id="2b218-135">Type:</span></span>

*   <span data-ttu-id="2b218-136">String</span><span class="sxs-lookup"><span data-stu-id="2b218-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b218-137">Требования</span><span class="sxs-lookup"><span data-stu-id="2b218-137">Requirements</span></span>

|<span data-ttu-id="2b218-138">Требование</span><span class="sxs-lookup"><span data-stu-id="2b218-138">Requirement</span></span>| <span data-ttu-id="2b218-139">Значение</span><span class="sxs-lookup"><span data-stu-id="2b218-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b218-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="2b218-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b218-141">1.6</span><span class="sxs-lookup"><span data-stu-id="2b218-141">1.6</span></span> |
|[<span data-ttu-id="2b218-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2b218-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2b218-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2b218-143">ReadItem</span></span>|
|[<span data-ttu-id="2b218-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2b218-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2b218-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2b218-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b218-146">Пример</span><span class="sxs-lookup"><span data-stu-id="2b218-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="2b218-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="2b218-147">displayName :String</span></span>

<span data-ttu-id="2b218-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="2b218-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="2b218-149">Тип:</span><span class="sxs-lookup"><span data-stu-id="2b218-149">Type:</span></span>

*   <span data-ttu-id="2b218-150">String</span><span class="sxs-lookup"><span data-stu-id="2b218-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b218-151">Требования</span><span class="sxs-lookup"><span data-stu-id="2b218-151">Requirements</span></span>

|<span data-ttu-id="2b218-152">Требование</span><span class="sxs-lookup"><span data-stu-id="2b218-152">Requirement</span></span>| <span data-ttu-id="2b218-153">Значение</span><span class="sxs-lookup"><span data-stu-id="2b218-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b218-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2b218-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b218-155">1.0</span><span class="sxs-lookup"><span data-stu-id="2b218-155">1.0</span></span>|
|[<span data-ttu-id="2b218-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2b218-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2b218-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2b218-157">ReadItem</span></span>|
|[<span data-ttu-id="2b218-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2b218-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2b218-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2b218-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b218-160">Пример</span><span class="sxs-lookup"><span data-stu-id="2b218-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="2b218-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="2b218-161">emailAddress :String</span></span>

<span data-ttu-id="2b218-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="2b218-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="2b218-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="2b218-163">Type:</span></span>

*   <span data-ttu-id="2b218-164">String</span><span class="sxs-lookup"><span data-stu-id="2b218-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b218-165">Требования</span><span class="sxs-lookup"><span data-stu-id="2b218-165">Requirements</span></span>

|<span data-ttu-id="2b218-166">Требование</span><span class="sxs-lookup"><span data-stu-id="2b218-166">Requirement</span></span>| <span data-ttu-id="2b218-167">Значение</span><span class="sxs-lookup"><span data-stu-id="2b218-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b218-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2b218-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b218-169">1.0</span><span class="sxs-lookup"><span data-stu-id="2b218-169">1.0</span></span>|
|[<span data-ttu-id="2b218-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2b218-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2b218-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2b218-171">ReadItem</span></span>|
|[<span data-ttu-id="2b218-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2b218-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2b218-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2b218-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b218-174">Пример</span><span class="sxs-lookup"><span data-stu-id="2b218-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="2b218-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="2b218-175">timeZone :String</span></span>

<span data-ttu-id="2b218-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="2b218-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="2b218-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="2b218-177">Type:</span></span>

*   <span data-ttu-id="2b218-178">String</span><span class="sxs-lookup"><span data-stu-id="2b218-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="2b218-179">Требования</span><span class="sxs-lookup"><span data-stu-id="2b218-179">Requirements</span></span>

|<span data-ttu-id="2b218-180">Требование</span><span class="sxs-lookup"><span data-stu-id="2b218-180">Requirement</span></span>| <span data-ttu-id="2b218-181">Значение</span><span class="sxs-lookup"><span data-stu-id="2b218-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="2b218-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="2b218-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="2b218-183">1.0</span><span class="sxs-lookup"><span data-stu-id="2b218-183">1.0</span></span>|
|[<span data-ttu-id="2b218-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="2b218-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="2b218-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="2b218-185">ReadItem</span></span>|
|[<span data-ttu-id="2b218-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="2b218-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="2b218-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="2b218-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="2b218-188">Пример</span><span class="sxs-lookup"><span data-stu-id="2b218-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```