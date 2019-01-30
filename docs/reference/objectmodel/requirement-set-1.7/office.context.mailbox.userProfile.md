---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.7
description: ''
ms.date: 10/31/2018
localization_priority: Normal
ms.openlocfilehash: b07ff5bee3adc18cc1006bb574e373182b29f5fe
ms.sourcegitcommit: 2e4b97f0252ff3dd908a3aa7a9720f0cb50b855d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/30/2019
ms.locfileid: "29635904"
---
# <a name="userprofile"></a><span data-ttu-id="a33c6-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="a33c6-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="a33c6-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="a33c6-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="a33c6-104">Требования</span><span class="sxs-lookup"><span data-stu-id="a33c6-104">Requirements</span></span>

|<span data-ttu-id="a33c6-105">Требование</span><span class="sxs-lookup"><span data-stu-id="a33c6-105">Requirement</span></span>| <span data-ttu-id="a33c6-106">Значение</span><span class="sxs-lookup"><span data-stu-id="a33c6-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="a33c6-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a33c6-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a33c6-108">1.0</span><span class="sxs-lookup"><span data-stu-id="a33c6-108">1.0</span></span>|
|[<span data-ttu-id="a33c6-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a33c6-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a33c6-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a33c6-110">ReadItem</span></span>|
|[<span data-ttu-id="a33c6-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a33c6-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a33c6-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a33c6-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="a33c6-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="a33c6-113">Members and methods</span></span>

| <span data-ttu-id="a33c6-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="a33c6-114">Member</span></span> | <span data-ttu-id="a33c6-115">Тип</span><span class="sxs-lookup"><span data-stu-id="a33c6-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="a33c6-116">accountType</span><span class="sxs-lookup"><span data-stu-id="a33c6-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="a33c6-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="a33c6-117">Member</span></span> |
| [<span data-ttu-id="a33c6-118">displayName</span><span class="sxs-lookup"><span data-stu-id="a33c6-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="a33c6-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="a33c6-119">Member</span></span> |
| [<span data-ttu-id="a33c6-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="a33c6-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="a33c6-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="a33c6-121">Member</span></span> |
| [<span data-ttu-id="a33c6-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="a33c6-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="a33c6-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="a33c6-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="a33c6-124">Элементы</span><span class="sxs-lookup"><span data-stu-id="a33c6-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="a33c6-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="a33c6-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="a33c6-126">Этот член в данный момент поддерживается только Outlook 2016 для Mac (построение 16.9.1212 или более поздней версии).</span><span class="sxs-lookup"><span data-stu-id="a33c6-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="a33c6-127">Возвращает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="a33c6-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="a33c6-128">Возможные значения перечислены в таблице ниже.</span><span class="sxs-lookup"><span data-stu-id="a33c6-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="a33c6-129">Значение</span><span class="sxs-lookup"><span data-stu-id="a33c6-129">Value</span></span> | <span data-ttu-id="a33c6-130">Описание</span><span class="sxs-lookup"><span data-stu-id="a33c6-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="a33c6-131">Почтовый ящик размещен на локальном сервере Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="a33c6-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="a33c6-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="a33c6-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="a33c6-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="a33c6-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="a33c6-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="a33c6-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="a33c6-135">Тип:</span><span class="sxs-lookup"><span data-stu-id="a33c6-135">Type:</span></span>

*   <span data-ttu-id="a33c6-136">String</span><span class="sxs-lookup"><span data-stu-id="a33c6-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a33c6-137">Требования</span><span class="sxs-lookup"><span data-stu-id="a33c6-137">Requirements</span></span>

|<span data-ttu-id="a33c6-138">Требование</span><span class="sxs-lookup"><span data-stu-id="a33c6-138">Requirement</span></span>| <span data-ttu-id="a33c6-139">Значение</span><span class="sxs-lookup"><span data-stu-id="a33c6-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="a33c6-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="a33c6-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a33c6-141">1.6</span><span class="sxs-lookup"><span data-stu-id="a33c6-141">1.6</span></span> |
|[<span data-ttu-id="a33c6-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a33c6-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a33c6-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a33c6-143">ReadItem</span></span>|
|[<span data-ttu-id="a33c6-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a33c6-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a33c6-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a33c6-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a33c6-146">Пример</span><span class="sxs-lookup"><span data-stu-id="a33c6-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="a33c6-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="a33c6-147">displayName :String</span></span>

<span data-ttu-id="a33c6-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="a33c6-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="a33c6-149">Тип:</span><span class="sxs-lookup"><span data-stu-id="a33c6-149">Type:</span></span>

*   <span data-ttu-id="a33c6-150">String</span><span class="sxs-lookup"><span data-stu-id="a33c6-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a33c6-151">Требования</span><span class="sxs-lookup"><span data-stu-id="a33c6-151">Requirements</span></span>

|<span data-ttu-id="a33c6-152">Требование</span><span class="sxs-lookup"><span data-stu-id="a33c6-152">Requirement</span></span>| <span data-ttu-id="a33c6-153">Значение</span><span class="sxs-lookup"><span data-stu-id="a33c6-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="a33c6-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a33c6-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a33c6-155">1.0</span><span class="sxs-lookup"><span data-stu-id="a33c6-155">1.0</span></span>|
|[<span data-ttu-id="a33c6-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a33c6-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a33c6-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a33c6-157">ReadItem</span></span>|
|[<span data-ttu-id="a33c6-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a33c6-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a33c6-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a33c6-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a33c6-160">Пример</span><span class="sxs-lookup"><span data-stu-id="a33c6-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="a33c6-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="a33c6-161">emailAddress :String</span></span>

<span data-ttu-id="a33c6-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="a33c6-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="a33c6-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="a33c6-163">Type:</span></span>

*   <span data-ttu-id="a33c6-164">String</span><span class="sxs-lookup"><span data-stu-id="a33c6-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a33c6-165">Требования</span><span class="sxs-lookup"><span data-stu-id="a33c6-165">Requirements</span></span>

|<span data-ttu-id="a33c6-166">Требование</span><span class="sxs-lookup"><span data-stu-id="a33c6-166">Requirement</span></span>| <span data-ttu-id="a33c6-167">Значение</span><span class="sxs-lookup"><span data-stu-id="a33c6-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="a33c6-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a33c6-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a33c6-169">1.0</span><span class="sxs-lookup"><span data-stu-id="a33c6-169">1.0</span></span>|
|[<span data-ttu-id="a33c6-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a33c6-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a33c6-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a33c6-171">ReadItem</span></span>|
|[<span data-ttu-id="a33c6-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a33c6-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a33c6-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a33c6-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a33c6-174">Пример</span><span class="sxs-lookup"><span data-stu-id="a33c6-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="a33c6-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="a33c6-175">timeZone :String</span></span>

<span data-ttu-id="a33c6-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="a33c6-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="a33c6-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="a33c6-177">Type:</span></span>

*   <span data-ttu-id="a33c6-178">String</span><span class="sxs-lookup"><span data-stu-id="a33c6-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="a33c6-179">Требования</span><span class="sxs-lookup"><span data-stu-id="a33c6-179">Requirements</span></span>

|<span data-ttu-id="a33c6-180">Требование</span><span class="sxs-lookup"><span data-stu-id="a33c6-180">Requirement</span></span>| <span data-ttu-id="a33c6-181">Значение</span><span class="sxs-lookup"><span data-stu-id="a33c6-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="a33c6-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="a33c6-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="a33c6-183">1.0</span><span class="sxs-lookup"><span data-stu-id="a33c6-183">1.0</span></span>|
|[<span data-ttu-id="a33c6-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="a33c6-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="a33c6-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="a33c6-185">ReadItem</span></span>|
|[<span data-ttu-id="a33c6-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="a33c6-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="a33c6-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="a33c6-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="a33c6-188">Пример</span><span class="sxs-lookup"><span data-stu-id="a33c6-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
