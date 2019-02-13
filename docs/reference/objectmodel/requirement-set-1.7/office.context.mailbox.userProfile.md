---
title: Office.context.mailbox.userProfile — набор обязательных элементов 1.7
description: ''
ms.date: 10/31/2018
localization_priority: Normal
ms.openlocfilehash: b07ff5bee3adc18cc1006bb574e373182b29f5fe
ms.sourcegitcommit: a59f4e322238efa187f388a75b7709462c71e668
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/13/2019
ms.locfileid: "29635904"
---
# <a name="userprofile"></a><span data-ttu-id="c1f60-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="c1f60-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="c1f60-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="c1f60-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1f60-104">Требования</span><span class="sxs-lookup"><span data-stu-id="c1f60-104">Requirements</span></span>

|<span data-ttu-id="c1f60-105">Требование</span><span class="sxs-lookup"><span data-stu-id="c1f60-105">Requirement</span></span>| <span data-ttu-id="c1f60-106">Значение</span><span class="sxs-lookup"><span data-stu-id="c1f60-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1f60-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="c1f60-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1f60-108">1.0</span><span class="sxs-lookup"><span data-stu-id="c1f60-108">1.0</span></span>|
|[<span data-ttu-id="c1f60-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c1f60-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1f60-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1f60-110">ReadItem</span></span>|
|[<span data-ttu-id="c1f60-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c1f60-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1f60-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c1f60-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="c1f60-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="c1f60-113">Members and methods</span></span>

| <span data-ttu-id="c1f60-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="c1f60-114">Member</span></span> | <span data-ttu-id="c1f60-115">Тип</span><span class="sxs-lookup"><span data-stu-id="c1f60-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="c1f60-116">accountType</span><span class="sxs-lookup"><span data-stu-id="c1f60-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="c1f60-117">Member</span><span class="sxs-lookup"><span data-stu-id="c1f60-117">Member</span></span> |
| [<span data-ttu-id="c1f60-118">displayName</span><span class="sxs-lookup"><span data-stu-id="c1f60-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="c1f60-119">Member</span><span class="sxs-lookup"><span data-stu-id="c1f60-119">Member</span></span> |
| [<span data-ttu-id="c1f60-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="c1f60-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="c1f60-121">Member</span><span class="sxs-lookup"><span data-stu-id="c1f60-121">Member</span></span> |
| [<span data-ttu-id="c1f60-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="c1f60-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="c1f60-123">Член</span><span class="sxs-lookup"><span data-stu-id="c1f60-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="c1f60-124">Members</span><span class="sxs-lookup"><span data-stu-id="c1f60-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="c1f60-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="c1f60-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="c1f60-126">Этот член в данный момент поддерживается только Outlook 2016 для Mac (построение 16.9.1212 или более поздней версии).</span><span class="sxs-lookup"><span data-stu-id="c1f60-126">This member is currently only supported by Outlook 2016 for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="c1f60-127">Возвращает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="c1f60-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="c1f60-128">Возможные значения перечислены в таблице ниже.</span><span class="sxs-lookup"><span data-stu-id="c1f60-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="c1f60-129">Значение</span><span class="sxs-lookup"><span data-stu-id="c1f60-129">Value</span></span> | <span data-ttu-id="c1f60-130">Описание</span><span class="sxs-lookup"><span data-stu-id="c1f60-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="c1f60-131">Почтовый ящик размещен на локальном сервере Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="c1f60-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="c1f60-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="c1f60-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="c1f60-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="c1f60-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="c1f60-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="c1f60-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="c1f60-135">Тип:</span><span class="sxs-lookup"><span data-stu-id="c1f60-135">Type:</span></span>

*   <span data-ttu-id="c1f60-136">String</span><span class="sxs-lookup"><span data-stu-id="c1f60-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1f60-137">Требования</span><span class="sxs-lookup"><span data-stu-id="c1f60-137">Requirements</span></span>

|<span data-ttu-id="c1f60-138">Требование</span><span class="sxs-lookup"><span data-stu-id="c1f60-138">Requirement</span></span>| <span data-ttu-id="c1f60-139">Значение</span><span class="sxs-lookup"><span data-stu-id="c1f60-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1f60-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="c1f60-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1f60-141">1.6</span><span class="sxs-lookup"><span data-stu-id="c1f60-141">1.6</span></span> |
|[<span data-ttu-id="c1f60-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c1f60-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1f60-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1f60-143">ReadItem</span></span>|
|[<span data-ttu-id="c1f60-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c1f60-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1f60-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c1f60-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1f60-146">Пример</span><span class="sxs-lookup"><span data-stu-id="c1f60-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="c1f60-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="c1f60-147">displayName :String</span></span>

<span data-ttu-id="c1f60-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="c1f60-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="c1f60-149">Тип:</span><span class="sxs-lookup"><span data-stu-id="c1f60-149">Type:</span></span>

*   <span data-ttu-id="c1f60-150">String</span><span class="sxs-lookup"><span data-stu-id="c1f60-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1f60-151">Требования</span><span class="sxs-lookup"><span data-stu-id="c1f60-151">Requirements</span></span>

|<span data-ttu-id="c1f60-152">Требование</span><span class="sxs-lookup"><span data-stu-id="c1f60-152">Requirement</span></span>| <span data-ttu-id="c1f60-153">Значение</span><span class="sxs-lookup"><span data-stu-id="c1f60-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1f60-154">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="c1f60-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1f60-155">1.0</span><span class="sxs-lookup"><span data-stu-id="c1f60-155">1.0</span></span>|
|[<span data-ttu-id="c1f60-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c1f60-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1f60-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1f60-157">ReadItem</span></span>|
|[<span data-ttu-id="c1f60-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c1f60-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1f60-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c1f60-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1f60-160">Пример</span><span class="sxs-lookup"><span data-stu-id="c1f60-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="c1f60-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="c1f60-161">emailAddress :String</span></span>

<span data-ttu-id="c1f60-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="c1f60-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="c1f60-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="c1f60-163">Type:</span></span>

*   <span data-ttu-id="c1f60-164">String</span><span class="sxs-lookup"><span data-stu-id="c1f60-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1f60-165">Требования</span><span class="sxs-lookup"><span data-stu-id="c1f60-165">Requirements</span></span>

|<span data-ttu-id="c1f60-166">Требование</span><span class="sxs-lookup"><span data-stu-id="c1f60-166">Requirement</span></span>| <span data-ttu-id="c1f60-167">Значение</span><span class="sxs-lookup"><span data-stu-id="c1f60-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1f60-168">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="c1f60-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1f60-169">1.0</span><span class="sxs-lookup"><span data-stu-id="c1f60-169">1.0</span></span>|
|[<span data-ttu-id="c1f60-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c1f60-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1f60-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1f60-171">ReadItem</span></span>|
|[<span data-ttu-id="c1f60-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c1f60-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1f60-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c1f60-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1f60-174">Пример</span><span class="sxs-lookup"><span data-stu-id="c1f60-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="c1f60-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="c1f60-175">timeZone :String</span></span>

<span data-ttu-id="c1f60-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="c1f60-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="c1f60-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="c1f60-177">Type:</span></span>

*   <span data-ttu-id="c1f60-178">String</span><span class="sxs-lookup"><span data-stu-id="c1f60-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="c1f60-179">Требования</span><span class="sxs-lookup"><span data-stu-id="c1f60-179">Requirements</span></span>

|<span data-ttu-id="c1f60-180">Требование</span><span class="sxs-lookup"><span data-stu-id="c1f60-180">Requirement</span></span>| <span data-ttu-id="c1f60-181">Значение</span><span class="sxs-lookup"><span data-stu-id="c1f60-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="c1f60-182">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="c1f60-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="c1f60-183">1.0</span><span class="sxs-lookup"><span data-stu-id="c1f60-183">1.0</span></span>|
|[<span data-ttu-id="c1f60-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="c1f60-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="c1f60-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="c1f60-185">ReadItem</span></span>|
|[<span data-ttu-id="c1f60-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="c1f60-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="c1f60-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="c1f60-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="c1f60-188">Пример</span><span class="sxs-lookup"><span data-stu-id="c1f60-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```
