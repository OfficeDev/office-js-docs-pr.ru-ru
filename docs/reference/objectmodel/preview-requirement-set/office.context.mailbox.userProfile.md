---
title: Office.context.mailbox.userProfile — предварительная версия набора обязательных элементов
description: ''
ms.date: 10/31/2018
ms.openlocfilehash: 061ee8367005f4af0795c4d9e1236d0b2443521a
ms.sourcegitcommit: 6f53df6f3ee91e084cd5160bb48afbbd49743b7e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/22/2018
ms.locfileid: "27432818"
---
# <a name="userprofile"></a><span data-ttu-id="79d43-102">userProfile</span><span class="sxs-lookup"><span data-stu-id="79d43-102">userProfile</span></span>

### <a name="officeofficemdcontextofficecontextmdmailboxofficecontextmailboxmduserprofile"></a><span data-ttu-id="79d43-103">[Office](Office.md)[.context](Office.context.md)[.mailbox](Office.context.mailbox.md).userProfile</span><span class="sxs-lookup"><span data-stu-id="79d43-103">Office.context.mailbox.userProfile</span></span>

##### <a name="requirements"></a><span data-ttu-id="79d43-104">Требования</span><span class="sxs-lookup"><span data-stu-id="79d43-104">Requirements</span></span>

|<span data-ttu-id="79d43-105">Требование</span><span class="sxs-lookup"><span data-stu-id="79d43-105">Requirement</span></span>| <span data-ttu-id="79d43-106">Значение</span><span class="sxs-lookup"><span data-stu-id="79d43-106">Value</span></span>|
|---|---|
|[<span data-ttu-id="79d43-107">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="79d43-107">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79d43-108">1.0</span><span class="sxs-lookup"><span data-stu-id="79d43-108">1.0</span></span>|
|[<span data-ttu-id="79d43-109">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="79d43-109">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79d43-110">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79d43-110">ReadItem</span></span>|
|[<span data-ttu-id="79d43-111">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="79d43-111">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79d43-112">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="79d43-112">Compose or read</span></span>|

##### <a name="members-and-methods"></a><span data-ttu-id="79d43-113">Элементы и методы</span><span class="sxs-lookup"><span data-stu-id="79d43-113">Members and methods</span></span>

| <span data-ttu-id="79d43-114">Элемент</span><span class="sxs-lookup"><span data-stu-id="79d43-114">Member</span></span> | <span data-ttu-id="79d43-115">Тип</span><span class="sxs-lookup"><span data-stu-id="79d43-115">Type</span></span> |
|--------|------|
| [<span data-ttu-id="79d43-116">accountType</span><span class="sxs-lookup"><span data-stu-id="79d43-116">accountType</span></span>](#accounttype-string) | <span data-ttu-id="79d43-117">Элемент</span><span class="sxs-lookup"><span data-stu-id="79d43-117">Member</span></span> |
| [<span data-ttu-id="79d43-118">displayName</span><span class="sxs-lookup"><span data-stu-id="79d43-118">displayName</span></span>](#displayname-string) | <span data-ttu-id="79d43-119">Элемент</span><span class="sxs-lookup"><span data-stu-id="79d43-119">Member</span></span> |
| [<span data-ttu-id="79d43-120">emailAddress</span><span class="sxs-lookup"><span data-stu-id="79d43-120">emailAddress</span></span>](#emailaddress-string) | <span data-ttu-id="79d43-121">Элемент</span><span class="sxs-lookup"><span data-stu-id="79d43-121">Member</span></span> |
| [<span data-ttu-id="79d43-122">timeZone</span><span class="sxs-lookup"><span data-stu-id="79d43-122">timeZone</span></span>](#timezone-string) | <span data-ttu-id="79d43-123">Элемент</span><span class="sxs-lookup"><span data-stu-id="79d43-123">Member</span></span> |

### <a name="members"></a><span data-ttu-id="79d43-124">Элементы</span><span class="sxs-lookup"><span data-stu-id="79d43-124">Members</span></span>

####  <a name="accounttype-string"></a><span data-ttu-id="79d43-125">accountType :String</span><span class="sxs-lookup"><span data-stu-id="79d43-125">accountType :String</span></span>

> [!NOTE]
> <span data-ttu-id="79d43-126">В настоящее время этот элемент поддерживается только в Outlook 2016 или более поздней версии для Mac (сборка 16.9.1212 или более поздняя версия).</span><span class="sxs-lookup"><span data-stu-id="79d43-126">This member is currently only supported in Outlook 2016 or later for Mac (build 16.9.1212 or later).</span></span>

<span data-ttu-id="79d43-127">Возвращает тип учетной записи пользователя, связанной с почтовым ящиком.</span><span class="sxs-lookup"><span data-stu-id="79d43-127">Gets the account type of the user associated with the mailbox.</span></span> <span data-ttu-id="79d43-128">Возможные значения перечислены в таблице ниже.</span><span class="sxs-lookup"><span data-stu-id="79d43-128">The possible values are listed in the following table.</span></span>

| <span data-ttu-id="79d43-129">Значение</span><span class="sxs-lookup"><span data-stu-id="79d43-129">Value</span></span> | <span data-ttu-id="79d43-130">Описание</span><span class="sxs-lookup"><span data-stu-id="79d43-130">Description</span></span> |
|-------|-------------|
| `enterprise` | <span data-ttu-id="79d43-131">Почтовый ящик размещен на локальном сервере Exchange Server.</span><span class="sxs-lookup"><span data-stu-id="79d43-131">The mailbox is on an on-premises Exchange server.</span></span> |
| `gmail` | <span data-ttu-id="79d43-132">Почтовый ящик связан с учетной записью Gmail.</span><span class="sxs-lookup"><span data-stu-id="79d43-132">The mailbox is associated with a Gmail account.</span></span> |
| `office365` | <span data-ttu-id="79d43-133">Почтовый ящик связан с рабочей или учебной учетной записью Office 365.</span><span class="sxs-lookup"><span data-stu-id="79d43-133">The mailbox is associated with an Office 365 work or school account.</span></span> |
| `outlookCom` | <span data-ttu-id="79d43-134">Почтовый ящик связан с личной учетной записью Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="79d43-134">The mailbox is associated with a personal Outlook.com account.</span></span> |

##### <a name="type"></a><span data-ttu-id="79d43-135">Тип:</span><span class="sxs-lookup"><span data-stu-id="79d43-135">Type:</span></span>

*   <span data-ttu-id="79d43-136">String</span><span class="sxs-lookup"><span data-stu-id="79d43-136">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79d43-137">Требования</span><span class="sxs-lookup"><span data-stu-id="79d43-137">Requirements</span></span>

|<span data-ttu-id="79d43-138">Требование</span><span class="sxs-lookup"><span data-stu-id="79d43-138">Requirement</span></span>| <span data-ttu-id="79d43-139">Значение</span><span class="sxs-lookup"><span data-stu-id="79d43-139">Value</span></span>|
|---|---|
|[<span data-ttu-id="79d43-140">Минимальная версия набора обязательных элементов для почтового ящика</span><span class="sxs-lookup"><span data-stu-id="79d43-140">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79d43-141">1.6</span><span class="sxs-lookup"><span data-stu-id="79d43-141">1.6</span></span> |
|[<span data-ttu-id="79d43-142">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="79d43-142">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79d43-143">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79d43-143">ReadItem</span></span>|
|[<span data-ttu-id="79d43-144">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="79d43-144">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79d43-145">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="79d43-145">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="79d43-146">Пример</span><span class="sxs-lookup"><span data-stu-id="79d43-146">Example</span></span>

```js
console.log(Office.context.mailbox.userProfile.accountType);
```

####  <a name="displayname-string"></a><span data-ttu-id="79d43-147">displayName :String</span><span class="sxs-lookup"><span data-stu-id="79d43-147">displayName :String</span></span>

<span data-ttu-id="79d43-148">Получает отображаемое имя пользователя.</span><span class="sxs-lookup"><span data-stu-id="79d43-148">Gets the user's display name.</span></span>

##### <a name="type"></a><span data-ttu-id="79d43-149">Тип:</span><span class="sxs-lookup"><span data-stu-id="79d43-149">Type:</span></span>

*   <span data-ttu-id="79d43-150">String</span><span class="sxs-lookup"><span data-stu-id="79d43-150">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79d43-151">Требования</span><span class="sxs-lookup"><span data-stu-id="79d43-151">Requirements</span></span>

|<span data-ttu-id="79d43-152">Требование</span><span class="sxs-lookup"><span data-stu-id="79d43-152">Requirement</span></span>| <span data-ttu-id="79d43-153">Значение</span><span class="sxs-lookup"><span data-stu-id="79d43-153">Value</span></span>|
|---|---|
|[<span data-ttu-id="79d43-154">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="79d43-154">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79d43-155">1.0</span><span class="sxs-lookup"><span data-stu-id="79d43-155">1.0</span></span>|
|[<span data-ttu-id="79d43-156">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="79d43-156">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79d43-157">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79d43-157">ReadItem</span></span>|
|[<span data-ttu-id="79d43-158">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="79d43-158">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79d43-159">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="79d43-159">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="79d43-160">Пример</span><span class="sxs-lookup"><span data-stu-id="79d43-160">Example</span></span>

```js
// Example: Allie Bellew
console.log(Office.context.mailbox.userProfile.displayName);
```

####  <a name="emailaddress-string"></a><span data-ttu-id="79d43-161">emailAddress :String</span><span class="sxs-lookup"><span data-stu-id="79d43-161">emailAddress :String</span></span>

<span data-ttu-id="79d43-162">Получает адрес электронной почты SMTP пользователя.</span><span class="sxs-lookup"><span data-stu-id="79d43-162">Gets the user's SMTP email address.</span></span>

##### <a name="type"></a><span data-ttu-id="79d43-163">Тип:</span><span class="sxs-lookup"><span data-stu-id="79d43-163">Type:</span></span>

*   <span data-ttu-id="79d43-164">String</span><span class="sxs-lookup"><span data-stu-id="79d43-164">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79d43-165">Требования</span><span class="sxs-lookup"><span data-stu-id="79d43-165">Requirements</span></span>

|<span data-ttu-id="79d43-166">Требование</span><span class="sxs-lookup"><span data-stu-id="79d43-166">Requirement</span></span>| <span data-ttu-id="79d43-167">Значение</span><span class="sxs-lookup"><span data-stu-id="79d43-167">Value</span></span>|
|---|---|
|[<span data-ttu-id="79d43-168">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="79d43-168">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79d43-169">1.0</span><span class="sxs-lookup"><span data-stu-id="79d43-169">1.0</span></span>|
|[<span data-ttu-id="79d43-170">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="79d43-170">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79d43-171">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79d43-171">ReadItem</span></span>|
|[<span data-ttu-id="79d43-172">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="79d43-172">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79d43-173">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="79d43-173">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="79d43-174">Пример</span><span class="sxs-lookup"><span data-stu-id="79d43-174">Example</span></span>

```js
// Example: allieb@contoso.com
console.log(Office.context.mailbox.userProfile.emailAddress);
```

####  <a name="timezone-string"></a><span data-ttu-id="79d43-175">timeZone :String</span><span class="sxs-lookup"><span data-stu-id="79d43-175">timeZone :String</span></span>

<span data-ttu-id="79d43-176">Получает часовой пояс пользователя по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="79d43-176">Gets the user's default time zone.</span></span>

##### <a name="type"></a><span data-ttu-id="79d43-177">Тип:</span><span class="sxs-lookup"><span data-stu-id="79d43-177">Type:</span></span>

*   <span data-ttu-id="79d43-178">String</span><span class="sxs-lookup"><span data-stu-id="79d43-178">String</span></span>

##### <a name="requirements"></a><span data-ttu-id="79d43-179">Требования</span><span class="sxs-lookup"><span data-stu-id="79d43-179">Requirements</span></span>

|<span data-ttu-id="79d43-180">Требование</span><span class="sxs-lookup"><span data-stu-id="79d43-180">Requirement</span></span>| <span data-ttu-id="79d43-181">Значение</span><span class="sxs-lookup"><span data-stu-id="79d43-181">Value</span></span>|
|---|---|
|[<span data-ttu-id="79d43-182">Версия минимального набора требований к почтовому ящику</span><span class="sxs-lookup"><span data-stu-id="79d43-182">Minimum mailbox requirement set version</span></span>](/office/dev/add-ins/reference/requirement-sets/outlook-api-requirement-sets)| <span data-ttu-id="79d43-183">1.0</span><span class="sxs-lookup"><span data-stu-id="79d43-183">1.0</span></span>|
|[<span data-ttu-id="79d43-184">Минимальный уровень разрешений</span><span class="sxs-lookup"><span data-stu-id="79d43-184">Minimum permission level</span></span>](https://docs.microsoft.com/outlook/add-ins/understanding-outlook-add-in-permissions)| <span data-ttu-id="79d43-185">ReadItem</span><span class="sxs-lookup"><span data-stu-id="79d43-185">ReadItem</span></span>|
|[<span data-ttu-id="79d43-186">Применимый режим Outlook</span><span class="sxs-lookup"><span data-stu-id="79d43-186">Applicable Outlook mode</span></span>](https://docs.microsoft.com/outlook/add-ins/#extension-points)| <span data-ttu-id="79d43-187">Создание или чтение</span><span class="sxs-lookup"><span data-stu-id="79d43-187">Compose or read</span></span>|

##### <a name="example"></a><span data-ttu-id="79d43-188">Пример</span><span class="sxs-lookup"><span data-stu-id="79d43-188">Example</span></span>

```js
// Example: Pacific Standard Time
console.log(Office.context.mailbox.userProfile.timeZone);
```