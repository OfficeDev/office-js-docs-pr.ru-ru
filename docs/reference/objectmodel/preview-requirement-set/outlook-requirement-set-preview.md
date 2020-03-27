---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: Функции и API, которые в настоящее время находятся в предварительной версии для надстроек Outlook и API JavaScript для Office.
ms.date: 03/26/2020
localization_priority: Normal
ms.openlocfilehash: 55de284932a53d2226258a15c86ead4f05361c30
ms.sourcegitcommit: 315a648cce38609c3e1c92bd4a339e268f8a2e1d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/26/2020
ms.locfileid: "42978622"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="2fa42-103">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="2fa42-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="2fa42-104">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="2fa42-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="2fa42-105">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="2fa42-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="2fa42-106">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="2fa42-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="2fa42-107">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="2fa42-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="2fa42-108">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="2fa42-108">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="2fa42-109">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="2fa42-109">Features in preview</span></span>

<span data-ttu-id="2fa42-110">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="2fa42-110">The following features are in preview.</span></span>

### <a name="append-on-send"></a><span data-ttu-id="2fa42-111">Добавление при отправке</span><span class="sxs-lookup"><span data-stu-id="2fa42-111">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="2fa42-112">Office. Context. Mailbox. Item. Body. Аппендонсендасинк</span><span class="sxs-lookup"><span data-stu-id="2fa42-112">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="2fa42-113">Добавлена новая функция для `Body` объекта, который добавляет данные в конец тела элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="2fa42-113">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="2fa42-114">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="2fa42-114">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="2fa42-115">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="2fa42-115">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="2fa42-116">Добавлен новый элемент в манифест, где `AppendOnSend` расширенное разрешение должно быть включено в коллекцию расширенных разрешений.</span><span class="sxs-lookup"><span data-stu-id="2fa42-116">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="2fa42-117">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="2fa42-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="2fa42-118">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="2fa42-118">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="2fa42-119">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="2fa42-119">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="2fa42-120">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="2fa42-120">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="2fa42-121">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="2fa42-121">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="2fa42-122">Подпись почты</span><span class="sxs-lookup"><span data-stu-id="2fa42-122">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="2fa42-123">Office. Context. Mailbox. Item. Body. Сетсигнатуреасинк</span><span class="sxs-lookup"><span data-stu-id="2fa42-123">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="2fa42-124">Добавлена новая функция для `Body` объекта, который добавляет или заменяет подпись в теле элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="2fa42-124">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="2fa42-125">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="2fa42-125">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="2fa42-126">Office. Context. Mailbox. Item. Дисаблеклиентсигнатуреасинк</span><span class="sxs-lookup"><span data-stu-id="2fa42-126">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="2fa42-127">Добавлена новая функция, которая отключает подпись клиента для отправляющего почтового ящика в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="2fa42-127">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="2fa42-128">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="2fa42-128">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="2fa42-129">Office. Context. Mailbox. Item. Жеткомпосетипеасинк</span><span class="sxs-lookup"><span data-stu-id="2fa42-129">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="2fa42-130">Добавлена новая функция, которая получает тип сообщения "создание" в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="2fa42-130">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="2fa42-131">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="2fa42-131">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="2fa42-132">Office. Context. Mailbox. Item. Исклиентсигнатуринабледасинк</span><span class="sxs-lookup"><span data-stu-id="2fa42-132">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="2fa42-133">Добавлена новая функция, проверяющая, включена ли подпись клиента для элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="2fa42-133">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="2fa42-134">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="2fa42-134">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="2fa42-135">Office. MailboxEnums. Компосетипе</span><span class="sxs-lookup"><span data-stu-id="2fa42-135">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="2fa42-136">Добавлено новое перечисление `ComposeType` , доступное в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="2fa42-136">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="2fa42-137">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="2fa42-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="2fa42-138">Тема Office</span><span class="sxs-lookup"><span data-stu-id="2fa42-138">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="2fa42-139">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="2fa42-139">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="2fa42-140">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="2fa42-140">Added ability to get Office theme.</span></span>

<span data-ttu-id="2fa42-141">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="2fa42-141">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="2fa42-142">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="2fa42-142">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="2fa42-143">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="2fa42-143">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="2fa42-144">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="2fa42-144">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="sso"></a><span data-ttu-id="2fa42-145">Единый вход</span><span class="sxs-lookup"><span data-stu-id="2fa42-145">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="2fa42-146">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="2fa42-146">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="2fa42-147">Добавлена возможность доступа к `getAccessToken`, что позволяет надстройкам [получать маркер доступа](../../../outlook/authenticate-a-user-with-an-sso-token.md) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="2fa42-147">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="2fa42-148">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="2fa42-148">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="2fa42-149">См. также</span><span class="sxs-lookup"><span data-stu-id="2fa42-149">See also</span></span>

- [<span data-ttu-id="2fa42-150">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="2fa42-150">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="2fa42-151">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="2fa42-151">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="2fa42-152">Начало работы</span><span class="sxs-lookup"><span data-stu-id="2fa42-152">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="2fa42-153">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="2fa42-153">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
