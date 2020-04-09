---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: Функции и API, которые в настоящее время находятся в предварительной версии для надстроек Outlook и API JavaScript для Office.
ms.date: 04/08/2020
localization_priority: Normal
ms.openlocfilehash: acc19c81f929596b0bd5622e696c1988cf31ee5c
ms.sourcegitcommit: c3bfea0818af1f01e71a1feff707fb2456a69488
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/08/2020
ms.locfileid: "43185416"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="d1e97-103">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d1e97-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="d1e97-104">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="d1e97-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="d1e97-105">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="d1e97-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="d1e97-106">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="d1e97-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="d1e97-107">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="d1e97-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

<span data-ttu-id="d1e97-108">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="d1e97-108">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="d1e97-109">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="d1e97-109">Features in preview</span></span>

<span data-ttu-id="d1e97-110">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="d1e97-110">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="d1e97-111">Дополнительные свойства календаря</span><span class="sxs-lookup"><span data-stu-id="d1e97-111">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="d1e97-112">исаллдайевент</span><span class="sxs-lookup"><span data-stu-id="d1e97-112">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="d1e97-113">Добавлен новый объект, представляющий свойство события "целый день" для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d1e97-113">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="d1e97-114">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-114">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="d1e97-115">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="d1e97-115">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="d1e97-116">Добавлен новый объект, представляющий чувствительность встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d1e97-116">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="d1e97-117">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-117">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="d1e97-118">Office. Context. Mailbox. Item. Исаллдайевент</span><span class="sxs-lookup"><span data-stu-id="d1e97-118">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="d1e97-119">Добавлено новое свойство, которое указывает, является ли встреча событием на целый день.</span><span class="sxs-lookup"><span data-stu-id="d1e97-119">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="d1e97-120">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-120">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="d1e97-121">Office. Context. Mailbox. Item. чувствительность</span><span class="sxs-lookup"><span data-stu-id="d1e97-121">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="d1e97-122">Добавлено новое свойство, представляющее чувствительность встречи.</span><span class="sxs-lookup"><span data-stu-id="d1e97-122">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="d1e97-123">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-123">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="d1e97-124">Office. MailboxEnums. Аппоинтментсенситивититипе</span><span class="sxs-lookup"><span data-stu-id="d1e97-124">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="d1e97-125">Добавлено новое перечисление `AppointmentSensitivityType` , представляющее параметры конфиденциальности, доступные для встречи.</span><span class="sxs-lookup"><span data-stu-id="d1e97-125">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="d1e97-126">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-126">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="d1e97-127">Добавление при отправке</span><span class="sxs-lookup"><span data-stu-id="d1e97-127">Append on send</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="d1e97-128">Office. Context. Mailbox. Item. Body. Аппендонсендасинк</span><span class="sxs-lookup"><span data-stu-id="d1e97-128">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="d1e97-129">Добавлена новая функция для `Body` объекта, который добавляет данные в конец тела элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d1e97-129">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="d1e97-130">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="d1e97-130">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="d1e97-131">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="d1e97-131">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="d1e97-132">Добавлен новый элемент в манифест, где `AppendOnSend` расширенное разрешение должно быть включено в коллекцию расширенных разрешений.</span><span class="sxs-lookup"><span data-stu-id="d1e97-132">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="d1e97-133">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="d1e97-133">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="d1e97-134">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="d1e97-134">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="d1e97-135">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="d1e97-135">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="d1e97-136">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="d1e97-136">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="d1e97-137">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="d1e97-137">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="d1e97-138">Подпись почты</span><span class="sxs-lookup"><span data-stu-id="d1e97-138">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="d1e97-139">Office. Context. Mailbox. Item. Body. Сетсигнатуреасинк</span><span class="sxs-lookup"><span data-stu-id="d1e97-139">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="d1e97-140">Добавлена новая функция для `Body` объекта, который добавляет или заменяет подпись в теле элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d1e97-140">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="d1e97-141">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-141">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="d1e97-142">Office. Context. Mailbox. Item. Дисаблеклиентсигнатуреасинк</span><span class="sxs-lookup"><span data-stu-id="d1e97-142">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="d1e97-143">Добавлена новая функция, которая отключает подпись клиента для отправляющего почтового ящика в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d1e97-143">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="d1e97-144">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-144">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="d1e97-145">Office. Context. Mailbox. Item. Жеткомпосетипеасинк</span><span class="sxs-lookup"><span data-stu-id="d1e97-145">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="d1e97-146">Добавлена новая функция, которая получает тип сообщения "создание" в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d1e97-146">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="d1e97-147">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-147">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="d1e97-148">Office. Context. Mailbox. Item. Исклиентсигнатуринабледасинк</span><span class="sxs-lookup"><span data-stu-id="d1e97-148">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="d1e97-149">Добавлена новая функция, проверяющая, включена ли подпись клиента для элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d1e97-149">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="d1e97-150">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-150">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="d1e97-151">Office. MailboxEnums. Компосетипе</span><span class="sxs-lookup"><span data-stu-id="d1e97-151">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="d1e97-152">Добавлено новое перечисление `ComposeType` , доступное в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="d1e97-152">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="d1e97-153">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-153">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="d1e97-154">Тема Office</span><span class="sxs-lookup"><span data-stu-id="d1e97-154">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="d1e97-155">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="d1e97-155">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="d1e97-156">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="d1e97-156">Added ability to get Office theme.</span></span>

<span data-ttu-id="d1e97-157">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-157">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="d1e97-158">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="d1e97-158">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="d1e97-159">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="d1e97-159">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="d1e97-160">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-160">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="online-meeting-provider-integration"></a><span data-ttu-id="d1e97-161">Интеграция поставщика собраний по сети</span><span class="sxs-lookup"><span data-stu-id="d1e97-161">Online meeting provider integration</span></span>

#### <a name="mobileonlinemeetingcommandsurface-extension-point"></a>[<span data-ttu-id="d1e97-162">Точка расширения Мобилеонлинемитингкоммандсурфаце</span><span class="sxs-lookup"><span data-stu-id="d1e97-162">MobileOnlineMeetingCommandSurface extension point</span></span>](../../manifest/extensionpoint.md#mobileonlinemeetingcommandsurface-preview)

<span data-ttu-id="d1e97-163">Добавлена `MobileOnlineMeetingCommandSurface` точка расширения для манифеста.</span><span class="sxs-lookup"><span data-stu-id="d1e97-163">Added `MobileOnlineMeetingCommandSurface` extension point to manifest.</span></span> <span data-ttu-id="d1e97-164">Он определяет интеграцию собраний по сети.</span><span class="sxs-lookup"><span data-stu-id="d1e97-164">It defines the online meeting integration.</span></span>

<span data-ttu-id="d1e97-165">**Доступно в**: Outlook на Android (подключено к подписке Office 365)</span><span class="sxs-lookup"><span data-stu-id="d1e97-165">**Available in**: Outlook on Android (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="sso"></a><span data-ttu-id="d1e97-166">Единый вход</span><span class="sxs-lookup"><span data-stu-id="d1e97-166">SSO</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="d1e97-167">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="d1e97-167">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="d1e97-168">Добавлена возможность доступа к `getAccessToken`, что позволяет надстройкам [получать маркер доступа](../../../outlook/authenticate-a-user-with-an-sso-token.md) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="d1e97-168">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="d1e97-169">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="d1e97-169">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="d1e97-170">См. также</span><span class="sxs-lookup"><span data-stu-id="d1e97-170">See also</span></span>

- [<span data-ttu-id="d1e97-171">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d1e97-171">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="d1e97-172">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="d1e97-172">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="d1e97-173">Начало работы</span><span class="sxs-lookup"><span data-stu-id="d1e97-173">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="d1e97-174">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="d1e97-174">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
