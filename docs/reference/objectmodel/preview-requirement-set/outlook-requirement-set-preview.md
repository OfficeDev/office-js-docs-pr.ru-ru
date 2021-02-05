---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: Функции и API, которые в настоящее время находятся в предварительной версии для надстройки Outlook.
ms.date: 02/02/2021
localization_priority: Normal
ms.openlocfilehash: 39dd1221f4dea9674c89cdaad20024ce408f8db3
ms.sourcegitcommit: 8546889a759590c3798ce56e311d9e46f0171413
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/04/2021
ms.locfileid: "50104842"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="3409d-103">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3409d-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="3409d-104">Подмножество API надстройки Outlook aPI JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="3409d-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3409d-105">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3409d-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="3409d-106">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="3409d-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="3409d-107">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="3409d-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="3409d-108">Вы можете просмотреть функции в Outlook в Интернете, настроив целевой выпуск в [клиенте Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="3409d-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="3409d-109">На этой странице отмечена "Настройка доступа к предварительному просмотру" для применимых функций.</span><span class="sxs-lookup"><span data-stu-id="3409d-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="3409d-110">Для других функций вы можете запросить доступ к битам предварительного просмотра для Outlook в Интернете с помощью учетной записи Microsoft 365, заполнив и передав [эту форму.](https://aka.ms/OWAPreview)</span><span class="sxs-lookup"><span data-stu-id="3409d-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="3409d-111">Для этих функций отмечена "Запрос предварительного доступа".</span><span class="sxs-lookup"><span data-stu-id="3409d-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="3409d-112">Набор предварительных требований включает все функции набора требований [1.9.](../requirement-set-1.9/outlook-requirement-set-1.9.md)</span><span class="sxs-lookup"><span data-stu-id="3409d-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="3409d-113">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="3409d-113">Features in preview</span></span>

<span data-ttu-id="3409d-114">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="3409d-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="3409d-115">Активация надстройки для элементов, защищенных с помощью управления правами на данные (IRM)</span><span class="sxs-lookup"><span data-stu-id="3409d-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="3409d-116">Надстройки теперь могут активироваться для элементов, защищенных СУИБ.</span><span class="sxs-lookup"><span data-stu-id="3409d-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="3409d-117">Чтобы включить эту возможность, администратор клиента должен включить право на использование, задав параметр политики "Разрешить программный доступ" `OBJMODEL` в Office. </span><span class="sxs-lookup"><span data-stu-id="3409d-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="3409d-118">Дополнительные [сведения см. в](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) описании и правах на использование.</span><span class="sxs-lookup"><span data-stu-id="3409d-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="3409d-119">**Доступно в**: Outlook для Windows, начиная со сборки 13229.10000 (подключенной к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3409d-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="3409d-120">Дополнительные свойства календаря</span><span class="sxs-lookup"><span data-stu-id="3409d-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="3409d-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="3409d-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="3409d-122">Добавлен новый объект, который представляет свойство события на весь день встречи в режиме compose.</span><span class="sxs-lookup"><span data-stu-id="3409d-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="3409d-123">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3409d-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="3409d-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="3409d-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="3409d-125">Добавлен новый объект, который представляет чувствительность встречи в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="3409d-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="3409d-126">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3409d-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="3409d-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="3409d-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="3409d-128">Добавлено новое свойство, которое представляет, является ли встреча событием на весь день.</span><span class="sxs-lookup"><span data-stu-id="3409d-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="3409d-129">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3409d-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="3409d-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="3409d-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="3409d-131">Добавлено новое свойство, которое представляет чувствительность встречи.</span><span class="sxs-lookup"><span data-stu-id="3409d-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="3409d-132">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3409d-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="3409d-133">Office.MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="3409d-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="3409d-134">Добавлено новое `AppointmentSensitivityType` enum, которое представляет параметры конфиденциальности, доступные для встречи.</span><span class="sxs-lookup"><span data-stu-id="3409d-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="3409d-135">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3409d-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="3409d-136">Активация на основе событий</span><span class="sxs-lookup"><span data-stu-id="3409d-136">Event-based activation</span></span>

<span data-ttu-id="3409d-137">Добавлена поддержка функций активации на основе событий в надстройки Outlook. Подробнее [см. в](../../../outlook/autolaunch.md) подстройке "Настройка надстройки Outlook для активации на основе событий".</span><span class="sxs-lookup"><span data-stu-id="3409d-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="3409d-138">Точка расширения LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="3409d-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="3409d-139">Добавлена `LaunchEvent` поддержка точек расширения для манифеста.</span><span class="sxs-lookup"><span data-stu-id="3409d-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="3409d-140">Он настраивает функции активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="3409d-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="3409d-141">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="3409d-141">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="3409d-142">Элемент манифеста LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="3409d-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="3409d-143">Добавлен `LaunchEvents` элемент манифеста.</span><span class="sxs-lookup"><span data-stu-id="3409d-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="3409d-144">Он поддерживает настройку функций активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="3409d-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="3409d-145">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="3409d-145">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="3409d-146">Элемент манифеста runtimes</span><span class="sxs-lookup"><span data-stu-id="3409d-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="3409d-147">Добавлена поддержка Outlook для `Runtimes` элемента манифеста.</span><span class="sxs-lookup"><span data-stu-id="3409d-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="3409d-148">Он ссылается на файлы HTML и JavaScript, необходимые для активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="3409d-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="3409d-149">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="3409d-149">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="3409d-150">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="3409d-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="3409d-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="3409d-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="3409d-152">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="3409d-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="3409d-153">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="3409d-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="3409d-154">Подпись почты</span><span class="sxs-lookup"><span data-stu-id="3409d-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="3409d-155">Office.context.mailbox.item.body.setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="3409d-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="3409d-156">Добавлена новая функция для объекта, которая добавляет или заменяет подпись в теле `Body` элемента в режиме compose.</span><span class="sxs-lookup"><span data-stu-id="3409d-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="3409d-157">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="3409d-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="3409d-158">Office.context.mailbox.item.disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="3409d-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="3409d-159">Добавлена новая функция, которая отключает подпись клиента для отправляемого почтового ящика в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="3409d-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="3409d-160">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="3409d-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="3409d-161">Office.context.mailbox.item.getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="3409d-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="3409d-162">Добавлена новая функция, которая получает тип составить сообщение в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="3409d-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="3409d-163">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="3409d-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="3409d-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="3409d-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="3409d-165">Добавлена новая функция, которая проверяет, включена ли подпись клиента для элемента в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="3409d-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="3409d-166">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="3409d-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="3409d-167">Office.MailboxEnums.ComposeType</span><span class="sxs-lookup"><span data-stu-id="3409d-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="3409d-168">Добавлено новое `ComposeType` enum, доступное в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="3409d-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="3409d-169">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="3409d-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="3409d-170">Уведомления с действиями</span><span class="sxs-lookup"><span data-stu-id="3409d-170">Notification messages with actions</span></span>

<span data-ttu-id="3409d-171">С помощью этой функции надстройка может добавить уведомление с  дополнительным действием, кроме действия "Отклонять" по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="3409d-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span> <span data-ttu-id="3409d-172">В современном Outlook в Интернете эта функция доступна только в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="3409d-172">In modern Outlook on the web, this feature is available in Compose mode only.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="3409d-173">Office.NotificationMessageDetails.actions</span><span class="sxs-lookup"><span data-stu-id="3409d-173">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="3409d-174">Добавлено новое свойство, которое позволяет добавить уведомление `InsightMessage` с помощью дополнительного действия.</span><span class="sxs-lookup"><span data-stu-id="3409d-174">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="3409d-175">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="3409d-175">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="3409d-176">Office.NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="3409d-176">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="3409d-177">Добавлен новый объект, в котором вы определяете дополнительное действие для `InsightMessage` уведомления.</span><span class="sxs-lookup"><span data-stu-id="3409d-177">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="3409d-178">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="3409d-178">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="3409d-179">Office.MailboxEnums.ActionType</span><span class="sxs-lookup"><span data-stu-id="3409d-179">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="3409d-180">Добавлено новое enum `ActionType` .</span><span class="sxs-lookup"><span data-stu-id="3409d-180">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="3409d-181">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="3409d-181">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="3409d-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span><span class="sxs-lookup"><span data-stu-id="3409d-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="3409d-183">Добавлен новый тип `InsightMessage` в `ItemNotificationMessageType` enum.</span><span class="sxs-lookup"><span data-stu-id="3409d-183">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="3409d-184">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="3409d-184">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="3409d-185">Тема Office</span><span class="sxs-lookup"><span data-stu-id="3409d-185">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="3409d-186">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="3409d-186">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="3409d-187">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="3409d-187">Added ability to get Office theme.</span></span>

<span data-ttu-id="3409d-188">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3409d-188">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="3409d-189">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="3409d-189">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="3409d-190">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="3409d-190">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="3409d-191">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3409d-191">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="3409d-192">Данные сеансов</span><span class="sxs-lookup"><span data-stu-id="3409d-192">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="3409d-193">Office.SessionData</span><span class="sxs-lookup"><span data-stu-id="3409d-193">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="3409d-194">Добавлен новый объект, который представляет данные сеанса элемента.</span><span class="sxs-lookup"><span data-stu-id="3409d-194">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="3409d-195">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3409d-195">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="3409d-196">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="3409d-196">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="3409d-197">Добавлено новое свойство для управления данными сеанса элемента в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="3409d-197">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="3409d-198">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="3409d-198">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

## <a name="see-also"></a><span data-ttu-id="3409d-199">См. также</span><span class="sxs-lookup"><span data-stu-id="3409d-199">See also</span></span>

- [<span data-ttu-id="3409d-200">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3409d-200">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="3409d-201">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3409d-201">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="3409d-202">Начало работы</span><span class="sxs-lookup"><span data-stu-id="3409d-202">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="3409d-203">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="3409d-203">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
