---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: Функции и API, которые в настоящее время находятся в предварительной версии для надстройки Outlook.
ms.date: 02/05/2021
localization_priority: Normal
ms.openlocfilehash: 92ba3510af0c8b9ebdf9ca4368c889b821a9cb3b
ms.sourcegitcommit: 4805454f7fc6c64368a35d014e24075faf3e7557
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/10/2021
ms.locfileid: "50173957"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="6d755-103">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="6d755-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="6d755-104">Подмножество API надстройки Outlook aPI JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройки Outlook.</span><span class="sxs-lookup"><span data-stu-id="6d755-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="6d755-105">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="6d755-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="6d755-106">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="6d755-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="6d755-107">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="6d755-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="6d755-108">Вы можете просмотреть функции в Outlook в Интернете, настроив целевой выпуск в [клиенте Microsoft 365.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="6d755-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="6d755-109">На этой странице отмечена "Настройка доступа к предварительному просмотру" для применимых функций.</span><span class="sxs-lookup"><span data-stu-id="6d755-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="6d755-110">Для других функций вы можете запросить доступ к битам предварительного просмотра для Outlook в Интернете с помощью учетной записи Microsoft 365, заполнив и передав [эту форму.](https://aka.ms/OWAPreview)</span><span class="sxs-lookup"><span data-stu-id="6d755-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="6d755-111">Для этих функций отмечено "Запрашивать предварительный доступ".</span><span class="sxs-lookup"><span data-stu-id="6d755-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="6d755-112">Набор предварительных требований включает все функции набора требований [1.9.](../requirement-set-1.9/outlook-requirement-set-1.9.md)</span><span class="sxs-lookup"><span data-stu-id="6d755-112">The Preview Requirement set includes all of the features of [Requirement set 1.9](../requirement-set-1.9/outlook-requirement-set-1.9.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="6d755-113">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="6d755-113">Features in preview</span></span>

<span data-ttu-id="6d755-114">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="6d755-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="6d755-115">Активация надстройки для элементов, защищенных с помощью управления правами на управление правами на данные (IRM)</span><span class="sxs-lookup"><span data-stu-id="6d755-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="6d755-116">Теперь надстройки могут активироваться для элементов, защищенных с защитой IRM.</span><span class="sxs-lookup"><span data-stu-id="6d755-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="6d755-117">Чтобы включить эту возможность, администратор клиента должен включить право на использование, задав параметр политики "Разрешить программный доступ" `OBJMODEL` в Office. </span><span class="sxs-lookup"><span data-stu-id="6d755-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="6d755-118">Дополнительные [сведения см. в описании](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) и правах на использование.</span><span class="sxs-lookup"><span data-stu-id="6d755-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="6d755-119">**Доступно в**: Outlook для Windows, начиная со сборки 13229.10000 (подключенной к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6d755-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="6d755-120">Дополнительные свойства календаря</span><span class="sxs-lookup"><span data-stu-id="6d755-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="6d755-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="6d755-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="6d755-122">Добавлен новый объект, который представляет свойство события на весь день встречи в режиме compose.</span><span class="sxs-lookup"><span data-stu-id="6d755-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="6d755-123">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6d755-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="6d755-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="6d755-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="6d755-125">Добавлен новый объект, который представляет чувствительность встречи в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="6d755-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="6d755-126">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6d755-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="6d755-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="6d755-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="6d755-128">Добавлено новое свойство, которое представляет, является ли встреча событием на весь день.</span><span class="sxs-lookup"><span data-stu-id="6d755-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="6d755-129">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6d755-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="6d755-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="6d755-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="6d755-131">Добавлено новое свойство, которое представляет чувствительность встречи.</span><span class="sxs-lookup"><span data-stu-id="6d755-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="6d755-132">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6d755-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="6d755-133">Office.MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="6d755-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="6d755-134">Добавлено новое `AppointmentSensitivityType` enum, которое представляет параметры конфиденциальности, доступные для встречи.</span><span class="sxs-lookup"><span data-stu-id="6d755-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="6d755-135">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6d755-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="6d755-136">Активация на основе событий</span><span class="sxs-lookup"><span data-stu-id="6d755-136">Event-based activation</span></span>

<span data-ttu-id="6d755-137">Добавлена поддержка функций активации на основе событий в надстройки Outlook. Подробнее [см. в](../../../outlook/autolaunch.md) подстройке "Настройка надстройки Outlook для активации на основе событий".</span><span class="sxs-lookup"><span data-stu-id="6d755-137">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="6d755-138">Точка расширения LaunchEvent</span><span class="sxs-lookup"><span data-stu-id="6d755-138">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="6d755-139">Добавлена `LaunchEvent` поддержка точки расширения для манифеста.</span><span class="sxs-lookup"><span data-stu-id="6d755-139">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="6d755-140">Он настраивает функции активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="6d755-140">It configures event-based activation functionality.</span></span>

<span data-ttu-id="6d755-141">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="6d755-141">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="6d755-142">Элемент манифеста LaunchEvents</span><span class="sxs-lookup"><span data-stu-id="6d755-142">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="6d755-143">Добавлен `LaunchEvents` элемент манифеста.</span><span class="sxs-lookup"><span data-stu-id="6d755-143">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="6d755-144">Он поддерживает настройку функций активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="6d755-144">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="6d755-145">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="6d755-145">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="6d755-146">Элемент манифеста runtimes</span><span class="sxs-lookup"><span data-stu-id="6d755-146">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="6d755-147">Добавлена поддержка Outlook для `Runtimes` элемента манифеста.</span><span class="sxs-lookup"><span data-stu-id="6d755-147">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="6d755-148">Он ссылается на файлы HTML и JavaScript, необходимые для активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="6d755-148">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="6d755-149">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="6d755-149">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="6d755-150">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="6d755-150">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="6d755-151">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="6d755-151">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="6d755-152">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="6d755-152">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="6d755-153">**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="6d755-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="6d755-154">Подпись почты</span><span class="sxs-lookup"><span data-stu-id="6d755-154">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="6d755-155">Office.context.mailbox.item.body.setSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="6d755-155">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="6d755-156">Добавлена новая функция для объекта, которая добавляет или заменяет подпись в теле `Body` элемента в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="6d755-156">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="6d755-157">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="6d755-157">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="6d755-158">Office.context.mailbox.item.disableClientSignatureAsync</span><span class="sxs-lookup"><span data-stu-id="6d755-158">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="6d755-159">Добавлена новая функция, которая отключает подпись клиента для отправляемого почтового ящика в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="6d755-159">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="6d755-160">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="6d755-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="6d755-161">Office.context.mailbox.item.getComposeTypeAsync</span><span class="sxs-lookup"><span data-stu-id="6d755-161">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="6d755-162">Добавлена новая функция, которая получает тип составить сообщение в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="6d755-162">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="6d755-163">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="6d755-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="6d755-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span><span class="sxs-lookup"><span data-stu-id="6d755-164">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="6d755-165">Добавлена новая функция, которая проверяет, включена ли подпись клиента для элемента в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="6d755-165">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="6d755-166">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="6d755-166">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="6d755-167">Office.MailboxEnums.ComposeType</span><span class="sxs-lookup"><span data-stu-id="6d755-167">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="6d755-168">Добавлено новое `ComposeType` enum, доступное в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="6d755-168">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="6d755-169">**Доступно в**: Outlook для Windows (подключенный к подписке Microsoft 365), Outlook в Интернете (современный, [настройка доступа к предварительной версии)](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="6d755-169">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="6d755-170">Уведомления с действиями</span><span class="sxs-lookup"><span data-stu-id="6d755-170">Notification messages with actions</span></span>

<span data-ttu-id="6d755-171">Эта функция позволяет надстройка включать уведомление с дополнительным действием, кроме действия по умолчанию **"Отклонять".**</span><span class="sxs-lookup"><span data-stu-id="6d755-171">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span> <span data-ttu-id="6d755-172">В современном Outlook в Интернете эта функция доступна только в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="6d755-172">In modern Outlook on the web, this feature is available in Compose mode only.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="6d755-173">Office.NotificationMessageDetails.actions</span><span class="sxs-lookup"><span data-stu-id="6d755-173">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="6d755-174">Добавлено новое свойство, которое позволяет добавить уведомление `InsightMessage` с помощью дополнительного действия.</span><span class="sxs-lookup"><span data-stu-id="6d755-174">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="6d755-175">**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="6d755-175">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="6d755-176">Office.NotificationMessageAction</span><span class="sxs-lookup"><span data-stu-id="6d755-176">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="6d755-177">Добавлен новый объект, в котором вы определяете дополнительное действие для `InsightMessage` уведомления.</span><span class="sxs-lookup"><span data-stu-id="6d755-177">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="6d755-178">**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="6d755-178">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="6d755-179">Office.MailboxEnums.ActionType</span><span class="sxs-lookup"><span data-stu-id="6d755-179">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="6d755-180">Добавлено новое enum `ActionType` .</span><span class="sxs-lookup"><span data-stu-id="6d755-180">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="6d755-181">**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="6d755-181">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="6d755-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span><span class="sxs-lookup"><span data-stu-id="6d755-182">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="6d755-183">Добавлен новый тип `InsightMessage` в `ItemNotificationMessageType` enum.</span><span class="sxs-lookup"><span data-stu-id="6d755-183">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="6d755-184">**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="6d755-184">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="6d755-185">Тема Office</span><span class="sxs-lookup"><span data-stu-id="6d755-185">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="6d755-186">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="6d755-186">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="6d755-187">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="6d755-187">Added ability to get Office theme.</span></span>

<span data-ttu-id="6d755-188">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6d755-188">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="6d755-189">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="6d755-189">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="6d755-190">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="6d755-190">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="6d755-191">**Доступно в**: Outlook для Windows (подключен к подписке На Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="6d755-191">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="6d755-192">Данные сеансов</span><span class="sxs-lookup"><span data-stu-id="6d755-192">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="6d755-193">Office.SessionData</span><span class="sxs-lookup"><span data-stu-id="6d755-193">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="6d755-194">Добавлен новый объект, который представляет данные сеанса элемента.</span><span class="sxs-lookup"><span data-stu-id="6d755-194">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="6d755-195">**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="6d755-195">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="6d755-196">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="6d755-196">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="6d755-197">Добавлено новое свойство для управления данными сеанса элемента в режиме составить.</span><span class="sxs-lookup"><span data-stu-id="6d755-197">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="6d755-198">**Доступно в**: Outlook для Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная версия)</span><span class="sxs-lookup"><span data-stu-id="6d755-198">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="6d755-199">См. также</span><span class="sxs-lookup"><span data-stu-id="6d755-199">See also</span></span>

- [<span data-ttu-id="6d755-200">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="6d755-200">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="6d755-201">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="6d755-201">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="6d755-202">Начало работы</span><span class="sxs-lookup"><span data-stu-id="6d755-202">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="6d755-203">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="6d755-203">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
