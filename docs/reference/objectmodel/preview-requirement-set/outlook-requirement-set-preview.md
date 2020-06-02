---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: Функции и API, которые в настоящее время находятся в режиме предварительной версии для надстроек Outlook.
ms.date: 05/29/2020
localization_priority: Normal
ms.openlocfilehash: 600aad32c394d35e62f4024808b185e8a9abe5e8
ms.sourcegitcommit: 09a8683ff29cf06d0d1d822be83cf0798f1ccdf9
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/01/2020
ms.locfileid: "44471347"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="3b0fe-103">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3b0fe-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="3b0fe-104">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="3b0fe-105">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="3b0fe-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="3b0fe-106">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="3b0fe-107">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="3b0fe-108">Вы можете предварительно просмотреть функции в Outlook в Интернете, [настроив целевой выпуск на клиенте Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="3b0fe-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="3b0fe-109">"Настройка предварительного доступа" отмечено на этой странице в соответствующих возможностях.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="3b0fe-110">Для других функций вы можете запросить доступ к предварительной версии BITS для Outlook в Интернете, используя свою учетную запись Microsoft 365, заполнив и отправив [эту форму](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="3b0fe-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="3b0fe-111">В этих функциях указано "запросить доступ к предварительному доступу".</span><span class="sxs-lookup"><span data-stu-id="3b0fe-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="3b0fe-112">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="3b0fe-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="3b0fe-113">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="3b0fe-113">Features in preview</span></span>

<span data-ttu-id="3b0fe-114">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="3b0fe-115">Дополнительные свойства календаря</span><span class="sxs-lookup"><span data-stu-id="3b0fe-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="3b0fe-116">исаллдайевент</span><span class="sxs-lookup"><span data-stu-id="3b0fe-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="3b0fe-117">Добавлен новый объект, представляющий свойство события "целый день" для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="3b0fe-118">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3b0fe-118">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="3b0fe-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="3b0fe-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="3b0fe-120">Добавлен новый объект, представляющий чувствительность встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="3b0fe-121">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3b0fe-121">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="3b0fe-122">Office. Context. Mailbox. Item. Исаллдайевент</span><span class="sxs-lookup"><span data-stu-id="3b0fe-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="3b0fe-123">Добавлено новое свойство, которое указывает, является ли встреча событием на целый день.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="3b0fe-124">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3b0fe-124">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="3b0fe-125">Office. Context. Mailbox. Item. чувствительность</span><span class="sxs-lookup"><span data-stu-id="3b0fe-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="3b0fe-126">Добавлено новое свойство, представляющее чувствительность встречи.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="3b0fe-127">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3b0fe-127">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="3b0fe-128">Office. MailboxEnums. Аппоинтментсенситивититипе</span><span class="sxs-lookup"><span data-stu-id="3b0fe-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="3b0fe-129">Добавлено новое перечисление `AppointmentSensitivityType` , представляющее параметры конфиденциальности, доступные для встречи.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="3b0fe-130">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3b0fe-130">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="3b0fe-131">Добавление при отправке</span><span class="sxs-lookup"><span data-stu-id="3b0fe-131">Append on send</span></span>

<span data-ttu-id="3b0fe-132">Для получения сведений об использовании функции "присоединение к отправке", ознакомьтесь со статьей [Реализация добавления при отправке в надстройке Outlook](../../../outlook/append-on-send.md).</span><span class="sxs-lookup"><span data-stu-id="3b0fe-132">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="3b0fe-133">Office. Context. Mailbox. Item. Body. Аппендонсендасинк</span><span class="sxs-lookup"><span data-stu-id="3b0fe-133">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="3b0fe-134">Добавлена новая функция для `Body` объекта, который добавляет данные в конец тела элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-134">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="3b0fe-135">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="3b0fe-135">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="3b0fe-136">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="3b0fe-136">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="3b0fe-137">Добавлен новый элемент в манифест, где `AppendOnSend` расширенное разрешение должно быть включено в коллекцию расширенных разрешений.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-137">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="3b0fe-138">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="3b0fe-138">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="3b0fe-139">Активация на основе событий</span><span class="sxs-lookup"><span data-stu-id="3b0fe-139">Event-based activation</span></span>

<span data-ttu-id="3b0fe-140">Добавлена поддержка функций активации на основе событий в надстройках Outlook. Чтобы узнать больше, ознакомьтесь со статьей [Настройка надстройки Outlook для активации на основе событий](../../../outlook/autolaunch.md) .</span><span class="sxs-lookup"><span data-stu-id="3b0fe-140">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="3b0fe-141">Точка расширения Лаунчевент</span><span class="sxs-lookup"><span data-stu-id="3b0fe-141">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="3b0fe-142">Добавлена `LaunchEvent` Поддержка точек расширения для манифеста.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-142">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="3b0fe-143">Он настраивает функции активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-143">It configures event-based activation functionality.</span></span>

<span data-ttu-id="3b0fe-144">**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="3b0fe-144">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="3b0fe-145">Элемент манифеста Лаунчевентс</span><span class="sxs-lookup"><span data-stu-id="3b0fe-145">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="3b0fe-146">Добавлен `LaunchEvents` элемент для манифеста.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-146">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="3b0fe-147">Он поддерживает настройку функций активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-147">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="3b0fe-148">**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="3b0fe-148">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="3b0fe-149">Элемент манифеста среды выполнения</span><span class="sxs-lookup"><span data-stu-id="3b0fe-149">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="3b0fe-150">Добавлена поддержка Outlook для `Runtimes` элемента manifest.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-150">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="3b0fe-151">Он ссылается на HTML-и JavaScript-файлы, необходимые для функции активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-151">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="3b0fe-152">**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="3b0fe-152">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="3b0fe-153">Получение всех настраиваемых свойств</span><span class="sxs-lookup"><span data-stu-id="3b0fe-153">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="3b0fe-154">CustomProperties. Жеталл</span><span class="sxs-lookup"><span data-stu-id="3b0fe-154">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

<span data-ttu-id="3b0fe-155">Добавлена новая функция для `CustomProperties` объекта, который получает все настраиваемые свойства.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-155">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="3b0fe-156">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современная), Outlook на Mac (подключено к подписке на Office 365), Outlook на Android, Outlook на iOS</span><span class="sxs-lookup"><span data-stu-id="3b0fe-156">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to Office 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="3b0fe-157">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="3b0fe-157">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="3b0fe-158">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="3b0fe-158">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="3b0fe-159">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="3b0fe-159">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="3b0fe-160">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="3b0fe-160">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="3b0fe-161">Подпись почты</span><span class="sxs-lookup"><span data-stu-id="3b0fe-161">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="3b0fe-162">Office. Context. Mailbox. Item. Body. Сетсигнатуреасинк</span><span class="sxs-lookup"><span data-stu-id="3b0fe-162">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="3b0fe-163">Добавлена новая функция для `Body` объекта, который добавляет или заменяет подпись в теле элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-163">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="3b0fe-164">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="3b0fe-164">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="3b0fe-165">Office. Context. Mailbox. Item. Дисаблеклиентсигнатуреасинк</span><span class="sxs-lookup"><span data-stu-id="3b0fe-165">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="3b0fe-166">Добавлена новая функция, которая отключает подпись клиента для отправляющего почтового ящика в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-166">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="3b0fe-167">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="3b0fe-167">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="3b0fe-168">Office. Context. Mailbox. Item. Жеткомпосетипеасинк</span><span class="sxs-lookup"><span data-stu-id="3b0fe-168">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="3b0fe-169">Добавлена новая функция, которая получает тип сообщения "создание" в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-169">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="3b0fe-170">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="3b0fe-170">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="3b0fe-171">Office. Context. Mailbox. Item. Исклиентсигнатуринабледасинк</span><span class="sxs-lookup"><span data-stu-id="3b0fe-171">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="3b0fe-172">Добавлена новая функция, проверяющая, включена ли подпись клиента для элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-172">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="3b0fe-173">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="3b0fe-173">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="3b0fe-174">Office. MailboxEnums. Компосетипе</span><span class="sxs-lookup"><span data-stu-id="3b0fe-174">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="3b0fe-175">Добавлено новое перечисление `ComposeType` , доступное в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-175">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="3b0fe-176">**Доступно в**: Outlook в Windows (подключено к подписке Office 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="3b0fe-176">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="3b0fe-177">Тема Office</span><span class="sxs-lookup"><span data-stu-id="3b0fe-177">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="3b0fe-178">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="3b0fe-178">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="3b0fe-179">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-179">Added ability to get Office theme.</span></span>

<span data-ttu-id="3b0fe-180">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3b0fe-180">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="3b0fe-181">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="3b0fe-181">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="3b0fe-182">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-182">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="3b0fe-183">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365)</span><span class="sxs-lookup"><span data-stu-id="3b0fe-183">**Available in**: Outlook on Windows (connected to Office 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="3b0fe-184">Единый вход (SSO)</span><span class="sxs-lookup"><span data-stu-id="3b0fe-184">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="3b0fe-185">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="3b0fe-185">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="3b0fe-186">Добавлена возможность доступа к `getAccessToken`, что позволяет надстройкам [получать маркер доступа](../../../outlook/authenticate-a-user-with-an-sso-token.md) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="3b0fe-186">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="3b0fe-187">**Доступно в** Outlook для Windows (версия, подключенная к подписке на Office 365), Outlook для Mac (версия, подключенная к подписке на Office 365), Outlook в Интернете (современная версия), Outlook в Интернете (классическая версия)</span><span class="sxs-lookup"><span data-stu-id="3b0fe-187">**Available in**: Outlook on Windows (connected to Office 365 subscription), Outlook on Mac (connected to Office 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="3b0fe-188">См. также</span><span class="sxs-lookup"><span data-stu-id="3b0fe-188">See also</span></span>

- [<span data-ttu-id="3b0fe-189">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3b0fe-189">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="3b0fe-190">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="3b0fe-190">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="3b0fe-191">Начало работы</span><span class="sxs-lookup"><span data-stu-id="3b0fe-191">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="3b0fe-192">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="3b0fe-192">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
