---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: Функции и API, которые в настоящее время находятся в режиме предварительной версии для надстроек Outlook.
ms.date: 07/10/2020
localization_priority: Normal
ms.openlocfilehash: 457195b7511d4dabca101242400d44154a57a781
ms.sourcegitcommit: 472b81642e9eb5fb2a55cd98a7b0826d37eb7f73
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2020
ms.locfileid: "45159222"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="37448-103">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="37448-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="37448-104">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="37448-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="37448-105">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="37448-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="37448-106">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="37448-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="37448-107">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="37448-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="37448-108">Вы можете предварительно просмотреть функции в Outlook в Интернете, [настроив целевой выпуск на клиенте Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="37448-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="37448-109">"Настройка предварительного доступа" отмечено на этой странице в соответствующих возможностях.</span><span class="sxs-lookup"><span data-stu-id="37448-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="37448-110">Для других функций вы можете запросить доступ к предварительной версии BITS для Outlook в Интернете, используя свою учетную запись Microsoft 365, заполнив и отправив [эту форму](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="37448-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="37448-111">В этих функциях указано "запросить доступ к предварительному доступу".</span><span class="sxs-lookup"><span data-stu-id="37448-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="37448-112">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="37448-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="37448-113">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="37448-113">Features in preview</span></span>

<span data-ttu-id="37448-114">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="37448-114">The following features are in preview.</span></span>

### <a name="additional-calendar-properties"></a><span data-ttu-id="37448-115">Дополнительные свойства календаря</span><span class="sxs-lookup"><span data-stu-id="37448-115">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="37448-116">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="37448-116">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview)

<span data-ttu-id="37448-117">Добавлен новый объект, представляющий свойство события "целый день" для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37448-117">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="37448-118">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-118">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="37448-119">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="37448-119">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview)

<span data-ttu-id="37448-120">Добавлен новый объект, представляющий чувствительность встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37448-120">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="37448-121">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-121">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="37448-122">Office. Context. Mailbox. Item. Исаллдайевент</span><span class="sxs-lookup"><span data-stu-id="37448-122">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="37448-123">Добавлено новое свойство, которое указывает, является ли встреча событием на целый день.</span><span class="sxs-lookup"><span data-stu-id="37448-123">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="37448-124">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-124">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="37448-125">Office. Context. Mailbox. Item. чувствительность</span><span class="sxs-lookup"><span data-stu-id="37448-125">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="37448-126">Добавлено новое свойство, представляющее чувствительность встречи.</span><span class="sxs-lookup"><span data-stu-id="37448-126">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="37448-127">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-127">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="37448-128">Office. MailboxEnums. Аппоинтментсенситивититипе</span><span class="sxs-lookup"><span data-stu-id="37448-128">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview)

<span data-ttu-id="37448-129">Добавлено новое перечисление `AppointmentSensitivityType` , представляющее параметры конфиденциальности, доступные для встречи.</span><span class="sxs-lookup"><span data-stu-id="37448-129">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="37448-130">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-130">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="37448-131">Добавление при отправке</span><span class="sxs-lookup"><span data-stu-id="37448-131">Append on send</span></span>

<span data-ttu-id="37448-132">Для получения сведений об использовании функции "присоединение к отправке", ознакомьтесь со статьей [Реализация добавления при отправке в надстройке Outlook](../../../outlook/append-on-send.md).</span><span class="sxs-lookup"><span data-stu-id="37448-132">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="37448-133">Office. Context. Mailbox. Item. Body. Аппендонсендасинк</span><span class="sxs-lookup"><span data-stu-id="37448-133">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#appendonsendasync-data--options--callback-)

<span data-ttu-id="37448-134">Добавлена новая функция для `Body` объекта, который добавляет данные в конец тела элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37448-134">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="37448-135">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="37448-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="37448-136">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="37448-136">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="37448-137">Добавлен новый элемент в манифест, где `AppendOnSend` расширенное разрешение должно быть включено в коллекцию расширенных разрешений.</span><span class="sxs-lookup"><span data-stu-id="37448-137">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="37448-138">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="37448-138">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="async-versions-of-display-apis"></a><span data-ttu-id="37448-139">Асинхронные версии `display` API</span><span class="sxs-lookup"><span data-stu-id="37448-139">Async versions of `display` APIs</span></span>

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[<span data-ttu-id="37448-140">Office. Context. Mailbox. Дисплайаппоинтментформасинк</span><span class="sxs-lookup"><span data-stu-id="37448-140">Office.context.mailbox.displayAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displayappointmentformasync-itemid--options--callback-)

<span data-ttu-id="37448-141">Добавлена новая функция для `Mailbox` объекта, отображающего существующую встречу.</span><span class="sxs-lookup"><span data-stu-id="37448-141">Added a new function to the `Mailbox` object that displays an existing appointment.</span></span> <span data-ttu-id="37448-142">Это асинхронная версия `displayAppointmentForm` метода.</span><span class="sxs-lookup"><span data-stu-id="37448-142">This is the async version of the `displayAppointmentForm` method.</span></span>

<span data-ttu-id="37448-143">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-143">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[<span data-ttu-id="37448-144">Office. Context. Mailbox. Дисплаймессажеформасинк</span><span class="sxs-lookup"><span data-stu-id="37448-144">Office.context.mailbox.displayMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaymessageformasync-itemid--options--callback-)

<span data-ttu-id="37448-145">Добавлена новая функция для `Mailbox` объекта, отображающего существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="37448-145">Added a new function to the `Mailbox` object that displays an existing message.</span></span> <span data-ttu-id="37448-146">Это асинхронная версия `displayMessageForm` метода.</span><span class="sxs-lookup"><span data-stu-id="37448-146">This is the async version of the `displayMessageForm` method.</span></span>

<span data-ttu-id="37448-147">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-147">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[<span data-ttu-id="37448-148">Office. Context. Mailbox. Дисплайневаппоинтментформасинк</span><span class="sxs-lookup"><span data-stu-id="37448-148">Office.context.mailbox.displayNewAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewappointmentformasync-parameters--options--callback-)

<span data-ttu-id="37448-149">Добавлена новая функция для `Mailbox` объекта, отображающего новую форму встречи.</span><span class="sxs-lookup"><span data-stu-id="37448-149">Added a new function to the `Mailbox` object that displays a new appointment form.</span></span> <span data-ttu-id="37448-150">Это асинхронная версия `displayNewAppointmentForm` метода.</span><span class="sxs-lookup"><span data-stu-id="37448-150">This is the async version of the `displayNewAppointmentForm` method.</span></span>

<span data-ttu-id="37448-151">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-151">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[<span data-ttu-id="37448-152">Office. Context. Mailbox. Дисплайневмессажеформасинк</span><span class="sxs-lookup"><span data-stu-id="37448-152">Office.context.mailbox.displayNewMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview#displaynewmessageformasync-parameters--options--callback-)

<span data-ttu-id="37448-153">Добавлена новая функция для `Mailbox` объекта, отображающего форму нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="37448-153">Added a new function to the `Mailbox` object that displays a new message form.</span></span> <span data-ttu-id="37448-154">Это асинхронная версия `displayNewMessageForm` метода.</span><span class="sxs-lookup"><span data-stu-id="37448-154">This is the async version of the `displayNewMessageForm` method.</span></span>

<span data-ttu-id="37448-155">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-155">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[<span data-ttu-id="37448-156">Office. Context. Mailbox. Item. Дисплайрепляллформасинк</span><span class="sxs-lookup"><span data-stu-id="37448-156">Office.context.mailbox.item.displayReplyAllFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="37448-157">Добавлена новая функция для `Item` объекта, отображающего форму "ответить всем" в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37448-157">Added a new function to the `Item` object that displays the "Reply all" form in Read mode.</span></span> <span data-ttu-id="37448-158">Это асинхронная версия `displayReplyAllForm` метода.</span><span class="sxs-lookup"><span data-stu-id="37448-158">This is the async version of the `displayReplyAllForm` method.</span></span>

<span data-ttu-id="37448-159">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-159">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[<span data-ttu-id="37448-160">Office. Context. Mailbox. Item. Дисплайреплиформасинк</span><span class="sxs-lookup"><span data-stu-id="37448-160">Office.context.mailbox.item.displayReplyFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="37448-161">Добавлена новая функция для `Item` объекта, отображающего форму "Reply" в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="37448-161">Added a new function to the `Item` object that displays the "Reply" form in Read mode.</span></span> <span data-ttu-id="37448-162">Это асинхронная версия `displayReplyForm` метода.</span><span class="sxs-lookup"><span data-stu-id="37448-162">This is the async version of the `displayReplyForm` method.</span></span>

<span data-ttu-id="37448-163">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="37448-164">Активация на основе событий</span><span class="sxs-lookup"><span data-stu-id="37448-164">Event-based activation</span></span>

<span data-ttu-id="37448-165">Добавлена поддержка функций активации на основе событий в надстройках Outlook. Чтобы узнать больше, ознакомьтесь со статьей [Настройка надстройки Outlook для активации на основе событий](../../../outlook/autolaunch.md) .</span><span class="sxs-lookup"><span data-stu-id="37448-165">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="37448-166">Точка расширения Лаунчевент</span><span class="sxs-lookup"><span data-stu-id="37448-166">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="37448-167">Добавлена `LaunchEvent` Поддержка точек расширения для манифеста.</span><span class="sxs-lookup"><span data-stu-id="37448-167">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="37448-168">Он настраивает функции активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="37448-168">It configures event-based activation functionality.</span></span>

<span data-ttu-id="37448-169">**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="37448-169">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="37448-170">Элемент манифеста Лаунчевентс</span><span class="sxs-lookup"><span data-stu-id="37448-170">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="37448-171">Добавлен `LaunchEvents` элемент для манифеста.</span><span class="sxs-lookup"><span data-stu-id="37448-171">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="37448-172">Он поддерживает настройку функций активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="37448-172">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="37448-173">**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="37448-173">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="37448-174">Элемент манифеста среды выполнения</span><span class="sxs-lookup"><span data-stu-id="37448-174">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="37448-175">Добавлена поддержка Outlook для `Runtimes` элемента manifest.</span><span class="sxs-lookup"><span data-stu-id="37448-175">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="37448-176">Он ссылается на HTML-и JavaScript-файлы, необходимые для функции активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="37448-176">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="37448-177">**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="37448-177">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="37448-178">Получение всех настраиваемых свойств</span><span class="sxs-lookup"><span data-stu-id="37448-178">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="37448-179">CustomProperties. Жеталл</span><span class="sxs-lookup"><span data-stu-id="37448-179">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview#getall--)

<span data-ttu-id="37448-180">Добавлена новая функция для `CustomProperties` объекта, который получает все настраиваемые свойства.</span><span class="sxs-lookup"><span data-stu-id="37448-180">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="37448-181">**Доступно в**: Outlook в Windows (подключенном к подписке на Microsoft 365), Outlook в Интернете (современная), Outlook на Mac (подключено к подписке Microsoft 365), Outlook на Android, Outlook на iOS</span><span class="sxs-lookup"><span data-stu-id="37448-181">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="37448-182">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="37448-182">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="37448-183">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="37448-183">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="37448-184">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="37448-184">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="37448-185">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (классическая)</span><span class="sxs-lookup"><span data-stu-id="37448-185">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="37448-186">Подпись почты</span><span class="sxs-lookup"><span data-stu-id="37448-186">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="37448-187">Office. Context. Mailbox. Item. Body. Сетсигнатуреасинк</span><span class="sxs-lookup"><span data-stu-id="37448-187">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview#setsignatureasync-data--options--callback-)

<span data-ttu-id="37448-188">Добавлена новая функция для `Body` объекта, который добавляет или заменяет подпись в теле элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37448-188">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="37448-189">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="37448-189">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="37448-190">Office. Context. Mailbox. Item. Дисаблеклиентсигнатуреасинк</span><span class="sxs-lookup"><span data-stu-id="37448-190">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="37448-191">Добавлена новая функция, которая отключает подпись клиента для отправляющего почтового ящика в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37448-191">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="37448-192">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="37448-192">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="37448-193">Office. Context. Mailbox. Item. Жеткомпосетипеасинк</span><span class="sxs-lookup"><span data-stu-id="37448-193">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview#getcomposetypeasync-options--callback-)

<span data-ttu-id="37448-194">Добавлена новая функция, которая получает тип сообщения "создание" в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37448-194">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="37448-195">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="37448-195">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="37448-196">Office. Context. Mailbox. Item. Исклиентсигнатуринабледасинк</span><span class="sxs-lookup"><span data-stu-id="37448-196">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="37448-197">Добавлена новая функция, проверяющая, включена ли подпись клиента для элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37448-197">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="37448-198">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="37448-198">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="37448-199">Office. MailboxEnums. Компосетипе</span><span class="sxs-lookup"><span data-stu-id="37448-199">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview)

<span data-ttu-id="37448-200">Добавлено новое перечисление `ComposeType` , доступное в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="37448-200">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="37448-201">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="37448-201">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="37448-202">Тема Office</span><span class="sxs-lookup"><span data-stu-id="37448-202">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="37448-203">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="37448-203">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="37448-204">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="37448-204">Added ability to get Office theme.</span></span>

<span data-ttu-id="37448-205">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-205">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="37448-206">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="37448-206">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="37448-207">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="37448-207">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="37448-208">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="37448-208">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="37448-209">Единый вход (SSO)</span><span class="sxs-lookup"><span data-stu-id="37448-209">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="37448-210">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="37448-210">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="37448-211">Добавлена возможность доступа к `getAccessToken`, что позволяет надстройкам [получать маркер доступа](../../../outlook/authenticate-a-user-with-an-sso-token.md) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="37448-211">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="37448-212">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook на Mac (подключен к подписке Microsoft 365), Outlook в Интернете (современный), Outlook в Интернете (классическая)</span><span class="sxs-lookup"><span data-stu-id="37448-212">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="37448-213">См. также</span><span class="sxs-lookup"><span data-stu-id="37448-213">See also</span></span>

- [<span data-ttu-id="37448-214">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="37448-214">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="37448-215">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="37448-215">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="37448-216">Начало работы</span><span class="sxs-lookup"><span data-stu-id="37448-216">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="37448-217">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="37448-217">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
