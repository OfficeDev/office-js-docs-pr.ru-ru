---
title: Предварительная версия набора обязательных элементов API для надстройки Outlook
description: Функции и API, которые в настоящее время находятся в режиме предварительной версии для надстроек Outlook.
ms.date: 09/02/2020
localization_priority: Normal
ms.openlocfilehash: 24cad394f0f3ffb95a05a81ccb38ee4aa72a3797
ms.sourcegitcommit: 83f9a2fdff81ca421cd23feea103b9b60895cab4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/11/2020
ms.locfileid: "47431068"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="9cbcb-103">Предварительная версия набора обязательных элементов API для надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="9cbcb-103">Outlook add-in API Preview requirement set</span></span>

<span data-ttu-id="9cbcb-104">Подмножество API надстройки Outlook в API JavaScript для Office включает объекты, методы, свойства и события, которые можно использовать в надстройке Outlook.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="9cbcb-105">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="9cbcb-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="9cbcb-106">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="9cbcb-107">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="9cbcb-108">Вы можете предварительно просмотреть функции в Outlook в Интернете, [настроив целевой выпуск на клиенте Microsoft 365](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span><span class="sxs-lookup"><span data-stu-id="9cbcb-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="9cbcb-109">"Настройка предварительного доступа" отмечено на этой странице в соответствующих возможностях.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="9cbcb-110">Для других функций вы можете запросить доступ к предварительной версии BITS для Outlook в Интернете, используя свою учетную запись Microsoft 365, заполнив и отправив [эту форму](https://aka.ms/OWAPreview).</span><span class="sxs-lookup"><span data-stu-id="9cbcb-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="9cbcb-111">В этих функциях указано "запросить доступ к предварительному доступу".</span><span class="sxs-lookup"><span data-stu-id="9cbcb-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="9cbcb-112">Предварительная версия набора обязательных элементов включает все возможности [набора обязательных элементов 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="9cbcb-112">The Preview Requirement set includes all of the features of [Requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="9cbcb-113">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="9cbcb-113">Features in preview</span></span>

<span data-ttu-id="9cbcb-114">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="9cbcb-115">Активация надстройки для элементов, защищенных службой управления правами на доступ к данным (IRM)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="9cbcb-116">Теперь надстройки можно активировать на элементах, защищенных с помощью управления правами на доступ к данным.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="9cbcb-117">Чтобы включить эту возможность, администратору клиента необходимо включить `OBJMODEL` право на использование, установив параметр **Разрешить программный доступ к** настраиваемой политике в Office.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="9cbcb-118">Для получения дополнительных сведений ознакомьтесь [с разрешениями и описаниями использования](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) .</span><span class="sxs-lookup"><span data-stu-id="9cbcb-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="9cbcb-119">**Доступно в**: Outlook в Windows, начиная с сборки 13229,10000 (подключены к подписке Microsoft 365).</span><span class="sxs-lookup"><span data-stu-id="9cbcb-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="9cbcb-120">Дополнительные свойства календаря</span><span class="sxs-lookup"><span data-stu-id="9cbcb-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="9cbcb-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="9cbcb-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9cbcb-122">Добавлен новый объект, представляющий свойство события "целый день" для встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="9cbcb-123">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="9cbcb-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="9cbcb-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9cbcb-125">Добавлен новый объект, представляющий чувствительность встречи в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="9cbcb-126">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="9cbcb-127">Office. Context. Mailbox. Item. Исаллдайевент</span><span class="sxs-lookup"><span data-stu-id="9cbcb-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9cbcb-128">Добавлено новое свойство, которое указывает, является ли встреча событием на целый день.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="9cbcb-129">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="9cbcb-130">Office. Context. Mailbox. Item. чувствительность</span><span class="sxs-lookup"><span data-stu-id="9cbcb-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9cbcb-131">Добавлено новое свойство, представляющее чувствительность встречи.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="9cbcb-132">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="9cbcb-133">Office. MailboxEnums. Аппоинтментсенситивититипе</span><span class="sxs-lookup"><span data-stu-id="9cbcb-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9cbcb-134">Добавлено новое перечисление `AppointmentSensitivityType` , представляющее параметры конфиденциальности, доступные для встречи.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="9cbcb-135">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="append-on-send"></a><span data-ttu-id="9cbcb-136">Добавление при отправке</span><span class="sxs-lookup"><span data-stu-id="9cbcb-136">Append on send</span></span>

<span data-ttu-id="9cbcb-137">Для получения сведений об использовании функции "присоединение к отправке", ознакомьтесь со статьей [Реализация добавления при отправке в надстройке Outlook](../../../outlook/append-on-send.md).</span><span class="sxs-lookup"><span data-stu-id="9cbcb-137">To learn about using the append-on-send feature, see [Implement append on send in your Outlook add-in](../../../outlook/append-on-send.md).</span></span>

#### <a name="officecontextmailboxitembodyappendonsendasync"></a>[<span data-ttu-id="9cbcb-138">Office. Context. Mailbox. Item. Body. Аппендонсендасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-138">Office.context.mailbox.item.body.appendOnSendAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#appendonsendasync-data--options--callback-)

<span data-ttu-id="9cbcb-139">Добавлена новая функция для `Body` объекта, который добавляет данные в конец тела элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-139">Added a new function to the `Body` object that appends data to the end of the item body in Compose mode.</span></span>

<span data-ttu-id="9cbcb-140">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9cbcb-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="extendedpermissions"></a>[<span data-ttu-id="9cbcb-141">ExtendedPermissions</span><span class="sxs-lookup"><span data-stu-id="9cbcb-141">ExtendedPermissions</span></span>](../../manifest/extendedpermissions.md)

<span data-ttu-id="9cbcb-142">Добавлен новый элемент в манифест, где `AppendOnSend` расширенное разрешение должно быть включено в коллекцию расширенных разрешений.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-142">Added a new element to the manifest where the `AppendOnSend` extended permission must be included in the collection of extended permissions.</span></span>

<span data-ttu-id="9cbcb-143">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9cbcb-143">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="async-versions-of-display-apis"></a><span data-ttu-id="9cbcb-144">Асинхронные версии `display` API</span><span class="sxs-lookup"><span data-stu-id="9cbcb-144">Async versions of `display` APIs</span></span>

#### <a name="officecontextmailboxdisplayappointmentformasync"></a>[<span data-ttu-id="9cbcb-145">Office. Context. Mailbox. Дисплайаппоинтментформасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-145">Office.context.mailbox.displayAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displayappointmentformasync-itemid--options--callback-)

<span data-ttu-id="9cbcb-146">Добавлена новая функция для `Mailbox` объекта, отображающего существующую встречу.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-146">Added a new function to the `Mailbox` object that displays an existing appointment.</span></span> <span data-ttu-id="9cbcb-147">Это асинхронная версия `displayAppointmentForm` метода.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-147">This is the async version of the `displayAppointmentForm` method.</span></span>

<span data-ttu-id="9cbcb-148">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaymessageformasync"></a>[<span data-ttu-id="9cbcb-149">Office. Context. Mailbox. Дисплаймессажеформасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-149">Office.context.mailbox.displayMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaymessageformasync-itemid--options--callback-)

<span data-ttu-id="9cbcb-150">Добавлена новая функция для `Mailbox` объекта, отображающего существующее сообщение.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-150">Added a new function to the `Mailbox` object that displays an existing message.</span></span> <span data-ttu-id="9cbcb-151">Это асинхронная версия `displayMessageForm` метода.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-151">This is the async version of the `displayMessageForm` method.</span></span>

<span data-ttu-id="9cbcb-152">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-152">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaynewappointmentformasync"></a>[<span data-ttu-id="9cbcb-153">Office. Context. Mailbox. Дисплайневаппоинтментформасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-153">Office.context.mailbox.displayNewAppointmentFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewappointmentformasync-parameters--options--callback-)

<span data-ttu-id="9cbcb-154">Добавлена новая функция для `Mailbox` объекта, отображающего новую форму встречи.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-154">Added a new function to the `Mailbox` object that displays a new appointment form.</span></span> <span data-ttu-id="9cbcb-155">Это асинхронная версия `displayNewAppointmentForm` метода.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-155">This is the async version of the `displayNewAppointmentForm` method.</span></span>

<span data-ttu-id="9cbcb-156">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-156">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxdisplaynewmessageformasync"></a>[<span data-ttu-id="9cbcb-157">Office. Context. Mailbox. Дисплайневмессажеформасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-157">Office.context.mailbox.displayNewMessageFormAsync</span></span>](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#displaynewmessageformasync-parameters--options--callback-)

<span data-ttu-id="9cbcb-158">Добавлена новая функция для `Mailbox` объекта, отображающего форму нового сообщения.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-158">Added a new function to the `Mailbox` object that displays a new message form.</span></span> <span data-ttu-id="9cbcb-159">Это асинхронная версия `displayNewMessageForm` метода.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-159">This is the async version of the `displayNewMessageForm` method.</span></span>

<span data-ttu-id="9cbcb-160">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-160">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyallformasync"></a>[<span data-ttu-id="9cbcb-161">Office. Context. Mailbox. Item. Дисплайрепляллформасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-161">Office.context.mailbox.item.displayReplyAllFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9cbcb-162">Добавлена новая функция для `Item` объекта, отображающего форму "ответить всем" в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-162">Added a new function to the `Item` object that displays the "Reply all" form in Read mode.</span></span> <span data-ttu-id="9cbcb-163">Это асинхронная версия `displayReplyAllForm` метода.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-163">This is the async version of the `displayReplyAllForm` method.</span></span>

<span data-ttu-id="9cbcb-164">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-164">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemdisplayreplyformasync"></a>[<span data-ttu-id="9cbcb-165">Office. Context. Mailbox. Item. Дисплайреплиформасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-165">Office.context.mailbox.item.displayReplyFormAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9cbcb-166">Добавлена новая функция для `Item` объекта, отображающего форму "Reply" в режиме чтения.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-166">Added a new function to the `Item` object that displays the "Reply" form in Read mode.</span></span> <span data-ttu-id="9cbcb-167">Это асинхронная версия `displayReplyForm` метода.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-167">This is the async version of the `displayReplyForm` method.</span></span>

<span data-ttu-id="9cbcb-168">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-168">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="9cbcb-169">Активация на основе событий</span><span class="sxs-lookup"><span data-stu-id="9cbcb-169">Event-based activation</span></span>

<span data-ttu-id="9cbcb-170">Добавлена поддержка функций активации на основе событий в надстройках Outlook. Чтобы узнать больше, ознакомьтесь со статьей [Настройка надстройки Outlook для активации на основе событий](../../../outlook/autolaunch.md) .</span><span class="sxs-lookup"><span data-stu-id="9cbcb-170">Added support for event-based activation functionality in Outlook add-ins. See [Configure your Outlook add-in for event-based activation](../../../outlook/autolaunch.md) to learn more.</span></span>

#### <a name="launchevent-extension-point"></a>[<span data-ttu-id="9cbcb-171">Точка расширения Лаунчевент</span><span class="sxs-lookup"><span data-stu-id="9cbcb-171">LaunchEvent extension point</span></span>](../../manifest/extensionpoint.md#launchevent-preview)

<span data-ttu-id="9cbcb-172">Добавлена `LaunchEvent` Поддержка точек расширения для манифеста.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-172">Added `LaunchEvent` extension point support to manifest.</span></span> <span data-ttu-id="9cbcb-173">Он настраивает функции активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-173">It configures event-based activation functionality.</span></span>

<span data-ttu-id="9cbcb-174">**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="9cbcb-174">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="launchevents-manifest-element"></a>[<span data-ttu-id="9cbcb-175">Элемент манифеста Лаунчевентс</span><span class="sxs-lookup"><span data-stu-id="9cbcb-175">LaunchEvents manifest element</span></span>](../../manifest/launchevents.md)

<span data-ttu-id="9cbcb-176">Добавлен `LaunchEvents` элемент для манифеста.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-176">Added `LaunchEvents` element to manifest.</span></span> <span data-ttu-id="9cbcb-177">Он поддерживает настройку функций активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-177">It supports configuring event-based activation functionality.</span></span>

<span data-ttu-id="9cbcb-178">**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="9cbcb-178">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

#### <a name="runtimes-manifest-element"></a>[<span data-ttu-id="9cbcb-179">Элемент манифеста среды выполнения</span><span class="sxs-lookup"><span data-stu-id="9cbcb-179">Runtimes manifest element</span></span>](../../manifest/runtimes.md)

<span data-ttu-id="9cbcb-180">Добавлена поддержка Outlook для `Runtimes` элемента manifest.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-180">Added Outlook support to the `Runtimes` manifest element.</span></span> <span data-ttu-id="9cbcb-181">Он ссылается на HTML-и JavaScript-файлы, необходимые для функции активации на основе событий.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-181">It references the HTML and JavaScript files needed for event-based activation functionality.</span></span>

<span data-ttu-id="9cbcb-182">**Доступно в**: Outlook в Интернете (современный, [запрос предварительной версии Access](https://aka.ms/OWAPreview))</span><span class="sxs-lookup"><span data-stu-id="9cbcb-182">**Available in**: Outlook on the web (modern, [Request preview access](https://aka.ms/OWAPreview))</span></span>

<br>

---

---

### <a name="get-all-custom-properties"></a><span data-ttu-id="9cbcb-183">Получение всех настраиваемых свойств</span><span class="sxs-lookup"><span data-stu-id="9cbcb-183">Get all custom properties</span></span>

#### <a name="custompropertiesgetall"></a>[<span data-ttu-id="9cbcb-184">CustomProperties. Жеталл</span><span class="sxs-lookup"><span data-stu-id="9cbcb-184">CustomProperties.getAll</span></span>](/javascript/api/outlook/office.customproperties?view=outlook-js-preview&preserve-view=true#getall--)

<span data-ttu-id="9cbcb-185">Добавлена новая функция для `CustomProperties` объекта, который получает все настраиваемые свойства.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-185">Added a new function to the `CustomProperties` object that gets all custom properties.</span></span>

<span data-ttu-id="9cbcb-186">**Доступно в**: Outlook в Windows (подключенном к подписке на Microsoft 365), Outlook в Интернете (современная), Outlook на Mac (подключено к подписке Microsoft 365), Outlook на Android, Outlook на iOS</span><span class="sxs-lookup"><span data-stu-id="9cbcb-186">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on Android, Outlook on iOS</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="9cbcb-187">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="9cbcb-187">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="9cbcb-188">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="9cbcb-188">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9cbcb-189">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="9cbcb-189">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="9cbcb-190">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (классическая)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-190">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (classic)</span></span>

<br>

---

---

### <a name="mail-signature"></a><span data-ttu-id="9cbcb-191">Подпись почты</span><span class="sxs-lookup"><span data-stu-id="9cbcb-191">Mail signature</span></span>

#### <a name="officecontextmailboxitembodysetsignatureasync"></a>[<span data-ttu-id="9cbcb-192">Office. Context. Mailbox. Item. Body. Сетсигнатуреасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-192">Office.context.mailbox.item.body.setSignatureAsync</span></span>](/javascript/api/outlook/office.body?view=outlook-js-preview&preserve-view=true#setsignatureasync-data--options--callback-)

<span data-ttu-id="9cbcb-193">Добавлена новая функция для `Body` объекта, который добавляет или заменяет подпись в теле элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-193">Added a new function to the `Body` object that adds or replaces the signature in the item body in Compose mode.</span></span>

<span data-ttu-id="9cbcb-194">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9cbcb-194">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemdisableclientsignatureasync"></a>[<span data-ttu-id="9cbcb-195">Office. Context. Mailbox. Item. Дисаблеклиентсигнатуреасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-195">Office.context.mailbox.item.disableClientSignatureAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9cbcb-196">Добавлена новая функция, которая отключает подпись клиента для отправляющего почтового ящика в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-196">Added a new function that disables the client signature for the sending mailbox in Compose mode.</span></span>

<span data-ttu-id="9cbcb-197">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9cbcb-197">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemgetcomposetypeasync"></a>[<span data-ttu-id="9cbcb-198">Office. Context. Mailbox. Item. Жеткомпосетипеасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-198">Office.context.mailbox.item.getComposeTypeAsync</span></span>](/javascript/api/outlook/office.messagecompose?view=outlook-js-preview&preserve-view=true#getcomposetypeasync-options--callback-)

<span data-ttu-id="9cbcb-199">Добавлена новая функция, которая получает тип сообщения "создание" в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-199">Added a new function that gets the compose type of a message in Compose mode.</span></span>

<span data-ttu-id="9cbcb-200">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9cbcb-200">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officecontextmailboxitemisclientsignatureenabledasync"></a>[<span data-ttu-id="9cbcb-201">Office. Context. Mailbox. Item. Исклиентсигнатуринабледасинк</span><span class="sxs-lookup"><span data-stu-id="9cbcb-201">Office.context.mailbox.item.isClientSignatureEnabledAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="9cbcb-202">Добавлена новая функция, проверяющая, включена ли подпись клиента для элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-202">Added a new function that checks if the client signature is enabled on the item in Compose mode.</span></span>

<span data-ttu-id="9cbcb-203">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9cbcb-203">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

#### <a name="officemailboxenumscomposetype"></a>[<span data-ttu-id="9cbcb-204">Office. MailboxEnums. Компосетипе</span><span class="sxs-lookup"><span data-stu-id="9cbcb-204">Office.MailboxEnums.ComposeType</span></span>](/javascript/api/outlook/office.mailboxenums.composetype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="9cbcb-205">Добавлено новое перечисление `ComposeType` , доступное в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-205">Added a new enum `ComposeType` available in Compose mode.</span></span>

<span data-ttu-id="9cbcb-206">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современная, [Настройка предварительного доступа](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span><span class="sxs-lookup"><span data-stu-id="9cbcb-206">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern, [Configure preview access](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center))</span></span>

<br>

---

---

### <a name="notification-messages-with-actions"></a><span data-ttu-id="9cbcb-207">Сообщения уведомления с действиями</span><span class="sxs-lookup"><span data-stu-id="9cbcb-207">Notification messages with actions</span></span>

<span data-ttu-id="9cbcb-208">Эта функция позволяет надстройке включать сообщение уведомления с дополнительным **действием, кроме действия по** умолчанию.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-208">This feature allows your add-in to include a notification message with a custom action besides the default **Dismiss** action.</span></span>

#### <a name="officenotificationmessagedetailsactions"></a>[<span data-ttu-id="9cbcb-209">Office. NotificationMessageDetails. Actions</span><span class="sxs-lookup"><span data-stu-id="9cbcb-209">Office.NotificationMessageDetails.actions</span></span>](/javascript/api/outlook/office.notificationmessagedetails#actions)

<span data-ttu-id="9cbcb-210">Добавлено новое свойство, которое позволяет добавить `InsightMessage` уведомление с дополнительным действием.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-210">Added a new property that enables you to add an `InsightMessage` notification with a custom action.</span></span>

<span data-ttu-id="9cbcb-211">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-211">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officenotificationmessageaction"></a>[<span data-ttu-id="9cbcb-212">Office. Нотификатионмессажеактион</span><span class="sxs-lookup"><span data-stu-id="9cbcb-212">Office.NotificationMessageAction</span></span>](/javascript/api/outlook/office.notificationmessageaction)

<span data-ttu-id="9cbcb-213">Добавлен новый объект, в котором определяется дополнительное действие для `InsightMessage` уведомления.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-213">Added a new object where you define a custom action for your `InsightMessage` notification.</span></span>

<span data-ttu-id="9cbcb-214">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-214">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsactiontype"></a>[<span data-ttu-id="9cbcb-215">Office. MailboxEnums.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-215">Office.MailboxEnums.ActionType</span></span>](/javascript/api/outlook/office.mailboxenums.actiontype)

<span data-ttu-id="9cbcb-216">Добавлено новое перечисление `ActionType` .</span><span class="sxs-lookup"><span data-stu-id="9cbcb-216">Added a new enum `ActionType`.</span></span>

<span data-ttu-id="9cbcb-217">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-217">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officemailboxenumsitemnotificationmessagetypeinsightmessage"></a>[<span data-ttu-id="9cbcb-218">Office. MailboxEnums. Итемнотификатионмессажетипе. Инсигхтмессаже</span><span class="sxs-lookup"><span data-stu-id="9cbcb-218">Office.MailboxEnums.ItemNotificationMessageType.InsightMessage</span></span>](/javascript/api/outlook/office.mailboxenums.itemnotificationmessagetype)

<span data-ttu-id="9cbcb-219">Добавлен новый тип `InsightMessage` в `ItemNotificationMessageType` перечисление.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-219">Added a new type `InsightMessage` to the `ItemNotificationMessageType` enum.</span></span>

<span data-ttu-id="9cbcb-220">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook в Интернете (современный)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-220">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="9cbcb-221">Тема Office</span><span class="sxs-lookup"><span data-stu-id="9cbcb-221">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="9cbcb-222">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="9cbcb-222">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="9cbcb-223">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-223">Added ability to get Office theme.</span></span>

<span data-ttu-id="9cbcb-224">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-224">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="9cbcb-225">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="9cbcb-225">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="9cbcb-226">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-226">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="9cbcb-227">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-227">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="9cbcb-228">Данные сеансов</span><span class="sxs-lookup"><span data-stu-id="9cbcb-228">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="9cbcb-229">Office. Сессиондата</span><span class="sxs-lookup"><span data-stu-id="9cbcb-229">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="9cbcb-230">Добавлен новый объект, представляющий данные сеанса для элемента.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-230">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="9cbcb-231">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-231">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="9cbcb-232">Office. Context. Mailbox. Item. Сессиондата</span><span class="sxs-lookup"><span data-stu-id="9cbcb-232">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="9cbcb-233">Добавлено новое свойство для управления данными сеанса элемента в режиме создания.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-233">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="9cbcb-234">**Доступно в**: Outlook в Windows (подключено к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-234">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="single-sign-on-sso"></a><span data-ttu-id="9cbcb-235">Единый вход (SSO)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-235">Single sign-on (SSO)</span></span>

#### <a name="officeruntimeauthgetaccesstoken"></a>[<span data-ttu-id="9cbcb-236">OfficeRuntime.auth.getAccessToken</span><span class="sxs-lookup"><span data-stu-id="9cbcb-236">OfficeRuntime.auth.getAccessToken</span></span>](../../../develop/sso-in-office-add-ins.md#sso-api-reference)

<span data-ttu-id="9cbcb-237">Добавлена возможность доступа к `getAccessToken`, что позволяет надстройкам [получать маркер доступа](../../../outlook/authenticate-a-user-with-an-sso-token.md) для API Microsoft Graph.</span><span class="sxs-lookup"><span data-stu-id="9cbcb-237">Added access to `getAccessToken`, which allows add-ins to [get an access token](../../../outlook/authenticate-a-user-with-an-sso-token.md) for the Microsoft Graph API.</span></span>

<span data-ttu-id="9cbcb-238">**Доступно в**: Outlook в Windows (подключенном к подписке Microsoft 365), Outlook на Mac (подключен к подписке Microsoft 365), Outlook в Интернете (современный), Outlook в Интернете (классическая)</span><span class="sxs-lookup"><span data-stu-id="9cbcb-238">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on Mac (connected to a Microsoft 365 subscription), Outlook on the web (modern), Outlook on the web (classic)</span></span>

## <a name="see-also"></a><span data-ttu-id="9cbcb-239">См. также</span><span class="sxs-lookup"><span data-stu-id="9cbcb-239">See also</span></span>

- [<span data-ttu-id="9cbcb-240">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="9cbcb-240">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="9cbcb-241">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="9cbcb-241">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="9cbcb-242">Начало работы</span><span class="sxs-lookup"><span data-stu-id="9cbcb-242">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="9cbcb-243">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="9cbcb-243">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
