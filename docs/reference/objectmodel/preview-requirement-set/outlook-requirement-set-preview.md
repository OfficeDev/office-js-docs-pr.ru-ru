---
title: Outlook набор требований к предварительному просмотру API надстройки
description: Функции и API, которые в настоящее время находятся в предварительном Outlook надстройки.
ms.date: 06/15/2021
localization_priority: Normal
ms.openlocfilehash: f9d8afc2b4347a8fb13f8ab98a163fb63968123f
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007764"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="47bd5-103">Outlook набор требований к предварительному просмотру API надстройки</span><span class="sxs-lookup"><span data-stu-id="47bd5-103">Outlook add-in API preview requirement set</span></span>

<span data-ttu-id="47bd5-104">Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="47bd5-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="47bd5-105">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="47bd5-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="47bd5-106">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="47bd5-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="47bd5-107">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="47bd5-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="47bd5-108">Можно просмотреть функции в Outlook в Интернете, настроив целевой выпуск на [Microsoft 365 клиента.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="47bd5-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="47bd5-109">"Настройка доступа к предварительному просмотру" отмечена на этой странице для применимых функций.</span><span class="sxs-lookup"><span data-stu-id="47bd5-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="47bd5-110">Для других функций вы можете запросить доступ к битам предварительного просмотра для Outlook в Интернете с помощью Microsoft 365 учетной записи, заполнив и подав [эту форму.](https://aka.ms/OWAPreview)</span><span class="sxs-lookup"><span data-stu-id="47bd5-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="47bd5-111">В этих функциях отмечен "Запрос доступа к предварительному просмотру".</span><span class="sxs-lookup"><span data-stu-id="47bd5-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="47bd5-112">Набор требований предварительного просмотра включает все функции [набора требований 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="47bd5-112">The preview requirement set includes all of the features of [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="47bd5-113">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="47bd5-113">Features in preview</span></span>

<span data-ttu-id="47bd5-114">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="47bd5-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="47bd5-115">Активация надстройки для элементов, защищенных управлением правами на информацию (IRM)</span><span class="sxs-lookup"><span data-stu-id="47bd5-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="47bd5-116">Надстройки теперь могут активироваться в пунктах, защищенных IRM.</span><span class="sxs-lookup"><span data-stu-id="47bd5-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="47bd5-117">Чтобы включить эту возможность, администратору клиента необходимо включить право использования, установив в Office параметр Разрешить программный `OBJMODEL` доступ. </span><span class="sxs-lookup"><span data-stu-id="47bd5-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="47bd5-118">Дополнительные [сведения см. в](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) дополнительных сведениях о правах и описаниях использования.</span><span class="sxs-lookup"><span data-stu-id="47bd5-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="47bd5-119">**Доступно в**: Outlook на Windows, начиная со сборки 13229.10000 (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="47bd5-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="47bd5-120">Дополнительные свойства календаря</span><span class="sxs-lookup"><span data-stu-id="47bd5-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="47bd5-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="47bd5-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="47bd5-122">Добавлен новый объект, который представляет свойство события на весь день встречи в режиме Compose.</span><span class="sxs-lookup"><span data-stu-id="47bd5-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="47bd5-123">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="47bd5-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="47bd5-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="47bd5-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="47bd5-125">Добавлен новый объект, который представляет чувствительность встречи в режиме Compose.</span><span class="sxs-lookup"><span data-stu-id="47bd5-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="47bd5-126">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="47bd5-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="47bd5-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="47bd5-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="47bd5-128">Добавлено новое свойство, которое представляет, если встреча является событием на весь день.</span><span class="sxs-lookup"><span data-stu-id="47bd5-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="47bd5-129">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="47bd5-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="47bd5-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="47bd5-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="47bd5-131">Добавлено новое свойство, которое представляет чувствительность встречи.</span><span class="sxs-lookup"><span data-stu-id="47bd5-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="47bd5-132">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="47bd5-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="47bd5-133">Office. MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="47bd5-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="47bd5-134">Добавлен новый `AppointmentSensitivityType` переумыв, который представляет параметры чувствительности, доступные при встрече.</span><span class="sxs-lookup"><span data-stu-id="47bd5-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="47bd5-135">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="47bd5-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="47bd5-136">Активация на основе событий</span><span class="sxs-lookup"><span data-stu-id="47bd5-136">Event-based activation</span></span>

<span data-ttu-id="47bd5-137">Эта функция была выпущена в [наборе требований 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="47bd5-137">This feature was released in [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="47bd5-138">Однако дополнительные события теперь доступны в предварительном просмотре.</span><span class="sxs-lookup"><span data-stu-id="47bd5-138">However, additional events are now available in preview.</span></span> <span data-ttu-id="47bd5-139">Дополнительные дополнительные ссылки на [поддерживаемые события.](../../../outlook/autolaunch.md#supported-events)</span><span class="sxs-lookup"><span data-stu-id="47bd5-139">To learn more, refer to [Supported events](../../../outlook/autolaunch.md#supported-events).</span></span>

<span data-ttu-id="47bd5-140">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="47bd5-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="47bd5-141">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="47bd5-141">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="47bd5-142">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="47bd5-142">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="47bd5-143">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="47bd5-143">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="47bd5-144">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="47bd5-144">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="47bd5-145">Тема Office</span><span class="sxs-lookup"><span data-stu-id="47bd5-145">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="47bd5-146">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="47bd5-146">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="47bd5-147">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="47bd5-147">Added ability to get Office theme.</span></span>

<span data-ttu-id="47bd5-148">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="47bd5-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="47bd5-149">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="47bd5-149">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="47bd5-150">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="47bd5-150">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="47bd5-151">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="47bd5-151">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="47bd5-152">Данные сеансов</span><span class="sxs-lookup"><span data-stu-id="47bd5-152">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="47bd5-153">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="47bd5-153">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="47bd5-154">Добавлен новый объект, который представляет данные сеанса элемента.</span><span class="sxs-lookup"><span data-stu-id="47bd5-154">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="47bd5-155">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="47bd5-155">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="47bd5-156">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="47bd5-156">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="47bd5-157">Добавлено новое свойство для управления данными сеанса элемента в режиме Compose.</span><span class="sxs-lookup"><span data-stu-id="47bd5-157">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="47bd5-158">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="47bd5-158">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="shared-mailboxes"></a><span data-ttu-id="47bd5-159">Общие почтовые ящики</span><span class="sxs-lookup"><span data-stu-id="47bd5-159">Shared mailboxes</span></span>

<span data-ttu-id="47bd5-160">Поддержка функций для общих папок (т. е. доступа делегатов) была выпущена в наборе [требований 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="47bd5-160">Feature support for shared folders (that is, delegate access) was released in [requirement set 1.8](../requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="47bd5-161">Однако поддержка общих почтовых ящиков теперь доступна в предварительном просмотре.</span><span class="sxs-lookup"><span data-stu-id="47bd5-161">However, support for shared mailboxes is now available in preview.</span></span> <span data-ttu-id="47bd5-162">Чтобы узнать больше, обратитесь к [разделу Включить общие папки и сценарии общих почтовых ящиков.](../../../outlook/delegate-access.md)</span><span class="sxs-lookup"><span data-stu-id="47bd5-162">To learn more, refer to [Enable shared folders and shared mailbox scenarios](../../../outlook/delegate-access.md).</span></span>

<span data-ttu-id="47bd5-163">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="47bd5-163">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="47bd5-164">См. также</span><span class="sxs-lookup"><span data-stu-id="47bd5-164">See also</span></span>

- [<span data-ttu-id="47bd5-165">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="47bd5-165">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="47bd5-166">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="47bd5-166">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="47bd5-167">Начало работы</span><span class="sxs-lookup"><span data-stu-id="47bd5-167">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="47bd5-168">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="47bd5-168">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
