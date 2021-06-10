---
title: Outlook набор требований к предварительному просмотру API надстройки
description: Функции и API, которые в настоящее время находятся в предварительном Outlook надстройки.
ms.date: 06/08/2021
localization_priority: Normal
ms.openlocfilehash: c7ca92e6a30f3109baff5721ae4e9930ef23dc56
ms.sourcegitcommit: 5a151d4df81e5640363774406d0f329d6a0d3db8
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/09/2021
ms.locfileid: "52854013"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="96ec8-103">Outlook набор требований к предварительному просмотру API надстройки</span><span class="sxs-lookup"><span data-stu-id="96ec8-103">Outlook add-in API preview requirement set</span></span>

<span data-ttu-id="96ec8-104">Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="96ec8-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="96ec8-105">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="96ec8-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="96ec8-106">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="96ec8-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="96ec8-107">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="96ec8-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="96ec8-108">Вы можете просматривать функции Outlook в Интернете, настроив целевой выпуск на Microsoft 365 [клиента.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="96ec8-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="96ec8-109">"Настройка доступа к предварительному просмотру" отмечена на этой странице для применимых функций.</span><span class="sxs-lookup"><span data-stu-id="96ec8-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="96ec8-110">Для других функций вы можете запросить доступ к битам предварительного просмотра для Outlook веб-страницы с помощью Microsoft 365 учетной записи, заполнив и подав эту [форму.](https://aka.ms/OWAPreview)</span><span class="sxs-lookup"><span data-stu-id="96ec8-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="96ec8-111">В этих функциях отмечен "Запрос доступа к предварительному просмотру".</span><span class="sxs-lookup"><span data-stu-id="96ec8-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="96ec8-112">Набор требований предварительного просмотра включает все функции [набора требований 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="96ec8-112">The preview requirement set includes all of the features of [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="96ec8-113">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="96ec8-113">Features in preview</span></span>

<span data-ttu-id="96ec8-114">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="96ec8-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="96ec8-115">Активация надстройки для элементов, защищенных управлением правами на информацию (IRM)</span><span class="sxs-lookup"><span data-stu-id="96ec8-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="96ec8-116">Надстройки теперь могут активироваться в пунктах, защищенных IRM.</span><span class="sxs-lookup"><span data-stu-id="96ec8-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="96ec8-117">Чтобы включить эту возможность, администратору клиента необходимо включить право использования, установив в Office параметр Разрешить программный `OBJMODEL` доступ. </span><span class="sxs-lookup"><span data-stu-id="96ec8-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="96ec8-118">Дополнительные [сведения см. в](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) дополнительных сведениях о правах и описаниях использования.</span><span class="sxs-lookup"><span data-stu-id="96ec8-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="96ec8-119">**Доступно в**: Outlook на Windows, начиная со сборки 13229.10000 (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="96ec8-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="96ec8-120">Дополнительные свойства календаря</span><span class="sxs-lookup"><span data-stu-id="96ec8-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="96ec8-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="96ec8-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="96ec8-122">Добавлен новый объект, который представляет свойство события на весь день встречи в режиме Compose.</span><span class="sxs-lookup"><span data-stu-id="96ec8-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="96ec8-123">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="96ec8-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="96ec8-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="96ec8-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="96ec8-125">Добавлен новый объект, который представляет чувствительность встречи в режиме Compose.</span><span class="sxs-lookup"><span data-stu-id="96ec8-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="96ec8-126">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="96ec8-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="96ec8-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="96ec8-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="96ec8-128">Добавлено новое свойство, которое представляет, если встреча является событием на весь день.</span><span class="sxs-lookup"><span data-stu-id="96ec8-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="96ec8-129">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="96ec8-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="96ec8-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="96ec8-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="96ec8-131">Добавлено новое свойство, которое представляет чувствительность встречи.</span><span class="sxs-lookup"><span data-stu-id="96ec8-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="96ec8-132">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="96ec8-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="96ec8-133">Office. MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="96ec8-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="96ec8-134">Добавлен новый `AppointmentSensitivityType` переумыв, который представляет параметры чувствительности, доступные при встрече.</span><span class="sxs-lookup"><span data-stu-id="96ec8-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="96ec8-135">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="96ec8-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="event-based-activation"></a><span data-ttu-id="96ec8-136">Активация на основе событий</span><span class="sxs-lookup"><span data-stu-id="96ec8-136">Event-based activation</span></span>

<span data-ttu-id="96ec8-137">Эта функция была выпущена в [наборе требований 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="96ec8-137">This feature was released in [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span> <span data-ttu-id="96ec8-138">Однако дополнительные события теперь доступны в предварительном просмотре.</span><span class="sxs-lookup"><span data-stu-id="96ec8-138">However, additional events are now available in preview.</span></span> <span data-ttu-id="96ec8-139">Дополнительные дополнительные информации см. в [дополнительных подробной информации о поддерживаемых событиях.](../../../outlook/autolaunch.md#supported-events)</span><span class="sxs-lookup"><span data-stu-id="96ec8-139">To learn more, see [Supported events](../../../outlook/autolaunch.md#supported-events).</span></span>

<span data-ttu-id="96ec8-140">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="96ec8-140">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="96ec8-141">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="96ec8-141">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="96ec8-142">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="96ec8-142">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="96ec8-143">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="96ec8-143">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="96ec8-144">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="96ec8-144">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="96ec8-145">Тема Office</span><span class="sxs-lookup"><span data-stu-id="96ec8-145">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="96ec8-146">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="96ec8-146">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="96ec8-147">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="96ec8-147">Added ability to get Office theme.</span></span>

<span data-ttu-id="96ec8-148">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="96ec8-148">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="96ec8-149">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="96ec8-149">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="96ec8-150">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="96ec8-150">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="96ec8-151">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="96ec8-151">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="96ec8-152">Данные сеансов</span><span class="sxs-lookup"><span data-stu-id="96ec8-152">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="96ec8-153">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="96ec8-153">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="96ec8-154">Добавлен новый объект, который представляет данные сеанса элемента.</span><span class="sxs-lookup"><span data-stu-id="96ec8-154">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="96ec8-155">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="96ec8-155">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="96ec8-156">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="96ec8-156">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="96ec8-157">Добавлено новое свойство для управления данными сеанса элемента в режиме Compose.</span><span class="sxs-lookup"><span data-stu-id="96ec8-157">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="96ec8-158">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="96ec8-158">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="96ec8-159">См. также</span><span class="sxs-lookup"><span data-stu-id="96ec8-159">See also</span></span>

- [<span data-ttu-id="96ec8-160">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="96ec8-160">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="96ec8-161">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="96ec8-161">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="96ec8-162">Начало работы</span><span class="sxs-lookup"><span data-stu-id="96ec8-162">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="96ec8-163">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="96ec8-163">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
