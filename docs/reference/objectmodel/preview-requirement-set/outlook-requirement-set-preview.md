---
title: Outlook набор требований к предварительному просмотру API надстройки
description: Функции и API, которые в настоящее время находятся в предварительном Outlook надстройки.
ms.date: 05/17/2021
localization_priority: Normal
ms.openlocfilehash: 98bf56c169967ad7c994d1793afa8678d31f6892
ms.sourcegitcommit: 0d9fcdc2aeb160ff475fbe817425279267c7ff31
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/21/2021
ms.locfileid: "52591060"
---
# <a name="outlook-add-in-api-preview-requirement-set"></a><span data-ttu-id="efadd-103">Outlook набор требований к предварительному просмотру API надстройки</span><span class="sxs-lookup"><span data-stu-id="efadd-103">Outlook add-in API preview requirement set</span></span>

<span data-ttu-id="efadd-104">Подмножество API Outlook надстройки aPI Office JavaScript включает объекты, методы, свойства и события, которые можно использовать в Outlook надстройки.</span><span class="sxs-lookup"><span data-stu-id="efadd-104">The Outlook add-in API subset of the Office JavaScript API includes objects, methods, properties, and events that you can use in an Outlook add-in.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="efadd-105">Эта документация относится к **предварительной версии** [набора обязательных элементов](../../requirement-sets/outlook-api-requirement-sets.md).</span><span class="sxs-lookup"><span data-stu-id="efadd-105">This documentation is for a **preview** [requirement set](../../requirement-sets/outlook-api-requirement-sets.md).</span></span> <span data-ttu-id="efadd-106">Этот набор обязательных элементов еще не полностью реализован, а клиенты будут неправильно сообщать о его поддержке.</span><span class="sxs-lookup"><span data-stu-id="efadd-106">This requirement set is not fully implemented yet, and clients will not accurately report support for it.</span></span> <span data-ttu-id="efadd-107">Не следует указывать этот набор обязательных элементов в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="efadd-107">You should not specify this requirement set in your add-in manifest.</span></span>

[!INCLUDE [Information about using preview APIs](../../../includes/using-preview-apis-host.md)]

> [!TIP]
> <span data-ttu-id="efadd-108">Вы можете просматривать функции Outlook в Интернете, настроив целевой выпуск на Microsoft 365 [клиента.](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center)</span><span class="sxs-lookup"><span data-stu-id="efadd-108">You may be able to preview features in Outlook on the web by [configuring targeted release on your Microsoft 365 tenant](/microsoft-365/admin/manage/release-options-in-office-365?view=o365-worldwide&preserve-view=true#set-up-the-release-option-in-the-admin-center).</span></span> <span data-ttu-id="efadd-109">"Настройка доступа к предварительному просмотру" отмечена на этой странице для применимых функций.</span><span class="sxs-lookup"><span data-stu-id="efadd-109">"Configure preview access" is noted on this page for applicable features.</span></span>
>
> <span data-ttu-id="efadd-110">Для других функций вы можете запросить доступ к битам предварительного просмотра для Outlook веб-страницы с помощью Microsoft 365 учетной записи, заполнив и подав эту [форму.](https://aka.ms/OWAPreview)</span><span class="sxs-lookup"><span data-stu-id="efadd-110">For other features, you may be able to request access to preview bits for Outlook on the web using your Microsoft 365 account by completing and submitting [this form](https://aka.ms/OWAPreview).</span></span> <span data-ttu-id="efadd-111">В этих функциях отмечен "Запрос доступа к предварительному просмотру".</span><span class="sxs-lookup"><span data-stu-id="efadd-111">"Request preview access" is noted on those features.</span></span>

<span data-ttu-id="efadd-112">Набор требований предварительного просмотра включает все функции [набора требований 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span><span class="sxs-lookup"><span data-stu-id="efadd-112">The preview requirement set includes all of the features of [requirement set 1.10](../requirement-set-1.10/outlook-requirement-set-1.10.md).</span></span>

## <a name="features-in-preview"></a><span data-ttu-id="efadd-113">Возможности предварительной версии</span><span class="sxs-lookup"><span data-stu-id="efadd-113">Features in preview</span></span>

<span data-ttu-id="efadd-114">Ниже перечислены возможности предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="efadd-114">The following features are in preview.</span></span>

### <a name="add-in-activation-on-items-protected-by-information-rights-management-irm"></a><span data-ttu-id="efadd-115">Активация надстройки для элементов, защищенных управлением правами на информацию (IRM)</span><span class="sxs-lookup"><span data-stu-id="efadd-115">Add-in activation on items protected by Information Rights Management (IRM)</span></span>

<span data-ttu-id="efadd-116">Надстройки теперь могут активироваться в пунктах, защищенных IRM.</span><span class="sxs-lookup"><span data-stu-id="efadd-116">Add-ins can now activate on IRM-protected items.</span></span> <span data-ttu-id="efadd-117">Чтобы включить эту возможность, администратору клиента необходимо включить право использования, установив в Office параметр Разрешить программный `OBJMODEL` доступ. </span><span class="sxs-lookup"><span data-stu-id="efadd-117">To turn on this capability, a tenant administrator needs to enable the `OBJMODEL` usage right by setting the **Allow programmatic access** custom policy option in Office.</span></span> <span data-ttu-id="efadd-118">Дополнительные [сведения см. в](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) дополнительных сведениях о правах и описаниях использования.</span><span class="sxs-lookup"><span data-stu-id="efadd-118">See [Usage rights and descriptions](/azure/information-protection/configure-usage-rights#usage-rights-and-descriptions) for more information.</span></span>

<span data-ttu-id="efadd-119">**Доступно в**: Outlook на Windows, начиная со сборки 13229.10000 (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="efadd-119">**Available in**: Outlook on Windows, starting with build 13229.10000 (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="additional-calendar-properties"></a><span data-ttu-id="efadd-120">Дополнительные свойства календаря</span><span class="sxs-lookup"><span data-stu-id="efadd-120">Additional calendar properties</span></span>

#### <a name="isalldayevent"></a>[<span data-ttu-id="efadd-121">IsAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="efadd-121">IsAllDayEvent</span></span>](/javascript/api/outlook/office.isalldayevent?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="efadd-122">Добавлен новый объект, который представляет свойство события на весь день встречи в режиме Compose.</span><span class="sxs-lookup"><span data-stu-id="efadd-122">Added a new object that represents the all-day event property of an appointment in Compose mode.</span></span>

<span data-ttu-id="efadd-123">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="efadd-123">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="sensitivity"></a>[<span data-ttu-id="efadd-124">Sensitivity</span><span class="sxs-lookup"><span data-stu-id="efadd-124">Sensitivity</span></span>](/javascript/api/outlook/office.sensitivity?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="efadd-125">Добавлен новый объект, который представляет чувствительность встречи в режиме Compose.</span><span class="sxs-lookup"><span data-stu-id="efadd-125">Added a new object that represents the sensitivity of an appointment in Compose mode.</span></span>

<span data-ttu-id="efadd-126">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="efadd-126">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemisalldayevent"></a>[<span data-ttu-id="efadd-127">Office.context.mailbox.item.isAllDayEvent</span><span class="sxs-lookup"><span data-stu-id="efadd-127">Office.context.mailbox.item.isAllDayEvent</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="efadd-128">Добавлено новое свойство, которое представляет, если встреча является событием на весь день.</span><span class="sxs-lookup"><span data-stu-id="efadd-128">Added a new property that represents if an appointment is an all-day event.</span></span>

<span data-ttu-id="efadd-129">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="efadd-129">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officecontextmailboxitemsensitivity"></a>[<span data-ttu-id="efadd-130">Office.context.mailbox.item.sensitivity</span><span class="sxs-lookup"><span data-stu-id="efadd-130">Office.context.mailbox.item.sensitivity</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="efadd-131">Добавлено новое свойство, которое представляет чувствительность встречи.</span><span class="sxs-lookup"><span data-stu-id="efadd-131">Added a new property that represents the sensitivity of an appointment.</span></span>

<span data-ttu-id="efadd-132">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="efadd-132">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officemailboxenumsappointmentsensitivitytype"></a>[<span data-ttu-id="efadd-133">Office. MailboxEnums.AppointmentSensitivityType</span><span class="sxs-lookup"><span data-stu-id="efadd-133">Office.MailboxEnums.AppointmentSensitivityType</span></span>](/javascript/api/outlook/office.mailboxenums.appointmentsensitivitytype?view=outlook-js-preview&preserve-view=true)

<span data-ttu-id="efadd-134">Добавлен новый `AppointmentSensitivityType` переумыв, который представляет параметры чувствительности, доступные при встрече.</span><span class="sxs-lookup"><span data-stu-id="efadd-134">Added a new enum `AppointmentSensitivityType` that represents the sensitivity options available on an appointment.</span></span>

<span data-ttu-id="efadd-135">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="efadd-135">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="integration-with-actionable-messages"></a><span data-ttu-id="efadd-136">Взаимодействие с интерактивными сообщениями</span><span class="sxs-lookup"><span data-stu-id="efadd-136">Integration with actionable messages</span></span>

#### <a name="officecontextmailboxitemgetinitializationcontextasync"></a>[<span data-ttu-id="efadd-137">Office.context.mailbox.item.getInitializationContextAsync</span><span class="sxs-lookup"><span data-stu-id="efadd-137">Office.context.mailbox.item.getInitializationContextAsync</span></span>](office.context.mailbox.item.md#methods)

<span data-ttu-id="efadd-138">Добавлена новая функция, которая возвращает данные инициализации, передаваемые при [активации надстройки интерактивным сообщением](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span><span class="sxs-lookup"><span data-stu-id="efadd-138">Added a new function that returns initialization data passed when the add-in is [activated by an actionable message](/outlook/actionable-messages/invoke-add-in-from-actionable-message).</span></span>

<span data-ttu-id="efadd-139">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="efadd-139">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

<br>

---

---

### <a name="office-theme"></a><span data-ttu-id="efadd-140">Тема Office</span><span class="sxs-lookup"><span data-stu-id="efadd-140">Office theme</span></span>

#### <a name="officecontextofficetheme"></a>[<span data-ttu-id="efadd-141">Office.context.officeTheme</span><span class="sxs-lookup"><span data-stu-id="efadd-141">Office.context.officeTheme</span></span>](/javascript/api/office/office.context#officetheme)

<span data-ttu-id="efadd-142">Добавлена возможность получения темы Office.</span><span class="sxs-lookup"><span data-stu-id="efadd-142">Added ability to get Office theme.</span></span>

<span data-ttu-id="efadd-143">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="efadd-143">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

#### <a name="officeeventtypeofficethemechanged"></a>[<span data-ttu-id="efadd-144">Office.EventType.OfficeThemeChanged</span><span class="sxs-lookup"><span data-stu-id="efadd-144">Office.EventType.OfficeThemeChanged</span></span>](/javascript/api/office/office.eventtype)

<span data-ttu-id="efadd-145">Добавлено событие `OfficeThemeChanged` для объекта `Mailbox`.</span><span class="sxs-lookup"><span data-stu-id="efadd-145">Added `OfficeThemeChanged` event to `Mailbox`.</span></span>

<span data-ttu-id="efadd-146">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365)</span><span class="sxs-lookup"><span data-stu-id="efadd-146">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription)</span></span>

<br>

---

---

### <a name="session-data"></a><span data-ttu-id="efadd-147">Данные сеансов</span><span class="sxs-lookup"><span data-stu-id="efadd-147">Session data</span></span>

#### <a name="officesessiondata"></a>[<span data-ttu-id="efadd-148">Office. SessionData</span><span class="sxs-lookup"><span data-stu-id="efadd-148">Office.SessionData</span></span>](/javascript/api/outlook/office.sessiondata)

<span data-ttu-id="efadd-149">Добавлен новый объект, который представляет данные сеанса элемента.</span><span class="sxs-lookup"><span data-stu-id="efadd-149">Added a new object that represents the session data of an item.</span></span>

<span data-ttu-id="efadd-150">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="efadd-150">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

#### <a name="officecontextmailboxitemsessiondata"></a>[<span data-ttu-id="efadd-151">Office.context.mailbox.item.sessionData</span><span class="sxs-lookup"><span data-stu-id="efadd-151">Office.context.mailbox.item.sessionData</span></span>](office.context.mailbox.item.md#properties)

<span data-ttu-id="efadd-152">Добавлено новое свойство для управления данными сеанса элемента в режиме Compose.</span><span class="sxs-lookup"><span data-stu-id="efadd-152">Added a new property to manage the session data of an item in Compose mode.</span></span>

<span data-ttu-id="efadd-153">**Доступно в**: Outlook на Windows (подключен к подписке Microsoft 365), Outlook в Интернете (современная)</span><span class="sxs-lookup"><span data-stu-id="efadd-153">**Available in**: Outlook on Windows (connected to a Microsoft 365 subscription), Outlook on the web (modern)</span></span>

## <a name="see-also"></a><span data-ttu-id="efadd-154">См. также</span><span class="sxs-lookup"><span data-stu-id="efadd-154">See also</span></span>

- [<span data-ttu-id="efadd-155">Надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="efadd-155">Outlook add-ins</span></span>](../../../outlook/outlook-add-ins-overview.md)
- [<span data-ttu-id="efadd-156">Примеры кода надстройки Outlook</span><span class="sxs-lookup"><span data-stu-id="efadd-156">Outlook add-in code samples</span></span>](https://developer.microsoft.com/outlook/gallery/?filterBy=Outlook,Samples,Add-ins)
- [<span data-ttu-id="efadd-157">Начало работы</span><span class="sxs-lookup"><span data-stu-id="efadd-157">Get started</span></span>](../../../quickstarts/outlook-quickstart.md)
- [<span data-ttu-id="efadd-158">Наборы обязательных элементов и поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="efadd-158">Requirement sets and supported clients</span></span>](../../requirement-sets/outlook-api-requirement-sets.md)
