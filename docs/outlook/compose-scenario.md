---
title: Разработка надстроек Outlook для форм создания
description: Узнайте о сценариях и возможностях надстроек Outlook для форм создания.
ms.date: 04/15/2019
localization_priority: Priority
ms.openlocfilehash: ea9a9bb245e8912111cad154a0cc66b0288d6eaa
ms.sourcegitcommit: a3ddfdb8a95477850148c4177e20e56a8673517c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/20/2020
ms.locfileid: "42166780"
---
# <a name="create-outlook-add-ins-for-compose-forms"></a><span data-ttu-id="9c3d8-103">Разработка надстроек Outlook для форм создания</span><span class="sxs-lookup"><span data-stu-id="9c3d8-103">Create Outlook add-ins for compose forms</span></span>

<span data-ttu-id="9c3d8-104">Начиная с версии 1.1 схемы манифестов для надстроек Office и версии 1.1 библиотеки Office.js, вы можете разрабатывать надстройки создания — надстройки Outlook, активируемые в формах создания.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-104">Starting with version 1.1 of the schema for Office Add-ins manifests and v1.1 of Office.js, you can create compose add-ins, which are Outlook add-ins activated in compose forms.</span></span> <span data-ttu-id="9c3d8-105">В отличие от надстроек чтения (активируемых в режиме чтения, когда пользователь просматривает сообщение или встречу), надстройки создания доступны в следующих сценариях:</span><span class="sxs-lookup"><span data-stu-id="9c3d8-105">In contrast with read add-ins (Outlook add-ins that are activated in read mode when a user is viewing a message or appointment), compose add-ins are available in the following user scenarios:</span></span>

- <span data-ttu-id="9c3d8-106">Создание сообщения, приглашения на собрание или встречи в отдельной форме.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-106">Composing a new message, meeting request, or appointment in a compose form.</span></span>

- <span data-ttu-id="9c3d8-107">Просмотр или редактирование существующих встречи или собрания, организованных пользователем.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-107">Viewing or editing an existing appointment, or meeting item in which the user is the organizer.</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="9c3d8-108">При просмотре организованной пользователем встречи в Outlook 2013 RTM или Exchange 2013 RTM доступны надстройки чтения.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-108">If the user is on the RTM release of Outlook 2013 and Exchange 2013 and is viewing a meeting item organized by the user, the user can find read add-ins available.</span></span> <span data-ttu-id="9c3d8-109">Начиная с выпуска Office 2013 с пакетом обновления 1 (SP1), только надстройки создания могут активироваться и быть доступными.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-109">Starting in the Office 2013 SP1 release, there's a change such that in the same scenario, only compose add-ins can activate and be available.</span></span>

- <span data-ttu-id="9c3d8-110">Создание ответа на сообщение (встроенного или в отдельной форме).</span><span class="sxs-lookup"><span data-stu-id="9c3d8-110">Composing an inline response message or replying to a message in a separate compose form.</span></span>

- <span data-ttu-id="9c3d8-111">Изменение ответа (**Принять**, **Под вопросом** или **Отклонить**) на приглашение на собрание или элемент собрания.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-111">Editing a response (**Accept**, **Tentative**, or **Decline**) to a meeting request or meeting item.</span></span>

- <span data-ttu-id="9c3d8-112">Предложение нового времени для элемента собрания.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-112">Proposing a new time for a meeting item.</span></span>

- <span data-ttu-id="9c3d8-113">Пересылка или ответ на приглашение на собрание или элемент собрания.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-113">Forwarding or replying to a meeting request or meeting item.</span></span>

<span data-ttu-id="9c3d8-114">В каждом из этих сценариев отображаются все определенные кнопки команд надстройки.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-114">In each of these compose scenarios, any add-in command buttons defined by the add-in are shown.</span></span> <span data-ttu-id="9c3d8-115">В старых надстройках, где не реализованы команды, пользователи могут выбрать **Надстройки Office** на ленте, чтобы открыть область выбора надстроек, а затем выбрать и запустить надстройку создания.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-115">For older add-ins that do not implement add-in commands, users can choose **Office Add-ins** in the ribbon to open the add-in selection pane, and then choose and start a compose add-in.</span></span> <span data-ttu-id="9c3d8-116">Ниже показаны команды надстройки в форме создания.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-116">The following figure shows add-in commands in a compose form.</span></span>

![Форма создания элемента Outlook с командами надстройки](../images/compose-form-commands.png)

<span data-ttu-id="9c3d8-118">На рисунке ниже показана область выбора надстроек, включающая две надстройки создания, в которых не реализованы команды. Она активируется при создании встроенного ответа в Outlook.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-118">The following figure shows the add-in selection pane consisting of two compose add-ins that do not implement add-in commands, activated when the user is composing an inline reply in Outlook.</span></span>

![Почтовое приложение, содержащее шаблоны, которое активировано в форме создания](../images/templates-app-selection.png)

## <a name="types-of-add-ins-available-in-compose-mode"></a><span data-ttu-id="9c3d8-120">Типы надстроек, доступные в режиме создания</span><span class="sxs-lookup"><span data-stu-id="9c3d8-120">Types of add-ins available in compose mode</span></span>

<span data-ttu-id="9c3d8-121">Надстройки создания реализуются в виде [команд надстроек Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="9c3d8-121">Compose add-ins are implemented as [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span> <span data-ttu-id="9c3d8-122">Чтобы надстройки активировались при создании писем или ответов на приглашения на собрания, в манифест включается [точка расширения MessageComposeCommandSurface](../reference/manifest/extensionpoint.md#messagecomposecommandsurface).</span><span class="sxs-lookup"><span data-stu-id="9c3d8-122">To activate add-ins for composing email or meeting responses, add-ins include a [MessageComposeCommandSurface extension point element](../reference/manifest/extensionpoint.md#messagecomposecommandsurface) in the manifest.</span></span> <span data-ttu-id="9c3d8-123">Чтобы надстройки активировались при создании или редактировании встреч или собраний, организованных пользователем, добавляется [точка расширения AppointmentOrganizerCommandSurface](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface).</span><span class="sxs-lookup"><span data-stu-id="9c3d8-123">To activate add-ins for composing or editing appointments or meetings where the user is the organizer, add-ins include a [AppointmentOrganizerCommandSurface extension point element](../reference/manifest/extensionpoint.md#appointmentorganizercommandsurface).</span></span>

> [!NOTE]
> <span data-ttu-id="9c3d8-124">На серверах или клиентах, не поддерживающих команды надстроек, используются [правила активации](activation-rules.md), указанные в элементе [Rule](../reference/manifest/rule.md), содержащемся в элементе [OfficeApp](../reference/manifest/officeapp.md).</span><span class="sxs-lookup"><span data-stu-id="9c3d8-124">Add-ins developed for servers or clients that do not support add-in commands use [activation rules](activation-rules.md) in a [Rule](../reference/manifest/rule.md) element contained in the [OfficeApp](../reference/manifest/officeapp.md) element.</span></span> <span data-ttu-id="9c3d8-125">Если надстройка не разрабатывается специально для устаревших клиентов и серверов, в ней следует использовать команды надстроек.</span><span class="sxs-lookup"><span data-stu-id="9c3d8-125">Unless the add-in is being specifically developed for older clients and servers, new add-ins should use add-in commands.</span></span>

## <a name="api-features-available-to-compose-add-ins"></a><span data-ttu-id="9c3d8-126">Функции API, доступные надстройкам создания</span><span class="sxs-lookup"><span data-stu-id="9c3d8-126">API features available to compose add-ins</span></span>

- [<span data-ttu-id="9c3d8-127">Добавление и удаление вложений в форме создания Outlook</span><span class="sxs-lookup"><span data-stu-id="9c3d8-127">Add and remove attachments to an item in a compose form in Outlook</span></span>](add-and-remove-attachments-to-an-item-in-a-compose-form.md)
- [<span data-ttu-id="9c3d8-128">Просмотр и изменение данных элемента в форме создания элементов Outlook</span><span class="sxs-lookup"><span data-stu-id="9c3d8-128">Get and set item data in a compose form in Outlook</span></span>](get-and-set-item-data-in-a-compose-form.md)
- [<span data-ttu-id="9c3d8-129">Просмотр, изменение или добавление получателей при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="9c3d8-129">Get, set, or add recipients when composing an appointment or message in Outlook</span></span>](get-set-or-add-recipients.md)
- [<span data-ttu-id="9c3d8-130">Просмотр или изменение темы при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="9c3d8-130">Get or set the subject when composing an appointment or message in Outlook</span></span>](get-or-set-the-subject.md)
- [<span data-ttu-id="9c3d8-131">Вставка данных в текст при создании встречи или сообщения в Outlook</span><span class="sxs-lookup"><span data-stu-id="9c3d8-131">Insert data in the body when composing an appointment or message in Outlook</span></span>](insert-data-in-the-body.md)
- [<span data-ttu-id="9c3d8-132">Просмотр или изменение расположения при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="9c3d8-132">Get or set the location when composing an appointment in Outlook</span></span>](get-or-set-the-location-of-an-appointment.md)
- [<span data-ttu-id="9c3d8-133">Просмотр или изменение времени при создании встречи в Outlook</span><span class="sxs-lookup"><span data-stu-id="9c3d8-133">Get or set the time when composing an appointment in Outlook</span></span>](get-or-set-the-time-of-an-appointment.md)

## <a name="see-also"></a><span data-ttu-id="9c3d8-134">См. также</span><span class="sxs-lookup"><span data-stu-id="9c3d8-134">See also</span></span>

- [<span data-ttu-id="9c3d8-135">Начало работы с надстройками Outlook для Office 365</span><span class="sxs-lookup"><span data-stu-id="9c3d8-135">Get Started with Outlook add-ins for Office 365</span></span>](../quickstarts/outlook-quickstart.md)
