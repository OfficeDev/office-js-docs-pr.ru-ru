---
title: Обзор надстроек Outlook
description: Надстройки Outlook — это встраиваемые в Outlook решения от сторонних разработчиков, использующие нашу веб-платформу.
ms.date: 06/15/2021
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: f0c1dbdd1cf9909310b629188d4f3d3d5de6b6bb
ms.sourcegitcommit: 0bf0e076f705af29193abe3dba98cbfcce17b24f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/18/2021
ms.locfileid: "53007813"
---
# <a name="outlook-add-ins-overview"></a><span data-ttu-id="3a2f4-103">Обзор надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="3a2f4-103">Outlook add-ins overview</span></span>

<span data-ttu-id="3a2f4-p101">Надстройки Outlook — это встраиваемые в Outlook решения сторонних разработчиков, использующие нашу веб-платформу. Три ключевых аспекта надстроек Outlook:</span><span class="sxs-lookup"><span data-stu-id="3a2f4-p101">Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform. Outlook add-ins have three key aspects:</span></span>

- <span data-ttu-id="3a2f4-106">В классических приложениях (Outlook для Windows и Mac), веб-приложениях (Microsoft 365 и Outlook.com) и мобильных решениях используются одинаковые логика надстроек и бизнес-логика.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-106">The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Microsoft 365 and Outlook.com), and mobile.</span></span>
- <span data-ttu-id="3a2f4-107">Надстройка Outlook состоит из манифеста, в котором описан способ интеграции надстройки с Outlook (например, при помощи кнопки или области задач), и кода JavaScript или HTML, который составляет пользовательский интерфейс и бизнес-логику надстройки.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-107">Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.</span></span>
- <span data-ttu-id="3a2f4-108">Пользователи и администраторы могут получать надстройки Outlook из [AppSource](https://appsource.microsoft.com) или [загружать их в неопубликованном виде](sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="3a2f4-108">Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.</span></span>

<span data-ttu-id="3a2f4-p102">Надстройки Outlook отличаются от надстроек COM и VSTO, которые предназначены исключительно для Outlook в Windows. В отличие от надстроек COM, надстройки Outlook не устанавливают никакой код непосредственно на устройство пользователя или его клиент Outlook. В случае надстройки Outlook ее клиент считывает манифест и подключает указанные элементы управления в пользовательском интерфейсе, а затем считывает код JavaScript и HTML. Эти веб-компоненты функционируют в контексте изолированного браузера.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-p102">Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows. Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client. For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML. The web components all run in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="3a2f4-p103">Элементы Outlook, поддерживающие надстройки, включают письма, приглашения на собрание, ответы и данные об отменах, а также сведения о встречах. Каждая надстройка Outlook определяет контекст, в котором она доступна, в том числе типы элементов и то, просматривает ли пользователь элемент или создает его.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-p103">The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments. Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.</span></span>

[!INCLUDE [publish policies note](../includes/note-publish-policies.md)]

## <a name="extension-points"></a><span data-ttu-id="3a2f4-115">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="3a2f4-115">Extension points</span></span>

<span data-ttu-id="3a2f4-p104">Надстройка использует точки расширения для интеграции с Outlook. Это можно сделать следующими способами:</span><span class="sxs-lookup"><span data-stu-id="3a2f4-p104">Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:</span></span>

- <span data-ttu-id="3a2f4-p105">Надстройки могут объявлять кнопки, которые отображаются на панелях команд в сообщениях и встречах. Дополнительные сведения см. в статье [Команды надстроек Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="3a2f4-p105">Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

    <span data-ttu-id="3a2f4-120">**Надстройка с кнопками на ленте**</span><span class="sxs-lookup"><span data-stu-id="3a2f4-120">**An add-in with command buttons on the ribbon**</span></span>

    ![Команда надстройки для фигуры без интерфейса](../images/uiless-command-shape.png)

- <span data-ttu-id="3a2f4-p106">Надстройки могут активироваться по совпадениям с регулярными выражениями или обнаруженным сущностям в сообщениях и встречах. Дополнительные сведения см. в статье [Контекстно-зависимые надстройки Outlook](contextual-outlook-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="3a2f4-p106">Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).</span></span>

    <span data-ttu-id="3a2f4-124">**Контекстная надстройка для выделенной сущности (адреса)**</span><span class="sxs-lookup"><span data-stu-id="3a2f4-124">**A contextual add-in for a highlighted entity (an address)**</span></span>

    ![Показывает контекстное приложение на карте](../images/outlook-detected-entity-card.png)

## <a name="mailbox-items-available-to-add-ins"></a><span data-ttu-id="3a2f4-126">Элементы почтовых ящиков, доступные надстройкам</span><span class="sxs-lookup"><span data-stu-id="3a2f4-126">Mailbox items available to add-ins</span></span>

<span data-ttu-id="3a2f4-127">Надстройки Outlook активизируются при создании или чтении сообщения либо встречи, но не других типов элементов.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-127">Outlook add-ins activate when the user is composing or reading a message or appointment, but not other item types.</span></span> <span data-ttu-id="3a2f4-128">При этом надстройки *не* активизируются, если текущий элемент сообщения в форме создания или просмотра имеет одну из следующих особенностей:</span><span class="sxs-lookup"><span data-stu-id="3a2f4-128">However, add-ins are *not* activated if the current message item, in a compose or read form, is one of the following:</span></span>

- <span data-ttu-id="3a2f4-p108">Защищен управлением правами на доступ к данным (IRM) или зашифрован каким-либо другим способом. Один из примеров — сообщение, подписанное цифровой подписью, так как в этом случае используется один из указанных выше механизмов.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-p108">Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.</span></span>

  > [!IMPORTANT]
  >
  > - <span data-ttu-id="3a2f4-131">Надстройки активируют сообщения с цифровой подписью в Outlook, связанном с подпиской на Microsoft 365.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-131">Add-ins activate on digitally signed messages in Outlook associated with a Microsoft 365 subscription.</span></span> <span data-ttu-id="3a2f4-132">В Windows эта поддержка представлена в сборке 8711.1000.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-132">On Windows, this support was introduced with build 8711.1000.</span></span>
  >
  > - <span data-ttu-id="3a2f4-133">Начиная с Outlook сборки 13229.10000 в Windows, надстройки могут активировать элементы, защищенные с помощью IRM.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-133">Starting with Outlook build 13229.10000 on Windows, add-ins can now activate on items protected by IRM.</span></span> <span data-ttu-id="3a2f4-134">Дополнительные сведения об этой функции в предварительной версии см. в статье [Активация надстроек для элементов, защищенных службами управления правами на доступ к данным (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span><span class="sxs-lookup"><span data-stu-id="3a2f4-134">For more information about this feature in preview, refer to [Add-in activation on items protected by Information Rights Management (IRM)](../reference/objectmodel/preview-requirement-set/outlook-requirement-set-preview.md#add-in-activation-on-items-protected-by-information-rights-management-irm).</span></span>

- <span data-ttu-id="3a2f4-135">Отчет или уведомление о доставке имеет класс сообщения IPM.Report.\*, включая отчеты о доставке, о недоставке, а также уведомления о прочтении, о непрочтении и о задержке.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-135">A delivery report or notification that has the message class IPM.Report.\*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.</span></span>

- <span data-ttu-id="3a2f4-136">MSG- или EML-файл, представляющий собой вложение в другое сообщение.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-136">A .msg or .eml file which is an attachment to another message.</span></span>

- <span data-ttu-id="3a2f4-137">MSG- или EML-файл, открытый из файловой системы.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-137">A .msg or .eml file opened from the file system.</span></span>

- <span data-ttu-id="3a2f4-138">В [групповом почтовом ящике](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes), общем почтовом ящике\*, почтовом ящике другого пользователя \*, архивном почтовом ящике или общедоступной папке.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-138">In a [group mailbox](/microsoft-365/admin/create-groups/compare-groups?view=o365-worldwide&preserve-view=true#shared-mailboxes), in a shared mailbox\*, in another user's mailbox\*, in an archive mailbox, or in a public folder.</span></span>

  > [!IMPORTANT]
  > <span data-ttu-id="3a2f4-139">\* Поддержка сценариев делегирования доступа (например, папок, полученных из почтового ящика другого пользователя) была представлена в [наборе требований 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span><span class="sxs-lookup"><span data-stu-id="3a2f4-139">\* Support for delegate access scenarios (for example, folders shared from another user's mailbox) was introduced in [requirement set 1.8](../reference/objectmodel/requirement-set-1.8/outlook-requirement-set-1.8.md).</span></span> <span data-ttu-id="3a2f4-140">Поддержка общих почтовых ящиков теперь доступна в предварительной версии.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-140">Shared mailbox support is now in preview.</span></span> <span data-ttu-id="3a2f4-141">Дополнительные сведения приводятся в статье [Включение сценариев общих папок и общих почтовых ящиков](delegate-access.md).</span><span class="sxs-lookup"><span data-stu-id="3a2f4-141">To learn more, refer to [Enable shared folders and shared mailbox scenarios](delegate-access.md).</span></span>

- <span data-ttu-id="3a2f4-142">Использование настраиваемой формы.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-142">Using a custom form.</span></span>

<span data-ttu-id="3a2f4-143">В общем случае Outlook может активировать надстройки в формах просмотра для элементов в папке "Отправленные", за исключением надстроек, активируемых на основании совпадений строк для известных сущностей.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-143">In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities.</span></span> <span data-ttu-id="3a2f4-144">Дополнительные сведения о причинах этого см. "Поддержка известных сущностей" в статье [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="3a2f4-144">For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>

## <a name="supported-clients"></a><span data-ttu-id="3a2f4-145">Поддерживаемые клиенты</span><span class="sxs-lookup"><span data-stu-id="3a2f4-145">Supported clients</span></span>

<span data-ttu-id="3a2f4-146">Надстройки Outlook поддерживают Outlook 2013 или более поздней версии для Windows, Outlook 2016 или более поздней версии для Mac, Outlook в Интернете для локальной версии Exchange 2013 и более поздних версий, Outlook для iOS, Outlook для Android, Outlook в Интернете и Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-146">Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web and Outlook.com.</span></span> <span data-ttu-id="3a2f4-147">Не все новые функции поддерживаются сразу всеми [клиентами](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="3a2f4-147">Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time.</span></span> <span data-ttu-id="3a2f4-148">Просмотрите статьи и справочные материалы по API для этих функций, чтобы узнать, в каких приложениях они поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-148">Please refer to articles and API references for those features to see which applications they may or may not be supported in.</span></span>

## <a name="get-started-building-outlook-add-ins"></a><span data-ttu-id="3a2f4-149">Знакомство с разработкой надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="3a2f4-149">Get started building Outlook add-ins</span></span>

<span data-ttu-id="3a2f4-150">Чтобы приступить к разработке надстроек Outlook, попробуйте приведенные ниже ресурсы.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-150">To get started building Outlook add-ins, try the following:</span></span>

- <span data-ttu-id="3a2f4-151">[Краткое руководство](../quickstarts/outlook-quickstart.md) — создание простой надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-151">[Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.</span></span>
- <span data-ttu-id="3a2f4-152">[Учебник](../tutorials/outlook-tutorial.md) — узнайте, как создать надстройку, которая вставляет элементы gist с сайта GitHub в новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="3a2f4-152">[Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.</span></span>

## <a name="see-also"></a><span data-ttu-id="3a2f4-153">См. также</span><span class="sxs-lookup"><span data-stu-id="3a2f4-153">See also</span></span>

- [<span data-ttu-id="3a2f4-154">Сведения о программе для разработчиков Microsoft 365</span><span class="sxs-lookup"><span data-stu-id="3a2f4-154">Learn about the Microsoft 365 Developer Program</span></span>](https://developer.microsoft.com/microsoft-365/dev-program)
- [<span data-ttu-id="3a2f4-155">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="3a2f4-155">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="3a2f4-156">Рекомендации по проектированию надстроек Office</span><span class="sxs-lookup"><span data-stu-id="3a2f4-156">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="3a2f4-157">Лицензирование надстроек Office и SharePoint</span><span class="sxs-lookup"><span data-stu-id="3a2f4-157">License your Office and SharePoint Add-ins</span></span>](/office/dev/store/license-your-add-ins)
- [<span data-ttu-id="3a2f4-158">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="3a2f4-158">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="3a2f4-159">Публикация решений в AppSource и в Office</span><span class="sxs-lookup"><span data-stu-id="3a2f4-159">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
