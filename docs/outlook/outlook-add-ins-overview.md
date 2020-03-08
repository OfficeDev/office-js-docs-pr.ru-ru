---
title: Обзор надстроек Outlook
description: Надстройки Outlook — это встраиваемые в Outlook решения от сторонних разработчиков, использующие нашу веб-платформу.
ms.date: 10/09/2019
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: cb6e19788390a804b0bbacb97666a3ca8a9d5971
ms.sourcegitcommit: 6c7c98f085dd20f827e0c388e672993412944851
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/06/2020
ms.locfileid: "42554699"
---
# <a name="outlook-add-ins-overview"></a><span data-ttu-id="5b00d-103">Обзор надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="5b00d-103">Outlook add-ins overview</span></span>

<span data-ttu-id="5b00d-104">Надстройки Outlook — это встраиваемые в Outlook решения от сторонних разработчиков, использующие нашу веб-платформу.</span><span class="sxs-lookup"><span data-stu-id="5b00d-104">Outlook add-ins are integrations built by third parties into Outlook by using our web-based platform.</span></span> <span data-ttu-id="5b00d-105">Три ключевых аспекта надстроек Outlook:</span><span class="sxs-lookup"><span data-stu-id="5b00d-105">Outlook add-ins have three key aspects:</span></span>

- <span data-ttu-id="5b00d-106">В классических приложениях (Outlook для Windows и Mac), веб-приложениях (Office 365 и Outlook.com) и мобильных решениях используются одинаковые логика надстроек и бизнес-логика.</span><span class="sxs-lookup"><span data-stu-id="5b00d-106">The same add-in and business logic works across desktop (Outlook on Windows and Mac), web (Office 365 and Outlook.com), and mobile.</span></span>
- <span data-ttu-id="5b00d-107">Надстройка Outlook состоит из манифеста, в котором описан способ интеграции надстройки с Outlook (например, при помощи кнопки или области задач), и кода JavaScript или HTML, который составляет пользовательский интерфейс и бизнес-логику надстройки.</span><span class="sxs-lookup"><span data-stu-id="5b00d-107">Outlook add-ins consist of a manifest, which describes how the add-in integrates into Outlook (for example, a button or a task pane), and JavaScript/HTML code, which makes up the UI and business logic of the add-in.</span></span>
- <span data-ttu-id="5b00d-108">Пользователи и администраторы могут получать надстройки Outlook из [AppSource](https://appsource.microsoft.com) или [загружать их в неопубликованном виде](sideload-outlook-add-ins-for-testing.md).</span><span class="sxs-lookup"><span data-stu-id="5b00d-108">Outlook add-ins can be acquired from [AppSource](https://appsource.microsoft.com) or [sideloaded](sideload-outlook-add-ins-for-testing.md) by end-users or administrators.</span></span>

<span data-ttu-id="5b00d-109">Надстройки Outlook отличаются от надстроек COM или VSTO, которые являются более ранними интеграциями, относящимися к Outlook под управлением Windows.</span><span class="sxs-lookup"><span data-stu-id="5b00d-109">Outlook add-ins are different from COM or VSTO add-ins, which are older integrations specific to Outlook running on Windows.</span></span> <span data-ttu-id="5b00d-110">В отличие от надстроек COM надстройки Outlook не имеют какого-либо кода, физически установленного на устройстве пользователя или в клиентах Outlook.</span><span class="sxs-lookup"><span data-stu-id="5b00d-110">Unlike COM add-ins, Outlook add-ins don't have any code physically installed on the user's device or Outlook client.</span></span> <span data-ttu-id="5b00d-111">В случае надстройки Outlook программа Outlook считывает манифест и подключает указанные элементы управления в пользовательском интерфейсе, а затем загружает JavaScript и HTML.</span><span class="sxs-lookup"><span data-stu-id="5b00d-111">For an Outlook add-in, Outlook reads the manifest and hooks up the specified controls in the UI, and then loads the JavaScript and HTML.</span></span> <span data-ttu-id="5b00d-112">Все веб-компоненты выполняется в "песочнице" в контексте браузера.</span><span class="sxs-lookup"><span data-stu-id="5b00d-112">The web components all run in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="5b00d-113">К элементам Outlook, поддерживающим надстройки, относятся письма, приглашения на собрание, ответы и данные об отменах, а также сведения о встречах.</span><span class="sxs-lookup"><span data-stu-id="5b00d-113">The Outlook items that support add-ins include email messages, meeting requests, responses and cancellations, and appointments.</span></span> <span data-ttu-id="5b00d-114">Каждая надстройка Outlook определяет контекст, в котором она доступна, в том числе типы элементов и то, просматривает пользователь элемент или создает его.</span><span class="sxs-lookup"><span data-stu-id="5b00d-114">Each Outlook add-in defines the context in which it is available, including the types of items and if the user is reading or composing an item.</span></span>

> [!NOTE]
> <span data-ttu-id="5b00d-p104">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource, она должна соответствовать [политикам проверки AppSource](/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и [статье о доступности надстроек Office в ведущих приложениях](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="5b00d-p104">When you build your add-in, if you plan to [publish](../publish/publish.md) your add-in to AppSource, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span>

## <a name="extension-points"></a><span data-ttu-id="5b00d-117">Точки расширения</span><span class="sxs-lookup"><span data-stu-id="5b00d-117">Extension points</span></span>

<span data-ttu-id="5b00d-p105">Надстройка использует точки расширения для интеграции с Outlook. Это можно сделать следующими способами:</span><span class="sxs-lookup"><span data-stu-id="5b00d-p105">Extension points are the ways that add-ins integrate with Outlook. The following are the ways this can be done:</span></span>

- <span data-ttu-id="5b00d-p106">Надстройки могут объявлять кнопки, которые отображаются на панелях команд в сообщениях и встречах. Дополнительные сведения см. в статье [Команды надстроек Outlook](add-in-commands-for-outlook.md).</span><span class="sxs-lookup"><span data-stu-id="5b00d-p106">Add-ins can declare buttons that appear in command surfaces across messages and appointments. For more information, see [Add-in commands for Outlook](add-in-commands-for-outlook.md).</span></span>

    <span data-ttu-id="5b00d-122">**Надстройка с кнопками на ленте**</span><span class="sxs-lookup"><span data-stu-id="5b00d-122">**An add-in with command buttons on the ribbon**</span></span>

    ![Команда надстройки для фигуры без интерфейса](../images/uiless-command-shape.png)

- <span data-ttu-id="5b00d-p107">Надстройки могут активироваться по совпадениям с регулярными выражениями или обнаруженным сущностям в сообщениях и встречах. Дополнительные сведения см. в статье [Контекстно-зависимые надстройки Outlook](contextual-outlook-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="5b00d-p107">Add-ins can link off regular expression matches or detected entities in messages and appointments. For more information, see [Contextual Outlook add-ins](contextual-outlook-add-ins.md).</span></span>

    <span data-ttu-id="5b00d-126">**Контекстная надстройка для выделенной сущности (адреса)**</span><span class="sxs-lookup"><span data-stu-id="5b00d-126">**A contextual add-in for a highlighted entity (an address)**</span></span>

    ![Показывает контекстное приложение на карточке](../images/outlook-detected-entity-card.png)


> [!NOTE]
> <span data-ttu-id="5b00d-128">Поскольку [настраиваемые области устарели](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/), следует убедиться, что вы используете поддерживаемую точку расширения.</span><span class="sxs-lookup"><span data-stu-id="5b00d-128">[Custom panes have been deprecated](https://developer.microsoft.com/outlook/blogs/make-your-add-ins-available-in-the-office-ribbon/) so please ensure that you're using a supported extension point.</span></span>

## <a name="mailbox-items-available-to-add-ins"></a><span data-ttu-id="5b00d-129">Элементы почтовых ящиков, доступные надстройкам</span><span class="sxs-lookup"><span data-stu-id="5b00d-129">Mailbox items available to add-ins</span></span>

<span data-ttu-id="5b00d-p108">Надстройки Outlook доступны при создании или просмотре сообщений или встреч. Outlook не активирует надстройки, если текущий элемент сообщения в форме создания или просмотра имеет одну из следующих особенностей:</span><span class="sxs-lookup"><span data-stu-id="5b00d-p108">Outlook add-ins are available on messages or appointments while composing or reading, but not other item types. Outlook does not activate add-ins if the current message item, in a compose or read form, is one of the following:</span></span>

- <span data-ttu-id="5b00d-p109">Защищен управлением правами на доступ к данным (IRM) или зашифрован каким-либо другим способом. Один из примеров — сообщение, подписанное цифровой подписью, так как в этом случае используется один из указанных выше механизмов.</span><span class="sxs-lookup"><span data-stu-id="5b00d-p109">Protected by Information Rights Management (IRM) or encrypted in other ways for protection. A digitally signed message is an example since digital signing relies on one of these mechanisms.</span></span>

- <span data-ttu-id="5b00d-134">Отчет или уведомление о доставке имеет класс сообщения IPM.Report.\*, включая отчеты о доставке, о недоставке, а также уведомления о прочтении, о непрочтении и о задержке.</span><span class="sxs-lookup"><span data-stu-id="5b00d-134">A delivery report or notification that has the message class IPM.Report.\*, including delivery and Non-Delivery Report (NDR) reports, and read, non-read, and delay notifications.</span></span>

- <span data-ttu-id="5b00d-135">Элемент является черновиком (не имеет назначенного отправителя) или находится в папке черновиков Outlook.</span><span class="sxs-lookup"><span data-stu-id="5b00d-135">A draft (does not have a sender assigned to it), or in the Outlook Drafts folder.</span></span>

- <span data-ttu-id="5b00d-136">MSG- или EML-файл, представляющий собой вложение в другое сообщение.</span><span class="sxs-lookup"><span data-stu-id="5b00d-136">A .msg or .eml file which is an attachment to another message.</span></span>

- <span data-ttu-id="5b00d-137">MSG- или EML-файл, открытый из файловой системы.</span><span class="sxs-lookup"><span data-stu-id="5b00d-137">A .msg or .eml file opened from the file system.</span></span>

- <span data-ttu-id="5b00d-138">В общем почтовом ящике, почтовом ящике другого пользователя, архивном почтовом ящике или общедоступной папке.</span><span class="sxs-lookup"><span data-stu-id="5b00d-138">In a shared mailbox, in another user's mailbox, in an archive mailbox, or in a public folder.</span></span>

- <span data-ttu-id="5b00d-139">Использование настраиваемой формы.</span><span class="sxs-lookup"><span data-stu-id="5b00d-139">Using a custom form.</span></span>

<span data-ttu-id="5b00d-140">В общем случае Outlook может активировать надстройки в формах просмотра для элементов в папке "Отправленные", за исключением надстроек, активируемых на основании совпадений строк для известных сущностей.</span><span class="sxs-lookup"><span data-stu-id="5b00d-140">In general, Outlook can activate add-ins in read form for items in the Sent Items folder, with the exception of add-ins that activate based on string matches of well-known entities.</span></span> <span data-ttu-id="5b00d-141">Дополнительные сведения о причинах этого см. "Поддержка известных сущностей" в статье [Сопоставление строк в элементе Outlook как известных сущностей](match-strings-in-an-item-as-well-known-entities.md).</span><span class="sxs-lookup"><span data-stu-id="5b00d-141">For more information about the reasons behind this, see "Support for well-known entities" in [Match strings in an Outlook item as well-known entities](match-strings-in-an-item-as-well-known-entities.md).</span></span>

## <a name="supported-hosts"></a><span data-ttu-id="5b00d-142">Поддерживаемые ведущие приложения</span><span class="sxs-lookup"><span data-stu-id="5b00d-142">Supported hosts</span></span>

<span data-ttu-id="5b00d-143">Надстройки Outlook поддерживают Outlook 2013 или более поздней версии для Windows, Outlook 2016 или более поздней версии для Mac, Outlook в Интернете для локальной версии Exchange 2013 и более поздних версий, Outlook для iOS, Outlook для Android, Outlook в Интернете в Office 365 и Outlook.com.</span><span class="sxs-lookup"><span data-stu-id="5b00d-143">Outlook add-ins are supported in Outlook 2013 or later on Windows, Outlook 2016 or later on Mac, Outlook on the web for Exchange 2013 on-premises and later versions, Outlook on iOS, Outlook on Android, and Outlook on the web in Office 365 and Outlook.com.</span></span> <span data-ttu-id="5b00d-144">Не все новые функции поддерживаются сразу всеми [клиентами](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients).</span><span class="sxs-lookup"><span data-stu-id="5b00d-144">Not all of the newest features are supported in all [clients](../reference/requirement-sets/outlook-api-requirement-sets.md#requirement-sets-supported-by-exchange-servers-and-outlook-clients) at the same time.</span></span> <span data-ttu-id="5b00d-145">Просмотрите статьи и справочные материалы по API для этих функций, чтобы узнать, в каких ведущих приложениях они поддерживаются.</span><span class="sxs-lookup"><span data-stu-id="5b00d-145">Please refer to articles and API references for those features to see which hosts they may or may not be supported in.</span></span>


## <a name="get-started-building-outlook-add-ins"></a><span data-ttu-id="5b00d-146">Знакомство с разработкой надстроек Outlook</span><span class="sxs-lookup"><span data-stu-id="5b00d-146">Get started building Outlook add-ins</span></span>

<span data-ttu-id="5b00d-147">Чтобы приступить к разработке надстроек Outlook, попробуйте приведенные ниже ресурсы.</span><span class="sxs-lookup"><span data-stu-id="5b00d-147">To get started building Outlook add-ins, try the following.</span></span>

- <span data-ttu-id="5b00d-148">[Краткое руководство](../quickstarts/outlook-quickstart.md) — создание простой надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="5b00d-148">[Quick start](../quickstarts/outlook-quickstart.md) - Build a simple task pane.</span></span>
- <span data-ttu-id="5b00d-149">[Учебник](../tutorials/outlook-tutorial.md) — узнайте, как создать надстройку, которая вставляет элементы gist с сайта GitHub в новое сообщение.</span><span class="sxs-lookup"><span data-stu-id="5b00d-149">[Tutorial](../tutorials/outlook-tutorial.md) - Learn how to create an add-in that inserts GitHub gists into a new message.</span></span>


## <a name="see-also"></a><span data-ttu-id="5b00d-150">См. также</span><span class="sxs-lookup"><span data-stu-id="5b00d-150">See also</span></span>

- [<span data-ttu-id="5b00d-151">Рекомендации по разработке надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5b00d-151">Best practices for developing Office Add-ins</span></span>](../concepts/add-in-development-best-practices.md)
- [<span data-ttu-id="5b00d-152">Рекомендации по проектированию надстроек Office</span><span class="sxs-lookup"><span data-stu-id="5b00d-152">Design guidelines for Office Add-ins</span></span>](../design/add-in-design.md)
- [<span data-ttu-id="5b00d-153">Лицензирование надстроек Office и SharePoint</span><span class="sxs-lookup"><span data-stu-id="5b00d-153">License your Office and SharePoint Add-ins</span></span>](/office/dev/store/license-your-add-ins)
- [<span data-ttu-id="5b00d-154">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="5b00d-154">Publish your Office Add-in</span></span>](../publish/publish.md)
- [<span data-ttu-id="5b00d-155">Публикация решений в AppSource и в Office</span><span class="sxs-lookup"><span data-stu-id="5b00d-155">Make your solutions available in AppSource and within Office</span></span>](/office/dev/store/submit-to-the-office-store)
