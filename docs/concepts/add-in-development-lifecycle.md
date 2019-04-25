---
title: Жизненный цикл разработки надстроек Office
description: ''
ms.date: 03/19/2019
localization_priority: Priority
ms.openlocfilehash: 86c384128640d64c47185a290bc224ffe7b59274
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32448366"
---
# <a name="office-add-ins-development-lifecycle"></a><span data-ttu-id="26bb3-102">Жизненный цикл разработки надстроек Office</span><span class="sxs-lookup"><span data-stu-id="26bb3-102">Office Add-ins development lifecycle</span></span>

> [!NOTE]
> <span data-ttu-id="26bb3-p101">Если вы планируете [опубликовать](../publish/publish.md) надстройку в AppSource и сделать ее доступной в интерфейсе Office, убедитесь, что она соответствует [политикам проверки AppSource](/office/dev/store/validation-policies). Например, чтобы пройти проверку, надстройка должна работать на всех платформах, поддерживающих определенные вами методы. Дополнительные сведения см. в [разделе 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) и на [странице со сведениями о доступности и ведущих приложениях для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="26bb3-p101">If you plan to [publish](../publish/publish.md) your add-in to AppSource and make it available within the Office experience, make sure that you conform to the [AppSource validation policies](/office/dev/store/validation-policies). For example, to pass validation, your add-in must work across all platforms that support the methods that you define (for more information, see [section 4.12](/office/dev/store/validation-policies#4-apps-and-add-ins-behave-predictably) and the [Office Add-in host and availability page](../overview/office-add-in-availability.md)).</span></span> 

<span data-ttu-id="26bb3-105">Типичный жизненный цикл разработки надстройки Office состоит из перечисленных ниже этапов.</span><span class="sxs-lookup"><span data-stu-id="26bb3-105">The typical development lifecycle of an Office Add-in includes the following steps:</span></span>


## <a name="1-decide-on-the-purpose-of-the-add-in"></a><span data-ttu-id="26bb3-106">1. Определение назначения надстройки</span><span class="sxs-lookup"><span data-stu-id="26bb3-106">1. Decide on the purpose of the add-in</span></span>

<span data-ttu-id="26bb3-107">Задайте следующие вопросы:</span><span class="sxs-lookup"><span data-stu-id="26bb3-107">Ask the following questions:</span></span>

- <span data-ttu-id="26bb3-108">В чем польза от этой надстройки?</span><span class="sxs-lookup"><span data-stu-id="26bb3-108">How is the add-in useful?</span></span>

- <span data-ttu-id="26bb3-109">Как оно поможет пользователям повысить производительность своего труда?</span><span class="sxs-lookup"><span data-stu-id="26bb3-109">How does it help your customers be more productive?</span></span>

- <span data-ttu-id="26bb3-110">Какие сценарии поддерживают функции вашей надстройки?</span><span class="sxs-lookup"><span data-stu-id="26bb3-110">What scenarios does your add-in's features support?</span></span>

<span data-ttu-id="26bb3-111">Определите наиболее важные возможности и сценарии и сосредоточьтесь на них при разработке надстройки.</span><span class="sxs-lookup"><span data-stu-id="26bb3-111">Decide the most important features and scenarios and focus your design around them.</span></span>


## <a name="2-identify-the-data-and-data-source-for-the-add-in"></a><span data-ttu-id="26bb3-112">2. Определение данных и их источника для надстройки</span><span class="sxs-lookup"><span data-stu-id="26bb3-112">2. Identify the data and data source for the add-in</span></span>

- <span data-ttu-id="26bb3-113">Где находятся данные: в документе, книге, презентации, проекте или браузерной базе данных Access?</span><span class="sxs-lookup"><span data-stu-id="26bb3-113">Is the data in a document, workbook, presentation, project, or an Access browser-based database?</span></span>

- <span data-ttu-id="26bb3-114">Относятся ли данные к одному или нескольким элементам на сервере Exchange Server или в почтовом ящике Exchange Online?</span><span class="sxs-lookup"><span data-stu-id="26bb3-114">Is the data about an item or items in an Exchange Server or Exchange Online mailbox?</span></span>

- <span data-ttu-id="26bb3-115">Данные получены из внешнего источника (например, веб-службы)?</span><span class="sxs-lookup"><span data-stu-id="26bb3-115">Is the data from an external source such as a web service?</span></span>


## <a name="3-identify-the-type-of-add-in-and-office-host-applications-that-best-support-the-purpose-of-the-add-in"></a><span data-ttu-id="26bb3-116">3. Определение типа надстройки и ведущих приложений Office, наиболее подходящих для ее назначения</span><span class="sxs-lookup"><span data-stu-id="26bb3-116">3. Identify the type of add-in and Office host applications that best support the purpose of the add-in</span></span>

<span data-ttu-id="26bb3-117">Определяя сценарии, учитывайте следующее:</span><span class="sxs-lookup"><span data-stu-id="26bb3-117">Consider the following to identify the scenarios:</span></span>

- <span data-ttu-id="26bb3-p102">Будут ли клиенты использовать надстройку для улучшения содержимого документа или браузерной базы данных Access? Если это так, может быть целесообразно создать **контентную надстройку**.</span><span class="sxs-lookup"><span data-stu-id="26bb3-p102">Will customers use the add-in to enrich the content of a document or Access browser-based database? If so, you may want to consider creating a **content add-in**.</span></span>

- <span data-ttu-id="26bb3-p103">Будут ли клиенты использовать надстройку во время просмотра или создания электронного сообщения или встречи? Важна ли возможность отображать надстройку с учетом текущего контекста? Важно ли сделать надстройку доступной не только на настольных компьютерах, но и на планшетах?</span><span class="sxs-lookup"><span data-stu-id="26bb3-p103">Will customers use the add-in while viewing or composing an email message or appointment? Is being able to expose the add-in according to the current context important? Is making the add-in available on not just the desktop, but also on tablets and phones a priority?</span></span>

    <span data-ttu-id="26bb3-p104">Если вы ответили "да" на какой-либо из этих вопросов, рекомендуем создать **надстройку Outlook**. Определите, в каком контексте будет активироваться надстройка (например, в форме создания, для определенных типов сообщений, при наличии вложения, адреса, предложения задачи, приглашения на собрание или определенных строковых шаблонов в тексте сообщения или сведениях о встрече).</span><span class="sxs-lookup"><span data-stu-id="26bb3-p104">If you answer yes to any of these questions, consider creating an **Outlook add-in**. Identify the context that will trigger your add-in (for example, the user being in a compose form, specific message types, the presence of an attachment, address, task suggestion, or meeting suggestion, or certain string patterns in the contents of an email or appointment).</span></span> 

    <span data-ttu-id="26bb3-125">Сведения о том, как активировать надстройку Outlook в соответствии с контекстом, см. в статье [Правила активации контекстных надстроек Outlook](/outlook/add-ins/activation-rules).</span><span class="sxs-lookup"><span data-stu-id="26bb3-125">To find out how you can contextually activate the Outlook add-in, see [Activation rules for Outlook add-ins](/outlook/add-ins/activation-rules).</span></span>

- <span data-ttu-id="26bb3-p105">Будут ли клиенты использовать надстройку для расширения возможностей при просмотре или создании документа? Если это так, рекомендуем создать **надстройку области задач**.</span><span class="sxs-lookup"><span data-stu-id="26bb3-p105">Will customers use the add-in to enhance the viewing or authoring experience of a document? If so, you may want to consider creating a **task pane add-in**.</span></span>

<span data-ttu-id="26bb3-128">Поддержка некоторых API для надстроек может отличаться в зависимости от того, в каком приложении Office и на какой платформе они работают (в Windows, Mac, веб-приложениях и на мобильных устройствах).</span><span class="sxs-lookup"><span data-stu-id="26bb3-128">Support for certain add-in APIs may differ between Office applications and the platform they are running on (Windows, Mac, Web, Mobile).</span></span> <span data-ttu-id="26bb3-129">Список поддерживаемых API по клиентам и платформам представлен на странице [Доступность ведущих приложений и платформ для надстроек Office](../overview/office-add-in-availability.md).</span><span class="sxs-lookup"><span data-stu-id="26bb3-129">To see the current API coverage by client and platform, see our [Office Add-in host and platform availability](../overview/office-add-in-availability.md) page.</span></span>  


## <a name="4-design-and-implement-the-user-experience-and-user-interface-for-the-add-in"></a><span data-ttu-id="26bb3-130">4. Разработка и реализация пользовательского интерфейса надстройки и взаимодействия с пользователем</span><span class="sxs-lookup"><span data-stu-id="26bb3-130">4. Design and implement the user experience and user interface for the add-in</span></span>

<span data-ttu-id="26bb3-p107">Разработайте быстрый и удобный пользовательский интерфейс, который будет согласован, прост в освоении и позволит выполнять основные действия всего за всего несколько этапов. Используйте сторонние интерфейсы API и веб-службы, соответствующие назначению надстройки.</span><span class="sxs-lookup"><span data-stu-id="26bb3-p107">Design a fast and fluid user experience that is consistent, easy to learn, with primary scenarios that require only a few steps to complete. Depending on the purpose of the add-in, make use of third-party APIs or web services.</span></span>

<span data-ttu-id="26bb3-133">Для реализации пользовательского интерфейса можно пользоваться любым из множества доступных средств веб-разработки и применять языки HTML и JavaScript.</span><span class="sxs-lookup"><span data-stu-id="26bb3-133">You can choose from a variety of web development tools and use HTML and JavaScript to implement the user interface.</span></span>


## <a name="5-create-an-xml-manifest-file-based-on-the-office-add-ins-manifest-schema"></a><span data-ttu-id="26bb3-134">5. Создание XML-файла манифеста на основе схемы манифеста надстроек Office</span><span class="sxs-lookup"><span data-stu-id="26bb3-134">5. Create an XML manifest file based on the Office Add-ins manifest schema</span></span>

<span data-ttu-id="26bb3-135">Создайте XML-манифест для идентификации надстройки и ее требований, укажите местоположение файлов HTML, JavaScript и CSS, которые использует надстройка. Кроме того, укажите размер и разрешения по умолчанию в соответствии с типом надстройки.</span><span class="sxs-lookup"><span data-stu-id="26bb3-135">Create an XML manifest to identify the add-in and its requirements, specify the locations of the HTML and any JavaScript and CSS files that the add-in uses, and depending on the type of the add-in, the default size and permissions.</span></span>

<span data-ttu-id="26bb3-p108">Для надстроек Outlook можно указать основанный на текущем сообщении или встрече контекст, в котором надстройка станет актуальной и будет отображаться в пользовательском интерфейсе Outlook. Кроме того, вы можете выбрать, на каких устройствах будет работать надстройка. Укажите в манифесте контекст в виде правил активации и поддерживаемых устройств.</span><span class="sxs-lookup"><span data-stu-id="26bb3-p108">For Outlook add-ins, you can specify the context, based on the current message or appointment, under which your add-in is relevant and you would like Outlook to make available in the UI. You can also decide which devices you want the add-in to support. In the manifest, specify the context as activation rules and the supported devices.</span></span>


## <a name="6-install-and-test-the-add-in"></a><span data-ttu-id="26bb3-139">6. Установка и тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="26bb3-139">6. Install and test the add-in</span></span>

<span data-ttu-id="26bb3-p109">Поместите HTML-файлы и файлы JavaScript и CSS (если они есть) на веб-серверы, указанные в файле манифеста надстройки. Процесс установки надстройки зависит от ее типа. Дополнительные сведения см. в статье [Загрузка неопубликованных надстроек Office для тестирования](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="26bb3-p109">Place the HTML files and any JavaScript and CSS files on the web servers that are specified in the add-in manifest file. The process to install an add-in depends on the type of the add-in. For details, see [Sideload Office Add-ins for testing](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

<span data-ttu-id="26bb3-p110">Если это надстройка Outlook, установите ее в почтовый ящик Exchange и укажите расположение манифеста надстройки в Центре администрирования Exchange (EAC). Дополнительные сведения см. в статье [Развертывание и установка надстроек Outlook для тестирования](/outlook/add-ins/testing-and-tips).</span><span class="sxs-lookup"><span data-stu-id="26bb3-p110">For Outlook add-ins, install it in an Exchange mailbox, and specify the location of the add-in manifest file in the Exchange Admin Center (EAC). For more information, see [Deploy and install Outlook add-ins for testing](/outlook/add-ins/testing-and-tips).</span></span>


## <a name="7-publish-the-add-in"></a><span data-ttu-id="26bb3-145">7. Публикация надстройки</span><span class="sxs-lookup"><span data-stu-id="26bb3-145">7. Publish the add-in</span></span>

<span data-ttu-id="26bb3-p111">Вы можете отправить надстройку в AppSource, где пользователи смогут ее скачать и установить. Кроме того, надстройки области задач и контентные надстройки можно публиковать в каталоге надстроек личной папки SharePoint или в общей сетевой папке, а надстройку Outlook для вашей организации можно развернуть непосредственно на сервере Exchange Server. Дополнительные сведения см. в статье [Публикация надстройки Office](../publish/publish.md).</span><span class="sxs-lookup"><span data-stu-id="26bb3-p111">You can submit the add-in to AppSource, from which customers can install the add-in. In addition, you can publish task pane and content add-ins to a private folder add-in catalog on SharePoint or to a shared network folder, and you can deploy an Outlook add-in directly on an Exchange server for your organization. For details, see [Publish your Office Add-in](../publish/publish.md).</span></span>


## <a name="8-maintain-the-add-in"></a><span data-ttu-id="26bb3-149">8. Обслуживание надстройки</span><span class="sxs-lookup"><span data-stu-id="26bb3-149">8. Maintain the add-in</span></span>

<span data-ttu-id="26bb3-p112">Если надстройка вызывает веб-службу, а вы вносите изменения в веб-службу уже после публикации надстройки, повторно публиковать ее не нужно. Тем не менее, если изменить какие-либо элементы или данные, которые вы уже отправили для надстройки (например, манифест надстройки, снимки экрана, значки, файлы HTML или JavaScript), надстройку необходимо будет повторно опубликовать.</span><span class="sxs-lookup"><span data-stu-id="26bb3-p112">If your add-in calls a web service, and if you make updates to the web service after publishing the add-in, you do not have to republish the add-in. However, if you change any items or data you submitted for your add-in, such as the add-in manifest, screenshots, icons, HTML or JavaScript files, you will need to republish the add-in.</span></span> 

<span data-ttu-id="26bb3-p113">В частности, если надстройка опубликована в AppSource, потребуется отправить ее заново, чтобы в AppSource были реализованы эти изменения. Повторно отправлять надстройку следует вместе с обновленным манифестом, включающим новый номер версии. Кроме того, необходимо обновить номер версии надстройки в форме отправки, чтобы он совпадал с версией нового манифеста. В случае надстроек Outlook необходимо убедиться, что элемент [Id](/office/dev/add-ins/reference/manifest/id) в манифесте надстройки содержит другой UUID.</span><span class="sxs-lookup"><span data-stu-id="26bb3-p113">In particular, if you have published the add-in to AppSource, you'll need to resubmit your add-in so that AppSource can implement those changes. You must resubmit your add-in with an updated add-in manifest that includes a new version number. You must also make sure to update the add-in version number in the submission form to match the new manifest's version number. For Outlook add-ins, you should make sure the [Id](/office/dev/add-ins/reference/manifest/id) element contains a different UUID in the add-in manifest.</span></span>
