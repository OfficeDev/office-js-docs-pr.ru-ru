---
title: Обзор платформы надстроек Office | Microsoft Docs
description: Используйте привычные веб-технологии, например HTML, CSS и JavaScript, для взаимодействия с Word, Excel, PowerPoint, OneNote, Project и Outlook, а также для расширения возможностей этих приложений.
ms.date: 07/05/2019
ms.topic: overview
scenarios: getting-started
localization_priority: Priority
ms.openlocfilehash: 9ed898db70eb86f32efd2acc9bcf98b6bdbc5b4b
ms.sourcegitcommit: 49af31060aa56c1e1ec1e08682914d3cbefc3f1c
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/29/2019
ms.locfileid: "36672819"
---
# <a name="office-add-ins-platform-overview"></a><span data-ttu-id="8005a-103">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="8005a-103">Office Add-ins platform overview</span></span>

<span data-ttu-id="8005a-p101">Платформу надстроек Office можно использовать для создания решений, которые расширяют возможности приложений Office и взаимодействуют с содержимым документов Office. В случае надстроек Office можно использовать привычные веб-технологии, например HTML, CSS и JavaScript, для взаимодействия с Word, Excel, PowerPoint, OneNote, Project и Outlook, а также для расширения возможностей этих приложений. Ваше решение может работать в Office на нескольких платформах, включая Windows, Mac, iPad и браузеры.</span><span class="sxs-lookup"><span data-stu-id="8005a-p101">You can use the Office Add-ins platform to build solutions that extend Office applications and interact with content in Office documents. With Office Add-ins, you can use familiar web technologies such as HTML, CSS, and JavaScript to extend and interact with Word, Excel, PowerPoint, OneNote, Project, and Outlook. Your solution can run in Office across multiple platforms, including Office on Windows, Office Online, Office for Mac, and Office for iPad.</span></span>

<span data-ttu-id="8005a-p102">Надстройки Office могут делать почти все, на что способна веб-страница в браузере. Платформу надстроек Office можно использовать для следующих целей:</span><span class="sxs-lookup"><span data-stu-id="8005a-p102">Office Add-ins can do almost anything a webpage can do inside a browser. Use the Office Add-ins platform to:</span></span>

-  <span data-ttu-id="8005a-p103">**Добавление новых возможностей к клиентам Office.** Подключайте внешние данные к Office, автоматизируйте обработку документов Office, добавляйте в клиенты Office функции сторонних решений и многое другое. Например, с помощью API Microsoft Graph можно подключаться к данным, повышая производительность.</span><span class="sxs-lookup"><span data-stu-id="8005a-p103">**Add new functionality to Office clients** - Bring external data into Office, automate Office documents, expose third-party functionality in Office clients, and more. For example, use Microsoft Graph API to connect to data that drives productivity.</span></span>

-  <span data-ttu-id="8005a-111">**Создание оснащенных различными функциями интерактивных объектов, которые можно внедрить в документы Office.** Внедряйте карты, диаграммы и интерактивные визуализации, которые пользователи могут добавлять в свои электронные таблицы Excel и презентации PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="8005a-111">**Create new rich, interactive objects that can be embedded in Office documents** - Embed maps, charts, and interactive visualizations that users can add to their own Excel spreadsheets and PowerPoint presentations.</span></span>

## <a name="how-are-office-add-ins-different-from-com-and-vsto-add-ins"></a><span data-ttu-id="8005a-112">Чем надстройки Office отличаются от надстроек COM и VSTO?</span><span class="sxs-lookup"><span data-stu-id="8005a-112">How are Office Add-ins different from COM and VSTO add-ins?</span></span>

<span data-ttu-id="8005a-p104">Надстройки COM и VSTO представляют собой более ранние решения для интеграции Office, которые работают только в Office для Windows. В отличие от надстроек COM, надстройкам Office не требуется код, который выполняется на устройстве пользователя или в клиенте Office. В надстройках Office ведущее приложение, например Excel, считывает манифест надстройки и подключает настраиваемые кнопки ленты и команды меню надстройки в пользовательском интерфейсе. При необходимости оно загружает JavaScript и HTML-код надстройки, который выполняется в "песочнице" в контексте браузера.</span><span class="sxs-lookup"><span data-stu-id="8005a-p104">COM or VSTO add-ins are earlier Office integration solutions that run only on Office on Windows. Unlike COM add-ins, Office Add-ins don't involve code that runs on the user's device or in the Office client. For an Office Add-in, the host application, for example Excel, reads the add-in manifest and hooks up the add-in’s custom ribbon buttons and menu commands in the UI. When needed, it loads the add-in's JavaScript and HTML code, which executes in the context of a browser in a sandbox.</span></span>

<span data-ttu-id="8005a-117">Преимущества надстроек Office над надстройками, созданными с помощью VBA, модели COM или VSTO:</span><span class="sxs-lookup"><span data-stu-id="8005a-117">Office Add-ins provide the following advantages over add-ins built using VBA, COM, or VSTO:</span></span>

- <span data-ttu-id="8005a-p105">Кроссплатформенная поддержка. Надстройки Office работают в Office в Интернете, Office для Windows, Office для Mac и Office для iPad.</span><span class="sxs-lookup"><span data-stu-id="8005a-p105">Cross-platform support. Office Add-ins run in Office on Windows, Mac, iOS, and Office Online.</span></span>

- <span data-ttu-id="8005a-p106">Централизованное развертывание и распространение. Администраторы могут централизованно развертывать надстройки Office в организации.</span><span class="sxs-lookup"><span data-stu-id="8005a-p106">Centralized deployment and distribution. Admins can deploy Office Add-ins centrally across an organization.</span></span>

- <span data-ttu-id="8005a-p107">Легкий доступ через AppSource. Вы можете сделать свое решение доступным широкой аудитории, отправив его в AppSource.</span><span class="sxs-lookup"><span data-stu-id="8005a-p107">Easy access via AppSource. You can make your solution available to a broad audience by submitting it to AppSource.</span></span>

- <span data-ttu-id="8005a-p108">Стандартная веб-технология. Вы можете использовать любую библиотеку для создания надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="8005a-p108">Based on standard web technology. You can use any library you like to build Office Add-ins.</span></span>

## <a name="components-of-an-office-add-in"></a><span data-ttu-id="8005a-126">Компоненты надстройки Office</span><span class="sxs-lookup"><span data-stu-id="8005a-126">Components of an Office Add-in</span></span>

<span data-ttu-id="8005a-p109">Надстройка Office включает в себя два основных компонента — XML-файл манифеста и веб-приложение. Манифест определяет различные параметры, включая способ интеграции надстройки с клиентами Office. Веб-приложение должно быть размещено на веб-сервере или в службе веб-хостинга, например в Microsoft Azure.</span><span class="sxs-lookup"><span data-stu-id="8005a-p109">An Office Add-in includes two basic components: an XML manifest file, and your own web application. The manifest defines various settings, including how your add-in integrates with Office clients. Your web application needs to be hosted on a web server, or web hosting service, such as Microsoft Azure.</span></span>

<span data-ttu-id="8005a-130">*Рис. 1. Манифест надстройки (XML) + веб-страница (HTML, JS) = надстройка Office*</span><span class="sxs-lookup"><span data-stu-id="8005a-130">*Figure 1. Add-in manifest (XML) + webpage (HTML, JS) = an Office Add-in*</span></span>

![Манифест + веб-страница = надстройка Office](../images/about-addins-manifestwebpage.png)

### <a name="manifest"></a><span data-ttu-id="8005a-132">Манифест</span><span class="sxs-lookup"><span data-stu-id="8005a-132">Manifest</span></span>

<span data-ttu-id="8005a-133">Этот манифест представляет собой XML-файл, который определяет следующие параметры и возможности надстройки:</span><span class="sxs-lookup"><span data-stu-id="8005a-133">The manifest is an XML file that specifies settings and capabilities of the add-in, such as:</span></span>

- <span data-ttu-id="8005a-134">Отображаемое имя, описание, идентификатор, версию и языковой стандарт по умолчанию надстройки.</span><span class="sxs-lookup"><span data-stu-id="8005a-134">The add-in's display name, description, ID, version, and default locale.</span></span>

- <span data-ttu-id="8005a-135">Способ интеграции надстройки с Office.</span><span class="sxs-lookup"><span data-stu-id="8005a-135">How the add-in integrates with Office.</span></span>  

- <span data-ttu-id="8005a-136">Уровень разрешений и требования для доступа к данным для надстройки.</span><span class="sxs-lookup"><span data-stu-id="8005a-136">The permission level and data access requirements for the add-in.</span></span>

### <a name="web-app"></a><span data-ttu-id="8005a-137">Веб-приложение</span><span class="sxs-lookup"><span data-stu-id="8005a-137">Web app</span></span>

<span data-ttu-id="8005a-p110">Самая простая надстройка Office представляет собой статическую HTML-страницу, которая отображается в приложении Office, но не взаимодействует ни с документом Office, ни с каким-либо другим ресурсом в Интернете. Для создания кода, который взаимодействует с документами Office или позволяет пользователю взаимодействовать с веб-ресурсами из ведущего приложения Office, можно применять любые технологии, как клиентские, так и серверные, которые поддерживает ваш поставщик услуг размещения (например, ASP.NET, PHP или Node.js). Для взаимодействия с клиентами и документами Office можно использовать интерфейсы API JavaScript Office.js.</span><span class="sxs-lookup"><span data-stu-id="8005a-p110">The most basic Office Add-in consists of a static HTML page that is displayed inside an Office application, but that doesn't interact with either the Office document or any other Internet resource. However, to create an experience that interacts with Office documents or allows the user to interact with online resources from an Office host application, you can use any technologies, both client and server side, that your hosting provider supports (such as ASP.NET, PHP, or Node.js). To interact with Office clients and documents, you use the Office.js JavaScript APIs.</span></span>

<span data-ttu-id="8005a-141">*Рисунок 2. Компоненты надстройки Hello World для Office*</span><span class="sxs-lookup"><span data-stu-id="8005a-141">*Figure 2. Components of a Hello World Office Add-in*</span></span>

![Компоненты надстройки Hello World](../images/about-addins-componentshelloworldoffice.png)

## <a name="extending-and-interacting-with-office-clients"></a><span data-ttu-id="8005a-143">Расширение возможностей и взаимодействие с клиентами Office</span><span class="sxs-lookup"><span data-stu-id="8005a-143">Extending and interacting with Office clients</span></span>

<span data-ttu-id="8005a-144">Что позволяют надстройки Office в ведущем приложении Office:</span><span class="sxs-lookup"><span data-stu-id="8005a-144">Office Add-ins can do the following within an Office host application:</span></span>

-  <span data-ttu-id="8005a-145">расширение функциональных возможностей (в любом приложении Office);</span><span class="sxs-lookup"><span data-stu-id="8005a-145">Extend functionality (any Office application)</span></span>

-  <span data-ttu-id="8005a-146">создание новых объектов (Excel или PowerPoint).</span><span class="sxs-lookup"><span data-stu-id="8005a-146">Create new objects (Excel or PowerPoint)</span></span>
 
### <a name="extend-office-functionality"></a><span data-ttu-id="8005a-147">Расширение функциональных возможностей Office</span><span class="sxs-lookup"><span data-stu-id="8005a-147">Extend Office functionality</span></span>

<span data-ttu-id="8005a-148">Добавить новые возможности в приложения Office можно посредством следующего:</span><span class="sxs-lookup"><span data-stu-id="8005a-148">You can add new functionality to Office applications via the following:</span></span>  

-  <span data-ttu-id="8005a-149">настраиваемые кнопки ленты и команды меню (в совокупности зовутся "командами надстройки");</span><span class="sxs-lookup"><span data-stu-id="8005a-149">Custom ribbon buttons and menu commands (collectively called “add-in commands”)</span></span>

-  <span data-ttu-id="8005a-150">вставляемые области задач.</span><span class="sxs-lookup"><span data-stu-id="8005a-150">Insertable task panes</span></span>

<span data-ttu-id="8005a-151">Пользовательский интерфейс и области задач указаны в манифесте надстройки.</span><span class="sxs-lookup"><span data-stu-id="8005a-151">Custom UI and task panes are specified in the add-in manifest.</span></span>  

#### <a name="custom-buttons-and-menu-commands"></a><span data-ttu-id="8005a-152">Настраиваемые кнопки и команды меню</span><span class="sxs-lookup"><span data-stu-id="8005a-152">Custom buttons and menu commands</span></span>  

<span data-ttu-id="8005a-p111">Вы можете добавлять настраиваемые кнопки ленты и элементы меню на ленту в Office в Интернете и Office для Windows. Благодаря этому пользователям будет проще получать доступ к надстройке непосредственно из приложения Office. С помощью кнопок можно выполнять различные действия, например отображение области задач с пользовательским HTML или выполнение функции JavaScript.</span><span class="sxs-lookup"><span data-stu-id="8005a-p111">You can add custom ribbon buttons and menu items to the ribbon in Office on Windows Desktop and Office Online. This makes it easy for users to access your add-in directly from their Office application. Command buttons can launch different actions such as showing a task pane with custom HTML or executing a JavaScript function.</span></span>  

<span data-ttu-id="8005a-156">*Рис. 3. Команды надстройки на ленте*</span><span class="sxs-lookup"><span data-stu-id="8005a-156">*Figure 3. Add-in commands in the ribbon*</span></span>

![Настраиваемые кнопки и команды меню](../images/about-addins-addincommands.png)

#### <a name="task-panes"></a><span data-ttu-id="8005a-158">Области задач</span><span class="sxs-lookup"><span data-stu-id="8005a-158">Task panes</span></span>  

<span data-ttu-id="8005a-p112">Для работы с решением пользователи могут использовать не только команды надстройки, но и области задач. В клиентах, не поддерживающих команды надстроек (Office 2013 и Office для iPad), надстройка запускается в виде области задач. Пользователи запускают надстройки области задач с помощью кнопки **Мои надстройки** на вкладке **Вставка**.</span><span class="sxs-lookup"><span data-stu-id="8005a-p112">You can use task panes in addition to add-in commands to enable users to interact with your solution. Clients that do not support add-in commands (Office 2013 and Office for iPad) run your add-in as a task pane. Users launch task pane add-ins via the **My Add-ins** button on the **Insert** tab.</span></span>

<span data-ttu-id="8005a-162">*Рис. 4. Область задач*</span><span class="sxs-lookup"><span data-stu-id="8005a-162">*Figure 4. Task pane*</span></span>

![Использование областей задач в дополнение к командам надстроек](../images/about-addins-taskpane.png)

### <a name="extend-outlook-functionality"></a><span data-ttu-id="8005a-164">Расширение возможностей Outlook</span><span class="sxs-lookup"><span data-stu-id="8005a-164">Extend Outlook functionality</span></span>

<span data-ttu-id="8005a-p113">Надстройки Outlook могут расширять функциональные возможности ленты Office и в зависимости от контекста отображаться рядом с просматриваемым или создаваемым элементом Outlook. Они могут взаимодействовать с письмами, приглашениями на собрания, ответами на приглашения на собрания, сообщениями об отмене собраний или данными о встречах, когда пользователь просматривает полученный элемент, отвечает на него или создает новый.</span><span class="sxs-lookup"><span data-stu-id="8005a-p113">Outlook add-ins can extend the Office ribbon and also display contextually next to an Outlook item when you're viewing or composing it. They can work with an email message, meeting request, meeting response, meeting cancellation, or appointment when a user is viewing a received item or replying or creating a new item.</span></span> 

<span data-ttu-id="8005a-p114">Надстройки Outlook могут получать доступ к зависящей от контекста информации, предоставляемой элементом, например к адресу или идентификатору для отслеживания, а затем использовать эти данные для доступа к дополнительным сведениям на сервере и в веб-службах с целью повышения удобства работы пользователей. В большинстве случаев надстройка Outlook может без изменений работать в ведущем приложении Outlook, обеспечивая согласованность интерфейса на компьютерах, в Интернете, на планшетах и на мобильных устройствах.</span><span class="sxs-lookup"><span data-stu-id="8005a-p114">Outlook add-ins can access contextual information from the item, such as an address or tracking ID, and then use that data to access additional information on the server and from web services to create compelling user experiences. In most cases, an Outlook add-in runs without modification on the various supporting host applications, including Outlook, Outlook for Mac, Outlook Web App, and Outlook Web App for devices, to provide a seamless experience on the desktop, web, and tablet and mobile devices.</span></span>

<span data-ttu-id="8005a-169">Обзор надстроек Outlook см. в статье [Общие сведения о надстройках Outlook](/outlook/add-ins/).</span><span class="sxs-lookup"><span data-stu-id="8005a-169">For an overview of Outlook add-ins, see [Outlook add-ins overview](/outlook/add-ins/).</span></span>

### <a name="create-new-objects-in-office-documents"></a><span data-ttu-id="8005a-170">Создание новых объектов в документах Office</span><span class="sxs-lookup"><span data-stu-id="8005a-170">Create new objects in Office documents</span></span>

<span data-ttu-id="8005a-p115">Вы можете внедрить веб-объекты, или контентные надстройки, в документы Excel и PowerPoint. Благодаря контентным надстройкам можно интегрировать мультимедиа (например, видеопроигрыватель YouTube или галерею рисунков), полнофункциональные веб-визуализации данных и другое внешнее содержимое.</span><span class="sxs-lookup"><span data-stu-id="8005a-p115">You can embed web-based objects called content add-ins within Excel and PowerPoint documents. With content add-ins, you can integrate rich, web-based data visualizations, media (such as a YouTube video player or a picture gallery), and other external content.</span></span>

<span data-ttu-id="8005a-173">*Рис. 5. Контентная надстройка*</span><span class="sxs-lookup"><span data-stu-id="8005a-173">*Figure 5. Content add-in*</span></span>

![Внедрение веб-объектов, которые зовутся контентными надстройками](../images/about-addins-contentaddin.png)

## <a name="office-javascript-apis"></a><span data-ttu-id="8005a-175">API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="8005a-175">Office JavaScript APIs</span></span>

<span data-ttu-id="8005a-p116">API JavaScript для Office содержат объекты и элементы для создания надстроек и взаимодействия с содержимым Office и веб-службами. В Excel, Outlook, Word, PowerPoint, OneNote и Project используется общая объектная модель. Кроме того, существуют расширенные объектные модели для Excel и Word. Эти API предоставляют доступ к известным объектам, таким как абзацы и книги, что упрощает создание надстройки для определенного ведущего приложения.</span><span class="sxs-lookup"><span data-stu-id="8005a-p116">The Office JavaScript APIs contain objects and members for building add-ins and interacting with Office content and web services. There is a common object model that is shared by Excel, Outlook, Word, PowerPoint, OneNote and Project. There are also more extensive host-specific object models for Excel and Word. These APIs provide access to well-known objects such as paragraphs and workbooks, which makes it easier to create an add-in for a specific host.</span></span>  

## <a name="next-steps"></a><span data-ttu-id="8005a-180">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="8005a-180">Next steps</span></span>

<span data-ttu-id="8005a-181">Чтобы создать свою первую надстройку Office менее чем за 5 минут, попробуйте воспользоваться кратким руководством для [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md) или [Project](../quickstarts/project-quickstart.md).</span><span class="sxs-lookup"><span data-stu-id="8005a-181">To build your first Office Add-in in less than 5 minutes, try out the quick start for [Excel](../quickstarts/excel-quickstart-jquery.md), [Outlook](/outlook/add-ins/quick-start?context=office/dev/add-ins/context), [Word](../quickstarts/word-quickstart.md), [OneNote](../quickstarts/onenote-quickstart.md), [PowerPoint](../quickstarts/powerpoint-quickstart.md), or [Project](../quickstarts/project-quickstart.md).</span></span> <span data-ttu-id="8005a-182">Сразу же приступить к созданию надстроек можно с помощью Visual Studio или любого другого редактора.</span><span class="sxs-lookup"><span data-stu-id="8005a-182">You can start building add-ins right away using Visual Studio or any other editor.</span></span> 

<span data-ttu-id="8005a-183">Чтобы начать планировать создание решений с удобным и привлекательным интерфейсом, ознакомьтесь с [рекомендациями по проектированию](../design/add-in-design.md) и другими [рекомендациями](../concepts/add-in-development-best-practices.md), касающимися надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="8005a-183">To start planning solutions that create effective and compelling user experiences, get familiar with the [design guidelines](../design/add-in-design.md) and [best practices](../concepts/add-in-development-best-practices.md) for Office Add-ins.</span></span>

## <a name="see-also"></a><span data-ttu-id="8005a-184">См. также</span><span class="sxs-lookup"><span data-stu-id="8005a-184">See also</span></span>

- [<span data-ttu-id="8005a-185">Примеры надстроек Office</span><span class="sxs-lookup"><span data-stu-id="8005a-185">Office Add-in samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel,Outlook,PowerPoint,Word)
- [<span data-ttu-id="8005a-186">Общие сведения об API JavaScript для Office</span><span class="sxs-lookup"><span data-stu-id="8005a-186">Understanding the JavaScript API for Office</span></span>](../develop/understanding-the-javascript-api-for-office.md)
- [<span data-ttu-id="8005a-187">Доступность ведущих приложений и платформ для надстроек Office</span><span class="sxs-lookup"><span data-stu-id="8005a-187">Office Add-in host and platform availability</span></span>](../overview/office-add-in-availability.md)
