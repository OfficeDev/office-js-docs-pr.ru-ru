---
title: Создание и отладка надстроек Office в Visual Studio
description: ''
ms.date: 03/14/2018
ms.openlocfilehash: 3e4fbcd3919be0d5510b36ae77a6e3706eab9689
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19437607"
---
# <a name="create-and-debug-office-add-ins-in-visual-studio"></a><span data-ttu-id="7c693-102">Создание и отладка надстроек Office в Visual Studio</span><span class="sxs-lookup"><span data-stu-id="7c693-102">Create and debug Office Add-ins in Visual Studio</span></span>

<span data-ttu-id="7c693-p101">Из этой статьи можно узнать, как создать свою первую надстройку Office с помощью Visual Studio. Действия, описанные в этой статье, относятся к версии Visual Studio 2015. Если вы используете другую версию Visual Studio, действия могут немного отличаться.</span><span class="sxs-lookup"><span data-stu-id="7c693-p101">This article describes how to use Visual Studio to create your first Office Add-in. The steps in this article based on Visual Studio 2015. If you're using another version of Visual Studio, the procedures might vary slightly.</span></span>

> [!NOTE]
> <span data-ttu-id="7c693-106">Чтобы начать работу над надстройкой OneNote, см. статью [Создание первой надстройки OneNote](../onenote/onenote-add-ins-getting-started.md).</span><span class="sxs-lookup"><span data-stu-id="7c693-106">To get started with an add-in for OneNote, see [Build your first OneNote add-in](../onenote/onenote-add-ins-getting-started.md).</span></span>

## <a name="create-an-office-add-in-project-in-visual-studio"></a><span data-ttu-id="7c693-107">Создание проекта надстройки Office в Visual Studio</span><span class="sxs-lookup"><span data-stu-id="7c693-107">Create an Office Add-in project in Visual Studio</span></span>


<span data-ttu-id="7c693-p102">Для начала убедитесь, что у вас установлены инструменты [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) и версия Microsoft Office. Вы можете присоединиться к [программе для разработчиков Office 365](https://developer.microsoft.com/en-us/office/dev-program) или получить [последнюю версию](../develop/install-latest-office-version.md), следуя приведенным ниже инструкциям.</span><span class="sxs-lookup"><span data-stu-id="7c693-p102">To get started, make sure you have the [Office Developer Tools](https://www.visualstudio.com/features/office-tools-vs.aspx) installed, and a version of Microsoft Office. You can join the [Office 365 Developer Program](https://developer.microsoft.com/en-us/office/dev-program), or follow these instructions to get the [latest version](../develop/install-latest-office-version.md).</span></span>


1. <span data-ttu-id="7c693-110">В строке меню Visual Studio выберите **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="7c693-110">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="7c693-111">В списке типов проектов **Visual C#** или **Visual Basic** разверните узел **Office/SharePoint**, выберите **Веб-надстройки**, а затем выберите один из проектов надстроек.</span><span class="sxs-lookup"><span data-stu-id="7c693-111">In the list of project types under  **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose  **Web Add-ins**, and then select one of the Add-in projects.</span></span>  
    
3. <span data-ttu-id="7c693-112">Введите имя проекта и нажмите кнопку **ОК**, чтобы создать его.</span><span class="sxs-lookup"><span data-stu-id="7c693-112">Name the project, and then choose  **OK** to create the project.</span></span>
    
4. <span data-ttu-id="7c693-p103">В Visual Studio создается решение, и соответствующие два проекта отображаются в **обозревателе решений**. В Visual Studio откроется страница Home.html по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="7c693-p103">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The default Home.html page opens in Visual Studio.</span></span>
    
<span data-ttu-id="7c693-115">В Visual Studio 2015 некоторые шаблоны проектов надстроек обновлены с добавлением новых функций.</span><span class="sxs-lookup"><span data-stu-id="7c693-115">In Visual Studio 2015, some of the add-in project templates have been updated to reflect additional functionality:</span></span>


- <span data-ttu-id="7c693-p104">Контентные надстройки могут отображаться в основном тексте документов Access и PowerPoint, а не только в электронных таблицах Excel. Вы также можете выбрать параметр "Базовый проект", чтобы создать проект простой контентной надстройки с минимумом начального кода, или параметр "Проект визуализации документов" (только для Access и Excel), чтобы создать более функциональную контентную надстройку, которая включает начальный код для визуализации данных и привязки к ним.</span><span class="sxs-lookup"><span data-stu-id="7c693-p104">Content add-ins can appear in the body of Access and PowerPoint documents, in addition to Excel spreadsheets. You can also choose the Basic Project option to create a basic content add-in project with minimal starter code, or the Document Visualization Project option (for Access and Excel only) to create a more full-featured content add-in that includes starter code to visualize and bind to data.</span></span>
    
- <span data-ttu-id="7c693-118">Надстройки Outlook позволяют не только встраивать надстройку в сообщения электронной почты или встречи, но и указывать, доступна ли она при создании и просмотре сообщения или встречи.</span><span class="sxs-lookup"><span data-stu-id="7c693-118">Outlook add-ins include options not just for including your add-in in email messages or appointments, but also for specifying whether the add-in is available when an email message or appointment is being composed as well as read.</span></span>
    

> [!NOTE]
> <span data-ttu-id="7c693-p105">Назначение большинства параметров в Visual Studio очевидно из названия, кроме флажка **Сообщение электронной почты**. Установите этот флажок, если вам нужно создать надстройку Outlook, которая отображается не только с почтовыми элементами, но и с приглашениями на собрания, текстами ответов и отмен.</span><span class="sxs-lookup"><span data-stu-id="7c693-p105">In Visual Studio most options are understandable from their descriptions except for the  **Email Message** checkbox. Use that checkbox if you want to create an Outlook add-in that appears not just with mail items, but also with meeting requests, responses, and cancellations.</span></span>

<span data-ttu-id="7c693-121">После завершения работы мастера Visual Studio создает решение, которое содержит два проекта.</span><span class="sxs-lookup"><span data-stu-id="7c693-121">When you've completed the wizard, Visual Studio creates a solution for you that contains two projects.</span></span>



|<span data-ttu-id="7c693-122">**Проект**</span><span class="sxs-lookup"><span data-stu-id="7c693-122">**Project**</span></span>|<span data-ttu-id="7c693-123">**Описание**</span><span class="sxs-lookup"><span data-stu-id="7c693-123">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="7c693-124">Проект надстройки</span><span class="sxs-lookup"><span data-stu-id="7c693-124">Add-in project</span></span>|<span data-ttu-id="7c693-p106">Содержит только XML-файл манифеста, который содержит все параметры, описывающие надстройку. Эти параметры помогают ведущему приложению Office определять, когда и где должна активироваться надстройка. Visual Studio создает содержимое этого файла за вас, чтобы вы могли сразу запустить проект и начать использовать надстройку. Вы можете менять эти параметры в любой момент с помощью редактора манифеста.</span><span class="sxs-lookup"><span data-stu-id="7c693-p106">Contains only an XML manifest file, which contains all the settings that describe your add-in. These settings help the Office host determine when your add-in should be activated and where the add-in should appear. Visual Studio generates the contents of this file for you so that you can run the project and use your add-in immediately. You change these settings any time by using the Manifest editor.</span></span>|
|<span data-ttu-id="7c693-129">Проект веб-приложения</span><span class="sxs-lookup"><span data-stu-id="7c693-129">Web application project</span></span>|<span data-ttu-id="7c693-p107">Содержит страницы контента надстройки, включающие все файлы и ссылки на файлы, необходимые для разработки страниц HTML и JavaScript с поддержкой Office. При разработке надстройки Visual Studio размещает веб-приложение на локальном сервере IIS. Когда вы будете готовы опубликовать надстройку, потребуется найти сервер для размещения проекта. Дополнительные сведения о проектах веб-приложений ASP.NET см. в статье [Веб-проекты ASP.NET](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).</span><span class="sxs-lookup"><span data-stu-id="7c693-p107">Contains the content pages of your add-in, including all the files and file references that you need to develop Office-aware HTML and JavaScript pages. While you develop your add-in, Visual Studio hosts the web application on your local IIS server. When you're ready to publish, you'll have to find a server to host this project.To learn more about ASP.NET web application projects, see [ASP.NET Web Projects](http://msdn.microsoft.com/en-us/library/cdcd712f-96b0-4165-8b5d-9d0566650a28%28Office.15%29.aspx).</span></span>|

## <a name="modify-your-add-in-settings"></a><span data-ttu-id="7c693-133">Изменение параметров надстроек</span><span class="sxs-lookup"><span data-stu-id="7c693-133">Modify your add-in settings</span></span>


<span data-ttu-id="7c693-p108">Чтобы изменить параметры надстройки, отредактируйте XML-файл манифеста проекта. В **обозревателе решений** разверните узел проекта надстройки, откройте папку, содержащую XML-манифест, и выберите его. Вы можете навести указатель мыши на любой элемент в файле, чтобы увидеть подсказку с описанием назначения этого элемента. Дополнительные сведения о файле манифеста см. в статье [XML-манифест надстроек Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="7c693-p108">To modify the settings of your add-in, edit the XML manifest file of the project. In  **Solution Explorer**, expand the add-in project node, expand the folder that contains the XML manifest, and choose the XML manifest. You can point to any element in the file to view a tooltip that describes the purpose of the element. For more information about the manfiest file, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


## <a name="develop-the-contents-of-your-add-in"></a><span data-ttu-id="7c693-138">Разработка содержимого надстройки</span><span class="sxs-lookup"><span data-stu-id="7c693-138">Develop the contents of your add-in</span></span>


<span data-ttu-id="7c693-139">Проект надстройки позволяет изменить ее параметры, а веб-приложение предоставляет содержимое, которое отображается в надстройке.</span><span class="sxs-lookup"><span data-stu-id="7c693-139">While the add-in project lets you modify the settings that describe your add-in, the web application provides the content that appears in the add-in.</span></span> 

<span data-ttu-id="7c693-p109">Проект веб-приложения содержит HTML-файл страницы по умолчанию и файл JavaScript, которые можно использовать для начала работы. Проект также содержит файл JavaScript, который используется для всех страниц, добавленных в проект. Эти файлы очень важны, так как в них содержатся ссылки на другие библиотеки JavaScript, в том числе API JavaScript для Office.</span><span class="sxs-lookup"><span data-stu-id="7c693-p109">The web application project contains a default HTML page and JavaScript file that you can use to get started. The project also contains a JavaScript file that is common to all pages that you add to your project. These files are convenient because they contain references to other JavaScript libraries including the JavaScript API for Office.</span></span> 

<span data-ttu-id="7c693-p110">По мере усложнения надстройки вы можете добавлять дополнительные HTML-файлы и файлы JavaScript. Содержимое файлов HTML и JavaScript по умолчанию можно использовать как примеры типов ссылок, которые можно добавить на другие страницы проекта, чтобы они работали с вашей надстройкой. В следующей таблице описываются файлы HTML и JavaScript по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="7c693-p110">As your add-in becomes more sophisticated, you can add more HTML and JavaScript files. You can use the contents of the default HTML and JavaScript files as examples of the types of references you might want to add to other pages in your project to make them work with your add-in. The following table describes default HTML and JavaScript files.</span></span>



|<span data-ttu-id="7c693-146">**Файл**</span><span class="sxs-lookup"><span data-stu-id="7c693-146">**File**</span></span>|<span data-ttu-id="7c693-147">**Описание**</span><span class="sxs-lookup"><span data-stu-id="7c693-147">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="7c693-148">**Home.html**</span><span class="sxs-lookup"><span data-stu-id="7c693-148">**Home.html**</span></span>|<span data-ttu-id="7c693-p111">Эта HTML-страница надстройки по умолчанию располагается в папке **Home** проекта. Она отображается первой в надстройки при ее активации в сообщении электронной почты или элементе встречи. Этот файл удобен тем, что содержит все ссылки на файлы, необходимые для начала работы. Когда вы будете готовы к созданию своей первой надстройки, просто добавьте HTML-код в этот файл.</span><span class="sxs-lookup"><span data-stu-id="7c693-p111">Located in the  **Home** folder of the project, this is default HTML page of the add-in. This page appears as the first page inside of the add-in when it is activated in a document, email message or appointment item. This file is convenient because it contains all of the file references that you need to get started. When you are ready to create your first add-in, just add your HTML code to this file.</span></span>|
|<span data-ttu-id="7c693-153">**Home.js**</span><span class="sxs-lookup"><span data-stu-id="7c693-153">**Home.js**</span></span>|<span data-ttu-id="7c693-p112">Файл JavaScript по умолчанию, расположенный в папке **Home** проекта и связанный со страницей Home.js. В файле Home.js можно разместить код, связанный с поведением страницы Home.html. Файл Home.js содержит пример кода, который поможет вам с созданием приложения.</span><span class="sxs-lookup"><span data-stu-id="7c693-p112">Located in the  **Home** folder of the project, this is the JavaScript file associated with the Home.js page. You can place any code that is specific to the behavior of the Home.html page in the Home.js file. The Home.js file contains some example code to get you started.</span></span>|
|<span data-ttu-id="7c693-157">**App.js**</span><span class="sxs-lookup"><span data-stu-id="7c693-157">**App.js**</span></span>|<span data-ttu-id="7c693-p113">Файл JavaScript по умолчанию, расположенный в папке **Add-in**. Вы можете поместить в файл App.js код, который относится к поведению нескольких страниц вашей надстройки. Файл App.js содержит некоторые примеры кода, которые помогут вам приступить к работе.</span><span class="sxs-lookup"><span data-stu-id="7c693-p113">Located in the  **Add-in** folder of the project, this is the default JavaScript file of the entire add-in. You can place code that is common to the behavior of multiple pages of your add-in in the App.js file. The App.js file contains some example code to get you started.</span></span>|

> [!NOTE]
> <span data-ttu-id="7c693-p114">Эти файлы необязательно использовать. Вы можете добавлять в проект другие файлы. Если в качестве начальной страницы надстройки должен отображаться другой HTML-файл, откройте редактор манифестов и задайте значение для свойства **SourceLocation**, указав имя нужного файла.</span><span class="sxs-lookup"><span data-stu-id="7c693-p114">You don't have to use these files. Feel free to add other files to the project and use those instead. If you want another HTML file to appear as the initial page of the add-in, open the manifest editor, and then point the  **SourceLocation** property to the name of the file.</span></span>


## <a name="debug-your-add-in"></a><span data-ttu-id="7c693-164">Отладка надстройки</span><span class="sxs-lookup"><span data-stu-id="7c693-164">Debug your add-in</span></span>


<span data-ttu-id="7c693-165">Если все готово к запуску надстройки, просмотрите свойства, связанные с построением и отладкой, затем запустите решение.</span><span class="sxs-lookup"><span data-stu-id="7c693-165">When you are ready to start your add-in, review build and debug related properties, and then start the solution.</span></span>


### <a name="review-the-build-and-debug-properties"></a><span data-ttu-id="7c693-166">Просмотр свойств, связанных с построением и отладкой</span><span class="sxs-lookup"><span data-stu-id="7c693-166">Review the build and debug properties</span></span>

<span data-ttu-id="7c693-p115">Перед запуском решения убедитесь, что Visual Studio откроет нужное ведущее приложение. Эти сведения отображаются на страницах свойств проекта вместе с несколькими другими свойствами, относящимися к построению и отладке надстройки.</span><span class="sxs-lookup"><span data-stu-id="7c693-p115">Before you start the solution, verify that Visual Studio will open the host application that you want. That information appears in the property pages of the project along with several other properties that relate to building and debugging the add-in.</span></span>


### <a name="to-open-the-property-pages-of-a-project"></a><span data-ttu-id="7c693-169">Открытие страниц свойств проекта</span><span class="sxs-lookup"><span data-stu-id="7c693-169">To open the property pages of a project</span></span>


1. <span data-ttu-id="7c693-170">В **обозревателе решений** выберите имя проекта.</span><span class="sxs-lookup"><span data-stu-id="7c693-170">In  **Solution Explorer**, choose the project name.</span></span>
    
2. <span data-ttu-id="7c693-171">В панели меню выберите пункты **Вид** и **Окно свойств**.</span><span class="sxs-lookup"><span data-stu-id="7c693-171">On the menu bar, choose  **View**,  **Properties Window**.</span></span>
    
<span data-ttu-id="7c693-172">В следующей таблице описываются свойства проекта.</span><span class="sxs-lookup"><span data-stu-id="7c693-172">The following table describes the properties of the project.</span></span>



|<span data-ttu-id="7c693-173">**Свойство**</span><span class="sxs-lookup"><span data-stu-id="7c693-173">**Property**</span></span>|<span data-ttu-id="7c693-174">**Описание**</span><span class="sxs-lookup"><span data-stu-id="7c693-174">**Description**</span></span>|
|:-----|:-----|
|<span data-ttu-id="7c693-175">**Действие при запуске**</span><span class="sxs-lookup"><span data-stu-id="7c693-175">**Start Action**</span></span>|<span data-ttu-id="7c693-176">Указывает, где производить отладку надстройки — в классическом приложении Office или в клиенте Office Online в заданном браузере.</span><span class="sxs-lookup"><span data-stu-id="7c693-176">Specifies whether to debug your add-in in an Office desktop client or in an Office Online client in the specified browser.</span></span>|
|<span data-ttu-id="7c693-177">**Документ при запуске** (только для контентных надстроек и надстроек области задач)</span><span class="sxs-lookup"><span data-stu-id="7c693-177">**Start Document** (Content and task pane add-ins only)</span></span>|<span data-ttu-id="7c693-178">Указывает, какой документ следует открыть при запуске проекта.</span><span class="sxs-lookup"><span data-stu-id="7c693-178">Specifies what document to open when you start the project.</span></span>|
|<span data-ttu-id="7c693-179">**Веб-проект**</span><span class="sxs-lookup"><span data-stu-id="7c693-179">**Web Project**</span></span>|<span data-ttu-id="7c693-180">Определяет имя веб-проекта, связанного с надстройкой.</span><span class="sxs-lookup"><span data-stu-id="7c693-180">Specifies the name of the web project associated with the add-in.</span></span>|
|<span data-ttu-id="7c693-181">**Адрес электронной почты** (только надстройки Outlook)</span><span class="sxs-lookup"><span data-stu-id="7c693-181">**Email Address** (Outlook add-ins only)</span></span>|<span data-ttu-id="7c693-182">Указывает адрес электронной почты учетной записи пользователя на сервере Exchange Server или Exchange Online, с которой нужно проверить надстройкой Outlook.</span><span class="sxs-lookup"><span data-stu-id="7c693-182">Specifies the email address of the user account in Exchange Server or Exchange Online that you want to test your Outlook add-in with.</span></span>|
|<span data-ttu-id="7c693-183">**URL-адрес EWS** (только надстройки Outlook)</span><span class="sxs-lookup"><span data-stu-id="7c693-183">**EWS Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="7c693-184">URL веб-службы Exchange (пример: https://www.contoso.com/ews/exchange.aspx).</span><span class="sxs-lookup"><span data-stu-id="7c693-184">Exchange Web service URL (For example: https://www.contoso.com/ews/exchange.aspx).</span></span> |
|<span data-ttu-id="7c693-185">**URL-адрес OWA** (только надстройки Outlook)</span><span class="sxs-lookup"><span data-stu-id="7c693-185">**OWA Url** (Outlook add-ins only)</span></span>|<span data-ttu-id="7c693-186">URL-адрес Outlook Web App (пример: https://www.contoso.com/owa).</span><span class="sxs-lookup"><span data-stu-id="7c693-186">Outlook Web App URL (For example: https://www.contoso.com/owa).</span></span>|
|<span data-ttu-id="7c693-187">**Имя пользователя** (только надстройки Outlook)</span><span class="sxs-lookup"><span data-stu-id="7c693-187">**User name** (Outlook add-ins only)</span></span>|<span data-ttu-id="7c693-188">Указывает имя учетной записи пользователя в Exchange Server или Exchange Online.</span><span class="sxs-lookup"><span data-stu-id="7c693-188">Specifies the name of your user account in Exchange Server or Exchange Online.</span></span>|
|<span data-ttu-id="7c693-189">**Файл проекта**</span><span class="sxs-lookup"><span data-stu-id="7c693-189">**Project File**</span></span>|<span data-ttu-id="7c693-190">Задает имя файла, в котором указаны сборка, конфигурация и другие сведения о проекте.</span><span class="sxs-lookup"><span data-stu-id="7c693-190">Specifies the name of the file containing build, configuration, and other information about the project.</span></span>|
|<span data-ttu-id="7c693-191">**Папка проекта**</span><span class="sxs-lookup"><span data-stu-id="7c693-191">**Project Folder**</span></span>|<span data-ttu-id="7c693-192">Расположение файла проекта.</span><span class="sxs-lookup"><span data-stu-id="7c693-192">The location of the project file.</span></span>|

### <a name="use-an-existing-document-to-debug-the-add-in-content-and-task-pane-add-ins-only"></a><span data-ttu-id="7c693-193">Использование существующего документа для отладки надстройки (только для контентных надстроек и надстроек области задач)</span><span class="sxs-lookup"><span data-stu-id="7c693-193">Use an existing document to debug the add-in (content and task pane add-ins only)</span></span>


<span data-ttu-id="7c693-p116">В проект надстройки можно добавить документы. Если ваш документ содержит тестовые данные, которые необходимо использовать в надстройке, Visual Studio откроет документ при запуске проекта.</span><span class="sxs-lookup"><span data-stu-id="7c693-p116">You can add documents to the add-in project. If you have a document that contains test data that you want to use with your add-in, Visual Studio opens that document for you when you start the project.</span></span>


### <a name="to-use-an-existing-document-to-debug-the-add-in"></a><span data-ttu-id="7c693-196">Чтобы использовать существующий документ для отладки надстройки, выполните следующие действия:</span><span class="sxs-lookup"><span data-stu-id="7c693-196">To use an existing document to debug the add-in</span></span>


1. <span data-ttu-id="7c693-197">В **обозревателе решений** выберите папку проекта надстройки.</span><span class="sxs-lookup"><span data-stu-id="7c693-197">In  **Solution Explorer**, choose the add-in project folder.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="7c693-198">Выберите проект надстройки, а не проект веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="7c693-198">Choose the add-in project and not the web application project.</span></span>

2. <span data-ttu-id="7c693-199">В меню **Проект** выберите пункт **Добавить существующий элемент**.</span><span class="sxs-lookup"><span data-stu-id="7c693-199">On the  **Project** menu, choose **Add Existing Item**.</span></span>
    
3. <span data-ttu-id="7c693-200">В диалоговом окне **Добавление существующего элемента** найдите и выберите документ, который необходимо добавить.</span><span class="sxs-lookup"><span data-stu-id="7c693-200">In the  **Add Existing Item** dialog box, locate and select the document that you want to add.</span></span>
    
4. <span data-ttu-id="7c693-201">Нажмите кнопку **Добавить**, чтобы добавить документ в проект.</span><span class="sxs-lookup"><span data-stu-id="7c693-201">Choose the  **Add** button to add the document to your project.</span></span>
    
5. <span data-ttu-id="7c693-202">В **обозревателе решений** откройте контекстное меню проекта и выберите **Свойства**.</span><span class="sxs-lookup"><span data-stu-id="7c693-202">In  **Solution Explorer**, open the shortcut menu for the project, and then choose  **Properties**.</span></span>
    
    <span data-ttu-id="7c693-203">Появятся страницы свойств для проекта.</span><span class="sxs-lookup"><span data-stu-id="7c693-203">The property pages for the project appear.</span></span>
    
6. <span data-ttu-id="7c693-204">В списке **Документ при запуске** выберите добавляемый в проект документ и нажмите кнопку **ОК**, чтобы закрыть страницы свойств.</span><span class="sxs-lookup"><span data-stu-id="7c693-204">In the  **Start Document** list, choose the document that you added to the project, and then choose the **OK** button to close the property pages.</span></span>
    

### <a name="start-the-solution"></a><span data-ttu-id="7c693-205">Запуск решения</span><span class="sxs-lookup"><span data-stu-id="7c693-205">Start the solution</span></span>


<span data-ttu-id="7c693-p117">При запуске Visual Studio автоматически создает решение. Решение можно запустить в строке **Меню**, выбрав **Отладка**, **Пуск**.</span><span class="sxs-lookup"><span data-stu-id="7c693-p117">Visual Studio will automatically build the solution when you start it. You can start the solution from the  **Menu** bar by choosing **Debug**,  **Start**.</span></span> 


> [!NOTE]
> <span data-ttu-id="7c693-p118">Если отладка скриптов не включена в Internet Explorer, запустить отладчик в Visual Studio не удастся. Чтобы включить отладку, откройте диалоговое окно **Свойства браузера**, перейдите на вкладку **Дополнительно** и снимите флажки **Отключить отладку сценариев (Internet Explorer)** и **Отключить отладку сценариев (другие)**.</span><span class="sxs-lookup"><span data-stu-id="7c693-p118">If script debugging isn't enabled in Internet Explorer, you won't be able to start the debugger in Visual Studio. You can enable script debugging by opening the  **Internet Options** dialog box, choosing the **Advanced** tab, and then clearing the **Disable Script Debugging (Internet Explorer)** and **Disable Script Debugging (Other)** check boxes.</span></span>

<span data-ttu-id="7c693-210">Visual Studio создает проект и выполняет следующие действия:</span><span class="sxs-lookup"><span data-stu-id="7c693-210">Visual Studio builds the project and does the following:</span></span>


1. <span data-ttu-id="7c693-p119">Создает копию XML-файла манифеста и добавляет его в каталог _Имя_проекта_\Output. Ведущее приложение использует эту копию при запуске Visual Studio и отладке надстройки.</span><span class="sxs-lookup"><span data-stu-id="7c693-p119">Creates a copy of the XML manifest file and adds it to  _ProjectName_\Output directory. The host application consumes this copy when you start Visual Studio and debug the add-in.</span></span>
    
2. <span data-ttu-id="7c693-213">Создает набор записей реестра на компьютере, которые служат для включения отображения надстройки в ведущем приложении.</span><span class="sxs-lookup"><span data-stu-id="7c693-213">Creates a set of registry entries on your computer that enable the add-in to appear in the host application.</span></span>
    
3. <span data-ttu-id="7c693-214">Создает проект веб-приложения, а затем развертывает его на локальном веб-сервере IIS (http://localhost)).</span><span class="sxs-lookup"><span data-stu-id="7c693-214">Builds the web application project, and then deploys it to the local IIS web server (http://localhost).</span></span> 
    
<span data-ttu-id="7c693-215">Затем Visual Studio делает следующее:</span><span class="sxs-lookup"><span data-stu-id="7c693-215">Next, Visual Studio does the following:</span></span>


1. <span data-ttu-id="7c693-216">Изменяет элемент [SourceLocation](http://msdn.microsoft.com/en-us/library/e6ea8cd4-7c8b-1da7-d8f8-8d3c80a088bc%28Office.15%29.aspx) файла манифеста XML, заменяя маркер ~ remoteAppUrl полным адресом начальной страницы (например, http://localhost/MyAgave.html)).</span><span class="sxs-lookup"><span data-stu-id="7c693-216">Modifies the SourceLocation element of the XML manifest file by replacing the ~remoteAppUrl token with the fully qualified address of the start page (for example, http://localhost/MyAgave.html).</span></span>
    
2. <span data-ttu-id="7c693-217">Запускает проект веб-приложения в IIS Express.</span><span class="sxs-lookup"><span data-stu-id="7c693-217">Starts the web application project in IIS Express.</span></span>
    
3. <span data-ttu-id="7c693-218">Открывает ведущее приложение.</span><span class="sxs-lookup"><span data-stu-id="7c693-218">Opens the host application.</span></span> 
    
<span data-ttu-id="7c693-p120">Visual Studio не отображает ошибки проверки в окне **ВЫВОД** при построении проекта. Visual Studio сообщает об ошибках и предупреждениях в окне **ОШИБКИ** по мере их возникновения. Кроме того, Visual Studio также выделяет ошибки волнистой линией разных цветов прямо в коде и текстовом редакторе. Эти пометки уведомляют о проблемах, обнаруженных Visual Studio в коде. Дополнительные сведения можно узнать в разделе [Код и текстовый редактор](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx). Дополнительные сведения о включении или отключении проверки изложены в следующих статьях:</span><span class="sxs-lookup"><span data-stu-id="7c693-p120">Visual Studio doesn't show validation errors in the  **OUTPUT** window when you build the project. Visual Studio reports errors and warnings in the **ERRORLIST** window as they occur. Visual Studio also reports validation errors by showing wavy underlines (known as squiggles) of different colors in the code and text editor. These marks notify you of problems that Visual Studio detected in your code. For more information, see [Code and Text Editor](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx). For more information about how to enable or disable validation, see:</span></span> 

- <span data-ttu-id="7c693-225">["Параметры", "Текстовый редактор", JavaScript, IntelliSense](https://msdn.microsoft.com/en-us/library/hh362485(v=vs.140).aspx)</span><span class="sxs-lookup"><span data-stu-id="7c693-225">[Options, Text Editor, JavaScript, IntelliSense](https://msdn.microsoft.com/en-us/library/hh362485(v=vs.140).aspx)</span></span>
    
- <span data-ttu-id="7c693-226">[Практическое руководство. Установка параметров проверки редактирования HTML в Visual Web Developer](https://msdn.microsoft.com/en-us/library/0byxkfet(v=vs.100).aspx)</span><span class="sxs-lookup"><span data-stu-id="7c693-226">[How to: Set Validation Options for HTML Editing in Visual Web Developer](https://msdn.microsoft.com/en-us/library/0byxkfet(v=vs.100).aspx)</span></span>
    
- <span data-ttu-id="7c693-227">[Проверка, CSS, текстовый редактор, диалоговое окно "Параметры"](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx)</span><span class="sxs-lookup"><span data-stu-id="7c693-227">[CSS, see Validation, CSS, Text Editor, Options Dialog Box](https://msdn.microsoft.com/en-us/library/se2f663y(v=vs.140).aspx)</span></span>
    
<span data-ttu-id="7c693-228">Чтобы просмотреть правила проверки XML-файла манифеста проекта, ознакомьтесь с разделом [XML-манифест надстроек для Office](../develop/add-in-manifests.md).</span><span class="sxs-lookup"><span data-stu-id="7c693-228">To review the validation rules of the XML manifest file in your project, see [Office Add-ins XML manifest](../develop/add-in-manifests.md).</span></span>


### <a name="show-an-add-in-in-excel-word-or-project-and-step-through-your-code"></a><span data-ttu-id="7c693-229">Показ надстройки в Excel, Word или Project и пошаговое выполнение кода</span><span class="sxs-lookup"><span data-stu-id="7c693-229">Show an add-in in Excel, Word, or Project and step through your code</span></span>


<span data-ttu-id="7c693-p121">Если для свойства **Документ при запуске** проекта задать значение Excel или Word, Visual Studio создаст новый документ и отобразит надстройку. Если для свойства **Документ при запуске** проекта указать на использование существующего документа, Visual Studio откроет этот документ, но надстройку необходимо вставить вручную. Если для свойства **Документ при запуске** задано значение **Microsoft Project**, также необходимо вставить надстройку вручную.</span><span class="sxs-lookup"><span data-stu-id="7c693-p121">If you set the  **Start Document** property of the add-in project to Excel or Word, Visual Studio creates a new document and the add-in appears. If you set the **Start Document** property of the add-in project to use an existing document, Visual Studio opens the document, but you have to insert the add-in manually. If you set the **Start Document** to **Microsoft Project**, you also have to insert the add-in manually.</span></span>


### <a name="to-show-an-office-add-in-in-excel-or-word"></a><span data-ttu-id="7c693-233">Отображение Надстройка Office в Excel или Word</span><span class="sxs-lookup"><span data-stu-id="7c693-233">To show an Office Add-in in Excel or Word</span></span>


1. <span data-ttu-id="7c693-234">В Excel или Word на вкладке **Вставка** выберите **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="7c693-234">In Excel or Word, on the  **Insert** tab, choose **Office Add-ins**.</span></span>
    
2. <span data-ttu-id="7c693-235">Выберите надстройку в появившемся списке.</span><span class="sxs-lookup"><span data-stu-id="7c693-235">In the list that appears, choose your add-in.</span></span>
    

### <a name="to-show-an-office-add-in-in-project"></a><span data-ttu-id="7c693-236">Отображение Надстройка Office в Project</span><span class="sxs-lookup"><span data-stu-id="7c693-236">To show an Office Add-in in Project</span></span>


1. <span data-ttu-id="7c693-237">В Project на вкладке **Проект** выберите **Надстройки Office**.</span><span class="sxs-lookup"><span data-stu-id="7c693-237">In Project, on the  **Project** tab, choose **Office Add-ins**.</span></span>
    
2. <span data-ttu-id="7c693-238">Выберите надстройку в появившемся списке.</span><span class="sxs-lookup"><span data-stu-id="7c693-238">In the list that appears, choose your add-in.</span></span>
    
<span data-ttu-id="7c693-p122">В Visual Studio теперь можно создавать точки останова, а затем взаимодействовать с надстройкой и выполнять код в файлах HTML, JavaScript и C# или VB в пошаговом режиме.</span><span class="sxs-lookup"><span data-stu-id="7c693-p122">In Visual Studio, you can then set break-points. Then, as you interact with your add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span>


### <a name="show-the-outlook-add-in-in-outlook-and-step-through-your-code"></a><span data-ttu-id="7c693-241">Отображение надстройки Outlook в Outlook и пошаговый переход по коду</span><span class="sxs-lookup"><span data-stu-id="7c693-241">Show the Outlook add-in in Outlook and step through your code</span></span>


<span data-ttu-id="7c693-242">Чтобы просмотреть надстройку Outlook, откройте сообщение эл. почты или элемент встречи.</span><span class="sxs-lookup"><span data-stu-id="7c693-242">To view the add-in in Outlook, open an email message or appointment item.</span></span>

<span data-ttu-id="7c693-p123">Outlook активирует надстройка для этого элемента, если соблюдаются критерии активации. В верхней части окна инспектора или области чтения появляется строка надстройка, и надстройка Outlook отображается в строке надстройка в виде кнопки. Если в вашей надстройке есть команда, на ленте появится кнопка (либо на вкладке по умолчанию, либо на пользовательской вкладке), а надстройка не будет отображаться в области надстройка.</span><span class="sxs-lookup"><span data-stu-id="7c693-p123">Outlook activates the add-in for the item as long as the activation criteria are met. The add-in bar appears at the top of the Inspector window or Reading Pane, and your Outlook add-in appears as a button in the add-in bar. If your add-in has an add-in command, a button will appear in the ribbon, either in the default tab or a specified custom tab, and the add-in will not appear in the add-in bar.</span></span>

<span data-ttu-id="7c693-246">Чтобы просмотреть надстройку Outlook, нажмите соответствующую кнопку.</span><span class="sxs-lookup"><span data-stu-id="7c693-246">To view your Outlook add-in, choose the button for your Outlook add-in.</span></span>

<span data-ttu-id="7c693-p124">В Visual Studio можно создавать точки останова, а затем взаимодействовать с надстройкой Outlook и выполнять код в файлах HTML, JavaScript и C# или VB в пошаговом режиме.</span><span class="sxs-lookup"><span data-stu-id="7c693-p124">In Visual Studio, you can set break-points. Then, as you interact with your Outlook add-in and step through the code in your HTML, JavaScript, and C# or VB code files.</span></span> 

<span data-ttu-id="7c693-p125">Можно также изменить код и просмотреть результаты этих изменений в надстройке Outlook без необходимости закрывать Надстройка Office и повторного запуска проекта. В Outlook следует просто открыть контекстное меню для надстройки Outlook и нажать кнопку **Обновить**.</span><span class="sxs-lookup"><span data-stu-id="7c693-p125">You can also change your code and review the effects of those changes in your Outlook add-in without having to close the Office Add-in and start the project again. In Outlook, just open the shortcut menu for the Outlook add-in, and then choose  **Reload**.</span></span>


### <a name="modify-code-and-continue-to-debug-the-add-in-without-having-to-start-the-project-again"></a><span data-ttu-id="7c693-251">Изменение кода и продолжение отладки надстройки без необходимости повторного запуска проекта</span><span class="sxs-lookup"><span data-stu-id="7c693-251">Modify code and continue to debug the add-in without having to start the project again</span></span>


<span data-ttu-id="7c693-p126">Код можно изменить и просмотреть результаты выполненных изменений в вашей надстройке, не закрывая ведущее приложение и не запуская проект снова. После изменения кода откройте контекстное меню надстройки и выберите команду **Перезагрузить**. Когда вы перезагружаете приложение, оно отключается от отладчика Visual Studio. Поэтому можно просмотреть эффект изменений, но невозможно пошагово выполнить код, пока ко всем доступным процессам Iexplore.exe не будет подключен отладчик Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="7c693-p126">You can change your code and review the effects of those changes in your add-in without having to close the host application and start the project again. After you change your code, open the shortcut menu for the add-in, and then choose  **Reload**. When you reload the add-in it becomes disconnected with the Visual Studio debugger. Therefore, you can view the effects of your change, but you cannot step through your code again until you attach the Visual Studio debugger to all of the available Iexplore.exe processes.</span></span>


### <a name="to-attach-the-visual-studio-debugger-to-all-of-the-available-iexploreexe-processes"></a><span data-ttu-id="7c693-256">Подключение отладчика Visual Studio ко всем доступным процессам Iexplore.exe</span><span class="sxs-lookup"><span data-stu-id="7c693-256">To attach the Visual Studio debugger to all of the available Iexplore.exe processes</span></span>


1. <span data-ttu-id="7c693-257">В Visual Studio выберите команды **ОТЛАДКА**, **Присоединиться к процессу**.</span><span class="sxs-lookup"><span data-stu-id="7c693-257">In Visual Studio, choose  **DEBUG**,  **Attach to Process**.</span></span>
    
2. <span data-ttu-id="7c693-258">В диалоговом окне **Присоединение к процессу** выберите все доступные процессы **Iexplore.exe**, а затем нажмите кнопку **Присоединиться**.</span><span class="sxs-lookup"><span data-stu-id="7c693-258">In the  **Attach to Process** dialog box, choose all of the available **Iexplore.exe** processes, and then choose the **Attach** button.</span></span>
    

## <a name="next-steps"></a><span data-ttu-id="7c693-259">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="7c693-259">Next steps</span></span>

- [<span data-ttu-id="7c693-260">Развертывание и публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="7c693-260">Deploy and publish your Office Add-in</span></span>](../publish/publish.md)
    
