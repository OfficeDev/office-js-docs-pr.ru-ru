---
title: Надстройки области задач для Project
description: Узнайте о надстройках панели задач для Project.
ms.date: 09/26/2019
ms.topic: conceptual
ms.custom: scenarios:getting-started
localization_priority: Priority
ms.openlocfilehash: 1e471c53e39af8764840716d59a4d26719d3ac0a
ms.sourcegitcommit: 9609bd5b4982cdaa2ea7637709a78a45835ffb19
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/28/2020
ms.locfileid: "47292407"
---
# <a name="task-pane-add-ins-for-project"></a><span data-ttu-id="72793-103">Надстройки области задач для Project</span><span class="sxs-lookup"><span data-stu-id="72793-103">Task pane add-ins for Project</span></span>

<span data-ttu-id="72793-104">Project стандартный 2013 и Project профессиональный 2013 (версии 15.1 или более поздней) поддерживают надстройки области задач. Вы можете запускать стандартные надстройки области задач, разработанные для Word и Excel.</span><span class="sxs-lookup"><span data-stu-id="72793-104">Project Standard 2013 and Project Professional 2013 (version 15.1 or higher) both include support for task pane add-ins. You can run general task pane add-ins that are developed for Word or Excel.</span></span> <span data-ttu-id="72793-105">Вы также можете разрабатывать собственные надстройки, которые обрабатывают события выбора в Project и интегрируют данные задачи, ресурса, представления и другие данные уровня ячейки в проект со списками SharePoint, надстройками SharePoint, веб-частями, веб-службами и корпоративными приложениями.</span><span class="sxs-lookup"><span data-stu-id="72793-105">You can also develop custom add-ins that handle selection events in Project and integrate task, resource, view, and other cell-level data in a project with SharePoint lists, SharePoint Add-ins, Web Parts, web services, and enterprise applications.</span></span>

> [!NOTE]
> <span data-ttu-id="72793-p102">[Загружаемый пакет SDK Project 2013](https://www.microsoft.com/download/details.aspx?id=30435%20) включает в себя примеры надстроек, показывающие, как использовать объектную модель надстроек для Project и как использовать службу OData для отчетности в Project Server 2013. При извлечении и установке пакета SDK см. подкаталог `\Samples\Apps\`.</span><span class="sxs-lookup"><span data-stu-id="72793-p102">The [Project 2013 SDK download](https://www.microsoft.com/download/details.aspx?id=30435%20) includes sample add-ins that show how to use the add-in object model for Project, and how to use the OData service for reporting data in Project Server 2013. When you extract and install the SDK, see the `\Samples\Apps\` subdirectory.</span></span>

<span data-ttu-id="72793-108">Общие сведения о надстройках Office см. в статье [Обзор платформы надстроек Office](../overview/office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="72793-108">For an introduction to Office Add-ins, see [Office Add-ins platform overview](../overview/office-add-ins.md).</span></span>

## <a name="add-in-scenarios-for-project"></a><span data-ttu-id="72793-109">Сценарии надстроек для Project</span><span class="sxs-lookup"><span data-stu-id="72793-109">Add-in scenarios for Project</span></span>

<span data-ttu-id="72793-p103">Руководители проектов могут использовать надстройки области задач Project, чтобы упростить управление проектом. Вместо переключения из Project и открытия другого приложения для поиска часто используемой информации, руководители проектов могут осуществлять прямой доступ к этой информации в Project. Контент в надстройке области задач может быть контекстно-зависимым на основании выбранной задачи, ресурсов, представления или других данных из ячейки на диаграмме Ганта, в представлении использования задач или представлении использования ресурсов.</span><span class="sxs-lookup"><span data-stu-id="72793-p103">Project managers can use Project task pane add-ins to help with project management activities. Instead of leaving Project and opening another application to search for frequently used information, project managers can directly access the information within Project. The content in a task pane add-in can be context-sensitive, based on the selected task, resource, view, or other data in a cell in a Gantt chart, task usage view, or resource usage view.</span></span>

> [!NOTE]
> <span data-ttu-id="72793-113">С помощью Project профессиональный 2013 вы можете разрабатывать надстройки области задач, которые получают доступ к Project в Интернете, локальным установкам Project Server 2013, а также SharePoint 2013 в локальной среде и в сети.</span><span class="sxs-lookup"><span data-stu-id="72793-113">With Project Professional 2013, you can develop task pane add-ins that access Project on the web, on-premises installations of Project Server 2013, and on-premises or online SharePoint 2013.</span></span> <span data-ttu-id="72793-114">Project стандартный 2013 не поддерживает прямую интеграцию с данными Project Server или списками задач SharePoint, синхронизированными с Project Server.</span><span class="sxs-lookup"><span data-stu-id="72793-114">Project Standard 2013 does not support direct integration with Project Server data or SharePoint task lists that are synchronized with Project Server.</span></span>

<span data-ttu-id="72793-115">Возможны следующие сценарии использования надстроек для Project.</span><span class="sxs-lookup"><span data-stu-id="72793-115">Add-in scenarios for Project include the following:</span></span>

- <span data-ttu-id="72793-p105">**Составление графика проекта**. Просматривайте данные из связанных проектов, которые могут затрагивать график. Надстройка области задач может интегрировать необходимые данные из других проектов в Project Server 2013. Например, вы можете просматривать наборы проектов и даты этапов разработки для подразделений или просматривать данные на определенную дату из других проектов, основанных на выбранном настраиваемом поле.</span><span class="sxs-lookup"><span data-stu-id="72793-p105">**Project scheduling** View data from related projects that can affect scheduling. A task pane add-in can integrate relevant data from other projects in Project Server 2013. For example, you can view the departmental collection of projects and milestone dates, or view specified data from other projects that are based on a selected custom field.</span></span>

- <span data-ttu-id="72793-119">**Управление ресурсами**. Просматривайте полный пул ресурсов в Project Server 2013 или подмножество, основанное на указанных навыках, включая данные о затратах и доступность ресурсов, чтобы подобрать необходимые ресурсы.</span><span class="sxs-lookup"><span data-stu-id="72793-119">**Resource management** View the complete resource pool in Project Server 2013 or a subset based on specified skills, including cost data and resource availability, to help select appropriate resources.</span></span>

- <span data-ttu-id="72793-p106">**Определение состояния и утверждения**. Используйте веб-приложение в надстройке области задач, чтобы обновить или просмотреть данные из внешнего приложения планирования ресурсов предприятия (ERP), системы управления расписаниями или приложения учета. Либо создайте настраиваемую веб-часть утверждения состояния, которую можно использовать как в Project Web App, так и в Project профессиональный 2013.</span><span class="sxs-lookup"><span data-stu-id="72793-p106">**Statusing and approvals** Use a web application in a task pane add-in to update or view data from an external enterprise resource planning (ERP) application, timesheet system, or accounting application. Or, create a custom status approval Web Part that can be used within both Project Web App and Project Professional 2013.</span></span>

- <span data-ttu-id="72793-p107">**Общение в группе**. Взаимодействуйте с членами команды и ресурсами непосредственно из надстройки области задач в рамках контекста проекта. Либо с легкостью ведите для себя контекстно-зависимые заметки по мере работы над проектом.</span><span class="sxs-lookup"><span data-stu-id="72793-p107">**Team communication** Communicate with team members and resources directly from a task pane add-in, within the context of a project. Or, easily maintain a set of context-sensitive notes for yourself as you work in a project.</span></span>

- <span data-ttu-id="72793-p108">**Рабочие пакеты**. Выполняйте поиск определенных видов шаблонов проектов в библиотеках SharePoint и коллекциях шаблонов в Интернете. Например, выполняйте поиск шаблонов для строительных проектов и добавляйте их в коллекцию шаблонов Project.</span><span class="sxs-lookup"><span data-stu-id="72793-p108">**Work packages** Search for specified kinds of project templates within SharePoint libraries and online template collections. For example, find templates for construction projects and add them to your Project template collection.</span></span>

- <span data-ttu-id="72793-p109">**Связанные элементы**. Просматривайте метаданные, документы и сообщения, связанные с определенными задачами в плане проекта. Например, вы можете использовать Project профессиональный 2013 для управления проектом, импортированным из списка задач SharePoint, и одновременно синхронизировать этот список задач с изменениями в проекте. Надстройка области задач может отображать дополнительные поля или метаданные, которые не были импортированы Project для задач в списке SharePoint.</span><span class="sxs-lookup"><span data-stu-id="72793-p109">**Related items** View metadata, documents, and messages that are related to specific tasks in a project plan. For example, you can use Project Professional 2013 to manage a project that was imported from a SharePoint task list, and still synchronize the task list with changes in the project. A task pane add-in can show additional fields or metadata that Project did not import for tasks in the SharePoint list.</span></span>

- <span data-ttu-id="72793-p110">**Использование объектных моделей Project Server**. Используйте GUID выбранной задачи с методами в интерфейсе Project Server (PSI) или клиентской объектной модели (CSOM) Project Server. Например, веб-приложение для надстройки может считывать и обновлять данные определения состояния для выбранной задачи или выбранного ресурса либо осуществлять интеграцию с внешним приложением управления расписаниями.</span><span class="sxs-lookup"><span data-stu-id="72793-p110">**Use the Project Server object models** Use the GUID of a selected task with methods in the Project Server Interface (PSI) or the client-side object model (CSOM) of Project Server. For example, the web application for an add-in can read and update the statusing data of a selected task and resource, or integrate with an external timesheet application.</span></span>

- <span data-ttu-id="72793-p111">**Получение данных отчетов**. Используйте запросы REST, JavaScript или LINQ, чтобы найти связанные сведения для выбранной задачи или выбранного ресурса в службе OData для отчетных таблиц в Project Web App. Запросы, использующие службу OData, можно выполнять с помощью интернет-версии или локальной установки Project Server 2013.</span><span class="sxs-lookup"><span data-stu-id="72793-p111">**Get reporting data** Use Representational State Transfer (REST), JavaScript, or LINQ queries to find related information for a selected task or resource in the OData service for reporting tables in Project Web App. Queries that use the OData service can be done with an online or an on-premises installation of Project Server 2013.</span></span>

    <span data-ttu-id="72793-133">Пример представлен в статье [Создание надстройки Project, использующей REST с локальной службой OData Project Server](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).</span><span class="sxs-lookup"><span data-stu-id="72793-133">For example, see [Create a Project add-in that uses REST with an on-premises Project Server OData  service](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).</span></span>

## <a name="developing-project-add-ins"></a><span data-ttu-id="72793-134">Разработка надстроек для Project</span><span class="sxs-lookup"><span data-stu-id="72793-134">Developing Project add-ins</span></span>

<span data-ttu-id="72793-p112">Библиотека JavaScript для надстройки для Project содержит расширения псевдонима пространства имен **Office**, позволяющие разработчикам осуществлять доступ к свойствам приложения Project, а также задач, ресурсов и представлений в проекте. Расширения библиотеки JavaScript в файле Project-15.js используются в надстройке для Project, созданной с использованием Visual Studio 2015. Office.js, Office.debug.js, Project-15.js, Project-15.debug.js и другие связанные файлы также предоставлены в загружаемом пакете SDK Project 2013.</span><span class="sxs-lookup"><span data-stu-id="72793-p112">The JavaScript library for Project add-ins includes extensions of the **Office** namespace alias that enable developers to access properties of the Project application and tasks, resources, and views in a project. The JavaScript library extensions in the Project-15.js file are used in a Project add-in created with Visual Studio 2015. The Office.js, Office.debug.js, Project-15.js, Project-15.debug.js, and related files are also provided in the Project 2013 SDK download.</span></span>

<span data-ttu-id="72793-p113">Чтобы создать надстройку, вы можете использовать простой текстовый редактор для создания веб-страницы HTML, а также связанных файлов JavaScript, файлов CSS и запросов REST. Кроме HTML-страницы или веб-приложения, для конфигурации надстройки требуется XML-файл манифеста. Project может использовать файл манифеста, который включает в себя атрибут **type**, указанный в виде **TaskPaneExtension**. Файл манифеста может использоваться несколькими клиентскими приложениями Office 2013, либо вы можете создать файл манифеста специально для Project 2013. Подробнее см. в разделе _Основы разработки_ статьи [Обзор платформы надстроек Office](../overview/office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="72793-p113">To create an add-in, you can use a simple text editor to create an HTML webpage and related JavaScript files, CSS files, and REST queries. In addition to an HTML page or a web application, an add-in requires an XML manifest file for configuration. Project can use a manifest file that includes a **type** attribute that is specified as **TaskPaneExtension**. The manifest file can be used by multiple Office 2013 client applications, or you can create a manifest file that is specific for Project 2013. For more information, see the  _Development basics_ section in [Office Add-ins platform overview](../overview/office-add-ins.md).</span></span>

<span data-ttu-id="72793-143">При установке загружаемого пакета SDK для Project 2013 подкаталог `\Samples\Apps\` содержит следующие примеры надстроек:</span><span class="sxs-lookup"><span data-stu-id="72793-143">When you install the Project 2013 SDK download, the  `\Samples\Apps\` subdirectory includes the following sample add-ins:</span></span>

- <span data-ttu-id="72793-p114">**Поиск Bing**. Файл манифеста BingSearch.xml указывает на страницу поиска Bing для мобильных устройств. Поскольку в Интернете уже присутствует веб-приложение Bing, надстройка поиска Bing не использует другие файлы исходного кода или объектную модель надстроек для Project.</span><span class="sxs-lookup"><span data-stu-id="72793-p114">**Bing Search:** The BingSearch.xml manifest file points to the Bing search page for mobile devices. Because the Bing web app already exists on the Internet, the Bing Search add-in does not use other source code files or the add-in object model for Project.</span></span>

- <span data-ttu-id="72793-p115">**Тест объектной модели Project**. Вместе файл манифеста JSOM_SimpleOMCalls.xml и файл JSOM_Call.html представляют собой пример, тестирующий объектную модель и функциональные возможности надстройки в Project 2013. HTML-файл ссылается на файл JSOM_Sample.js, функции JavaScript которого используют файл Office.js и файлы Project-15.js для реализации основных функциональных возможностей. Загружаемый пакет SDK включает в себя все необходимые файлы исходного кода и XML-файл манифеста надстройки теста объектной модели Project. Разработка и установка примера теста объектной модели Project описана в статье [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span><span class="sxs-lookup"><span data-stu-id="72793-p115">**Project OM Test:** The JSOM_SimpleOMCalls.xml manifest file and the JSOM_Call.html file are, together, an example that tests the object model and add-in functionality in Project 2013. The HTML file references the JSOM_Sample.js file, which has JavaScript functions that use the Office.js file and the Project-15.js file for the primary functionality. The SDK download includes all of the necessary source code files and the manifest XML file for the Project OM Test add-in. The development and installation of the Project OM Test sample is described in [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>

- <span data-ttu-id="72793-p116">**HelloProject_OData**. Это решение Visual Studio для Project профессиональный 2013, которое формирует сводные данные по активному проекту, например, сведения о стоимости, работе и проценте завершения, а также сравнивает их со средними показателями для всех опубликованных проектов в том экземпляре Project Web App, где хранится активный проект. Разработка, установка и тестирование примера, использующего протокол REST в службе  **ProjectData** в Project Web App, описаны в статье [Создание надстройки Project, использующей REST с локальной службой OData Project Server](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).</span><span class="sxs-lookup"><span data-stu-id="72793-p116">**HelloProject_OData:** This is a Visual Studio solution for Project Professional 2013 that summarizes data from the active project, such as cost, work, and percent complete, and compares that with the average for all published projects in the Project Web App instance where the active project is stored. The development, installation, and testing of the sample, which uses the REST protocol with the **ProjectData** service in Project Web App, is described in [Create a Project add-in that uses REST with an on-premises Project Server OData service](../project/create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md).</span></span>

### <a name="creating-an-add-in-manifest-file"></a><span data-ttu-id="72793-152">Создание файла манифеста надстройки</span><span class="sxs-lookup"><span data-stu-id="72793-152">Creating an add-in manifest file</span></span>

<span data-ttu-id="72793-153">Файл манифеста указывает URL-адрес веб-страницы надстройки или веб-приложения, вид надстройки (надстройка области задач для Project), дополнительные URL-адреса контента для других языков и региональных параметров и другие свойства.</span><span class="sxs-lookup"><span data-stu-id="72793-153">The manifest file specifies the URL of the add-in webpage or web application, the kind of add-in (task pane for Project), optional URLs of content for other languages and locales, and other properties.</span></span>

### <a name="procedure-1-to-create-the-add-in-manifest-file-for-bing-search"></a><span data-ttu-id="72793-p117">Процедура 1. Создание файла манифеста для надстройки поиска Bing</span><span class="sxs-lookup"><span data-stu-id="72793-p117">Procedure 1. To create the add-in manifest file for Bing Search</span></span>

- <span data-ttu-id="72793-p118">Создайте XML-файл в локальном каталоге. Этот XML-файл включает в себя элемент **OfficeApp** и дочерние элементы, описанные в статье [XML-манифест надстроек для Office](../develop/add-in-manifests.md). Например, создайте файл BingSearch.xml со следующим XML-кодом.</span><span class="sxs-lookup"><span data-stu-id="72793-p118">Create an XML file in a local directory. The XML file includes the **OfficeApp** element and child elements, which are described in the [Office Add-ins XML manifest](../develop/add-in-manifests.md). For example, create a file named BingSearch.xml that contains the following XML.</span></span>

    ```XML
    <?xml version="1.0" encoding="utf-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
               xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
               xsi:type="TaskPaneApp">
      <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
      <Id>01234567-89ab-cedf-0123-456789abcdef</Id>
      <Version>15.0</Version>
      <ProviderName>Microsoft</ProviderName>
      <DefaultLocale>en-us</DefaultLocale>
      <DisplayName DefaultValue="Bing Search">
      </DisplayName>
      <Description DefaultValue="Search selected data on Bing">
      </Description>
      <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
      </IconUrl>
      <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
      <Capabilities>
        <Capability Name="Project"/>
      </Capabilities>
      <DefaultSettings>
        <SourceLocation DefaultValue="http://m.bing.com">
        </SourceLocation>
      </DefaultSettings>
      <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

- <span data-ttu-id="72793-159">Ниже приведены обязательные элементы манифеста надстройки:</span><span class="sxs-lookup"><span data-stu-id="72793-159">Following are the required elements in the add-in manifest:</span></span>
  - <span data-ttu-id="72793-160">Атрибут `xsi:type="TaskPaneApp"` в элементе **OfficeApp** указывает, что эта надстройка относится к типу области задач.</span><span class="sxs-lookup"><span data-stu-id="72793-160">In the **OfficeApp** element, the `xsi:type="TaskPaneApp"` attribute specifies that the add-in is a task pane type.</span></span>
  - <span data-ttu-id="72793-161">Элемент **Id** является UUID и должен быть уникальным.</span><span class="sxs-lookup"><span data-stu-id="72793-161">The **Id** element is a UUID and must be unique.</span></span>
  - <span data-ttu-id="72793-p119">Элемент **Version** указывает версию надстройки. Элемент **ProviderName** указывает название компании или имя разработчика, предоставивших надстройку. Элемент **DefaultLocale** указывает язык и региональные параметры по умолчанию для строк манифеста.</span><span class="sxs-lookup"><span data-stu-id="72793-p119">The **Version** element is the version of the add-in. The **ProviderName** element is the name of the company or developer who provides the add-in. The **DefaultLocale** element specifies the default locale for the strings in the manifest.</span></span>
  - <span data-ttu-id="72793-p120">Элемент **DisplayName** представляет собой имя, отображаемое в раскрывающемся списке **Надстройка области задач** на вкладке **Вид** ленты Project 2013. Это значение может содержать до 32 символов.</span><span class="sxs-lookup"><span data-stu-id="72793-p120">The **DisplayName** element is the name that shows in the **Task Pane Add-in** drop-down list in the **VIEW** tab of the ribbon in Project 2013. The value can contain up to 32 characters.</span></span>
  - <span data-ttu-id="72793-p121">Элемент **Description** содержит описание надстройки на языке по умолчанию. Это значение может содержать до 2000 символов.</span><span class="sxs-lookup"><span data-stu-id="72793-p121">The **Description** element contains the add-in description for the default locale. The value can contain up to 2000 characters.</span></span>
  - <span data-ttu-id="72793-169">Элемент **Capabilities** содержит один или несколько дочерних элементов **Capability**, указывающих приложение Office.</span><span class="sxs-lookup"><span data-stu-id="72793-169">The **Capabilities** element contains one or more **Capability** child elements that specify the Office application.</span></span>
  - <span data-ttu-id="72793-p122">Элемент **DefaultSettings** включает в себя элемент **SourceLocation**, который указывает путь к HTML-файлу в общей папке или URL-адрес веб-страницы, используемой надстройкой. Надстройка области задач игнорирует элементы **RequestedHeight** и **RequestedWidth**.</span><span class="sxs-lookup"><span data-stu-id="72793-p122">The **DefaultSettings** element includes the **SourceLocation** element, which specifies the path of an HTML file on a file share or the URL of a webpage that the add-in uses. A task pane add-in ignores the **RequestedHeight** element and the **RequestedWidth** element.</span></span>
  - <span data-ttu-id="72793-p123">Элемент **IconUrl** является необязательным. Он может быть значком в общей папке или URL-адресом значка в веб-приложении.</span><span class="sxs-lookup"><span data-stu-id="72793-p123">The **IconUrl** element is optional. It can be an icon on a file share or the URL of an icon in a web application.</span></span>

- <span data-ttu-id="72793-p124">(Необязательно) Добавьте элементы **Override**, имеющие значения для других региональных параметров и языка. Например, следующий манифест предоставляет элементы **Override** для значений **DisplayName**, **Description**, **IconUrl** и **SourceLocation** для французского языка.</span><span class="sxs-lookup"><span data-stu-id="72793-p124">(Optional) Add **Override** elements that have values for other locales. For example, the following manifest provides **Override** elements for French values of **DisplayName**, **Description**, **IconUrl**, and **SourceLocation**.</span></span>

    ```XML
    <?xml version="1.0" encoding="utf-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.0"
                xmlns:xsi="https://www.w3.org/2001/XMLSchema-instance"
              xsi:type="TaskPaneApp">
      <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
      <Id>01234567-89ab-cedf-0123-456789abcdef</Id>
      <Version>15.0</Version>
      <ProviderName>Microsoft</ProviderName>
      <DefaultLocale>en-us</DefaultLocale>
      <DisplayName DefaultValue="Bing Search">
        <Override Locale="fr-fr" Value="Bing Search"/>
      </DisplayName>
      <Description DefaultValue="Search selected data on Bing">
        <Override Locale="fr-fr" Value="Search selected data on Bing"></Override>
      </Description>
      <IconUrl DefaultValue="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg">
        <Override Locale="fr-fr" Value="http://officeimg.vo.msecnd.net/_layouts/images/general/office_logo.jpg"/>
      </IconUrl>
      <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
      <Capabilities>
        <Capability Name="Project"/>
      </Capabilities>
      <DefaultSettings>
        <SourceLocation DefaultValue="http://m.bing.com">
          <Override Locale="fr-fr" Value="http://m.bing.com"/>
        </SourceLocation>
      </DefaultSettings>
      <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

## <a name="installing-project-add-ins"></a><span data-ttu-id="72793-176">Установка надстроек Project</span><span class="sxs-lookup"><span data-stu-id="72793-176">Installing Project add-ins</span></span>

<span data-ttu-id="72793-p125">В Project 2013 вы можете устанавливать надстройки в виде изолированных решений в общей папке или в частном каталоге надстроек. Вы также можете просматривать и приобретать надстройки в AppSource.</span><span class="sxs-lookup"><span data-stu-id="72793-p125">In Project 2013, you can install add-ins as stand-alone solutions on a file share, or in a private add-in catalog. You can also review and purchase add-ins in AppSource.</span></span>

<span data-ttu-id="72793-p126">Общая папка может содержать несколько XML-файлов манифестов и несколько подкаталогов. Вы можете добавлять и удалять каталоги для хранения манифестов с помощью вкладки **Надежные каталоги надстроек** диалогового окна **Центр управления безопасностью** в Project 2013. Для отображения надстройки в Project элемент **SourceLocation** в манифесте должен указывать на существующий веб-сайт или исходный HTML-файл.</span><span class="sxs-lookup"><span data-stu-id="72793-p126">There can be multiple add-in manifest XML files and subdirectories in a file share. You can add or remove manifest directory locations and catalogs by using the **Trusted Add-in Catalogs** tab in the **Trust Center** dialog box in Project 2013. To show an add-in in Project, the **SourceLocation** element in a manifest must point to an existing website or HTML source file.</span></span>

> [!NOTE]
> <span data-ttu-id="72793-182">Если вы занимаетесь разработкой на компьютере с Windows, необходимо установить Internet Explorer или Microsoft Edge.</span><span class="sxs-lookup"><span data-stu-id="72793-182">If you are developing on a Windows computer, either Internet Explorer or Microsoft Edge must be installed.</span></span> <span data-ttu-id="72793-183">Дополнительные сведения см. в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="72793-183">For more information see [Browsers used by Office Add-ins](../concepts/browsers-used-by-office-web-add-ins.md).</span></span>

<span data-ttu-id="72793-p128">В процедуре 2 надстройка поиска Bing устанавливается на локальный компьютер с установленным Project 2013. Однако из-за того, что инфраструктура надстроек не использует локальные пути для файлов, такие как  `C:\Project\AppManifests`, вы можете создать на локальном компьютере сетевую папку. При необходимости вы можете создать общую папку на удаленном компьютере.</span><span class="sxs-lookup"><span data-stu-id="72793-p128">In Procedure 2, the Bing Search add-in is installed on the local computer where Project 2013 is installed. However, because the add-in infrastructure does not directly use local file paths such as  `C:\Project\AppManifests`, you can create a network share on the local computer. If you prefer, you can create a file share on a remote computer.</span></span>

### <a name="procedure-2-to-install-the-bing-search-add-in"></a><span data-ttu-id="72793-p129">Процедура 2. Установка надстройки поиска Bing</span><span class="sxs-lookup"><span data-stu-id="72793-p129">Procedure 2. To install the Bing Search add-in</span></span>

1. <span data-ttu-id="72793-p130">Создайте локальный каталог для манифестов надстроек. Например, создайте каталог  `C:\Project\AppManifests`.</span><span class="sxs-lookup"><span data-stu-id="72793-p130">Create a local directory for add-in manifests. For example, create the  `C:\Project\AppManifests` directory.</span></span>

2. <span data-ttu-id="72793-191">Разрешите для каталога  `C:\Project\AppManifests` общий доступ с именем AppManifests, чтобы сетевой путь к общей папке выглядел следующим образом: `\\ServerName\AppManifests`.</span><span class="sxs-lookup"><span data-stu-id="72793-191">Share the  `C:\Project\AppManifests` directory asAppManifests, so the network path to the file share becomes  `\\ServerName\AppManifests`.</span></span>

3. <span data-ttu-id="72793-192">Скопируйте файл манифеста BingSearch.xml в каталог  `C:\Project\AppManifests`.</span><span class="sxs-lookup"><span data-stu-id="72793-192">Copy the BingSearch.xml manifest file to the  `C:\Project\AppManifests` directory.</span></span>

4. <span data-ttu-id="72793-193">В Project 2013 откройте диалоговое окно **Параметры Project**, выберите **Центр управления безопасностью** и затем **Параметры центра управления безопасностью**.</span><span class="sxs-lookup"><span data-stu-id="72793-193">In Project 2013, open the **Project Options** dialog box, choose **Trust Center**, and then choose **Trust Center Settings**.</span></span>

5. <span data-ttu-id="72793-194">В диалоговом окне **Центр управления безопасностью** выберите в левой области **Доверенные каталоги надстроек**.</span><span class="sxs-lookup"><span data-stu-id="72793-194">In the **Trust Center** dialog box, in the left pane, choose **Trusted Add-in Catalogs**.</span></span>

6. <span data-ttu-id="72793-195">В области **Надежные каталоги надстроек** (см. рис. 1) добавьте путь `\\ServerName\AppManifests` в текстовое поле **URL-адрес каталога**, выберите элемент **Добавить каталог** и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="72793-195">In the **Trusted Add-in Catalogs** pane (see Figure 1), add the `\\ServerName\AppManifests` path in the **Catalog Url** text box, choose **Add Catalog**, and then choose **OK**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="72793-p131">На рис. 1 показаны две общие папки и один гипотетический URL-адрес частного каталога в списке **Адрес доверенного каталога**. Только одна общая папка может использоваться по умолчанию, и только один URL-адрес может указывать каталог по умолчанию. Например, если сделать `\\Server2\AppManifests` папкой по умолчанию, Project снимет флажок **По умолчанию** для папки `\\ServerName\AppManifests`. Если выделение по умолчанию будет изменено, вы можете нажать кнопку **Очистить**, чтобы удалить установленные надстройки, а затем перезапустить Project. Если вы добавите надстройку в общую папку по умолчанию или каталог SharePoint при запущенном приложении Project, необходимо перезапустить Project.</span><span class="sxs-lookup"><span data-stu-id="72793-p131">Figure 1 shows two file shares and one hypothetical URL for a private catalog in the **Trusted Catalog Address** list. Only one file share can be the default file share and only one catalog URL can be the default catalog. For example, if you set `\\Server2\AppManifests` as the default, Project clears the **Default** check box for `\\ServerName\AppManifests`.If you change the default selection, you can choose **Clear** to remove installed add-ins, and then restart Project. If you add an add-in to the default file share or SharePoint catalog while Project is open, you should restart Project.</span></span>

    <span data-ttu-id="72793-200">*Рис. 1. Добавление каталогов с манифестами надстроек с помощью центра управления безопасностью*</span><span class="sxs-lookup"><span data-stu-id="72793-200">*Figure 1. Using the Trust Center to add catalogs of add-in manifests*</span></span>

    ![Использование центра управления безопасностью для добавления манифестов приложений](../images/pj15-agave-overview-trust-centers.png)

7. <span data-ttu-id="72793-p132">На ленте **Project** выберите раскрывающееся меню **Apps Надстройки Office** и выберите элемент **Просмотреть все**. В диалоговом окне **Вставка надстройки** выберите **ОБЩАЯ ПАПКА** (см. рис. 2).</span><span class="sxs-lookup"><span data-stu-id="72793-p132">On the **Project** ribbon, choose the **Office Add-ins** drop-down menu, and then choose **See All**. In the **Insert Add-in** dialog box, choose **SHARED FOLDER** (see Figure 2).</span></span>

    <span data-ttu-id="72793-204">*Рис. 2. Запуск надстройки, расположенной в общей папке*</span><span class="sxs-lookup"><span data-stu-id="72793-204">*Figure 2. Starting an add-in that is on a file share*</span></span>

    ![Запуск приложения Office, расположенного в общей папке](../images/pj15-agave-overview-start-agave-apps.png)

8. <span data-ttu-id="72793-206">Выберите надстройку поиска Bing и нажмите кнопку **Вставить**.</span><span class="sxs-lookup"><span data-stu-id="72793-206">Select the Bing Search add-in, and then choose **Insert**.</span></span>

    <span data-ttu-id="72793-p133">Надстройка поиска Bing отображается в области задач, как показано на рисунке 3. Вы можете вручную изменить размер области задач и использовать надстройку поиска Bing.</span><span class="sxs-lookup"><span data-stu-id="72793-p133">The Bing Search add-in shows in a task pane, as in Figure 3. You can manually resize the task pane, and use the Bing Search add-in.</span></span>

    <span data-ttu-id="72793-209">*Рис. 3. Использование приложения поиска Bing*</span><span class="sxs-lookup"><span data-stu-id="72793-209">*Figure 3. Using the Bing Search add-in*</span></span>

    ![Использование приложения поиска Bing](../images/pj15-agave-overview-bing-search.png)

## <a name="distributing-project-add-ins"></a><span data-ttu-id="72793-211">Распространение надстроек Project</span><span class="sxs-lookup"><span data-stu-id="72793-211">Distributing Project add-ins</span></span>

<span data-ttu-id="72793-212">Вы можете распространять надстройки через общую папку, каталог приложений в библиотеке SharePoint или AppSource.</span><span class="sxs-lookup"><span data-stu-id="72793-212">You can distribute add-ins through a file share, an app catalog in a SharePoint library, or AppSource.</span></span> <span data-ttu-id="72793-213">Дополнительные сведения см. в статье [Публикация надстройки Office](../publish/publish.md).</span><span class="sxs-lookup"><span data-stu-id="72793-213">For more information, see [Publish your Office Add-in](../publish/publish.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="72793-214">См. также</span><span class="sxs-lookup"><span data-stu-id="72793-214">See also</span></span>

- [<span data-ttu-id="72793-215">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="72793-215">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
- [<span data-ttu-id="72793-216">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="72793-216">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
- [<span data-ttu-id="72793-217">Создание первой надстройки области задач для Project 2013 с помощью текстового редактора</span><span class="sxs-lookup"><span data-stu-id="72793-217">Create your first task pane add-in for Project 2013 by using a text editor</span></span>](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- [<span data-ttu-id="72793-218">Создание надстройки Project, использующей REST с локальной службой OData Project Server</span><span class="sxs-lookup"><span data-stu-id="72793-218">Create a Project add-in that uses REST with an on-premises Project Server OData service</span></span>](create-a-project-add-in-that-uses-rest-with-an-on-premises-odata-service.md)
- [<span data-ttu-id="72793-219">Загрузка пакета SDK для Project 2013</span><span class="sxs-lookup"><span data-stu-id="72793-219">Project 2013 SDK download</span></span>](https://www.microsoft.com/download/details.aspx?id=30435%20)
