---
title: Создание надстройки Project, использующей REST с локальной службой OData Project Server
description: ''
ms.date: 01/23/2018
ms.openlocfilehash: ce481438086f7e55dd27acb61010e61dff7153dc
ms.sourcegitcommit: c72c35e8389c47a795afbac1b2bcf98c8e216d82
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/23/2018
ms.locfileid: "19439210"
---
# <a name="create-a-project-add-in-that-uses-rest-with-an-on-premises-project-server-odata-service"></a><span data-ttu-id="bafa3-102">Создание надстройки Project, использующей REST с локальной службой OData Project Server</span><span class="sxs-lookup"><span data-stu-id="bafa3-102">Create a Project add-in that uses REST with an on-premises Project Server OData service</span></span>

<span data-ttu-id="bafa3-p101">В этой статье описывается создание надстройки области задач для Project профессиональный 2013, которая сравнивает данные по материальным и трудовым затратам в активном проекте со средними значениями из всех проектов в текущем экземпляре Project Web App. Надстройка использует REST с библиотекой jQuery для получения доступа к службе отчетов OData **ProjectData** в Project Server 2013.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p101">This article describes how to build a task pane add-in for Project Professional 2013 that compares cost and work data in the active project with the averages for all projects in the current Project Web App instance. The add-in uses REST with the jQuery library to access the  **ProjectData** OData reporting service in Project Server 2013.</span></span>


<span data-ttu-id="bafa3-105">Код в данной статье основан на примере, разработанном Саурабхом Сангхви (Saurabh Sanghvi) и Эрвиндом Лаиром (Arvind Iyer), сотрудниками корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="bafa3-105">The code in this article is based on a sample developed by Saurabh Sanghvi and Arvind Iyer, Microsoft Corporation.</span></span>

## <a name="prerequisites-for-creating-a-task-pane-add-in-that-reads-project-server-reporting-data"></a><span data-ttu-id="bafa3-106">Необходимые условия для создания надстроек области задач, читающей данные отчетов Project Server</span><span class="sxs-lookup"><span data-stu-id="bafa3-106">Prerequisites for creating a task pane add-in that reads Project Server reporting data</span></span>


<span data-ttu-id="bafa3-107">Далее приводятся необходимые условия для создания надстройки области задач Project, считывающей данные из службы **ProjectData** в экземпляре Project Web App локальной установки Project Server 2013:</span><span class="sxs-lookup"><span data-stu-id="bafa3-107">The following are the prerequisites for creating a Project task pane add-in that reads the  **ProjectData** service of a Project Web App instance in an on-premises installation of Project Server 2013:</span></span>


- <span data-ttu-id="bafa3-p102">Проверьте, что на локальном компьютере разработчика установлены самые последние пакеты обновления и обновления Windows. Операционной системой может быть Windows 7, Windows 8, Windows Server 2008 или Windows Server 2012.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p102">Ensure that you have installed the most recent service packs and Windows updates on your local development computer. The operating system can be Windows 7, Windows 8, Windows Server 2008, or Windows Server 2012.</span></span>
    
- <span data-ttu-id="bafa3-p103">Project профессиональный 2013 требуется для подключения к Project Web App. На компьютере разработчика должен быть установлен Project профессиональный 2013, чтобы включить отладку по клавише **F5** с помощью Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p103">Project Professional 2013 is required to connect with Project Web App. The development computer must have Project Professional 2013 installed to enable  **F5** debugging with Visual Studio.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="bafa3-112">С помощью Project стандартный 2013 можно размещать надстройки области задач, но невозможно войти в Project Web App.</span><span class="sxs-lookup"><span data-stu-id="bafa3-112">Project Standard 2013 can also host task pane add-ins, but cannot log on to Project Web App.</span></span>

- <span data-ttu-id="bafa3-113">Visual Studio 2015 с Инструменты разработчика Office для Visual Studio содержит шаблоны, позволяющие создавать Надстройки Office и SharePoint. Убедитесь, что у вас установлена самая последняя версия Office Developer Tools. См. раздел _Средства_ статьи [Надстройки Office и скачиваемые файлы для SharePoint](http://msdn.microsoft.com/en-us/office/apps/fp123627.aspx).</span><span class="sxs-lookup"><span data-stu-id="bafa3-113">Visual Studio 2015 with Office Developer Tools for Visual Studio includes templates for creating Office and SharePoint Add-ins. Ensure that you have installed the most recent version of Office Developer Tools; see the  _Tools_ section of the [Office Add-ins and SharePoint downloads](http://msdn.microsoft.com/en-us/office/apps/fp123627.aspx).</span></span>
    
- <span data-ttu-id="bafa3-p104">Процедуры и примеры кода, приведенные в этой статье, получают доступ к службе **ProjectData**, предоставляемой Project Server 2013 в локальном домене. Методы jQuery в этой статье не работают с Project Online.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p104">The procedures and code examples in this article access the  **ProjectData** service of Project Server 2013 in a local domain. The jQuery methods in this article do not work with Project Online.</span></span>
    
    <span data-ttu-id="bafa3-116">Убедитесь, что служба **ProjectData** доступна на компьютере разработчика.</span><span class="sxs-lookup"><span data-stu-id="bafa3-116">Verify that the  **ProjectData** service is accessible from your development computer.</span></span>
    

### <a name="procedure-1-to-verify-that-the-projectdata-service-is-accessible"></a><span data-ttu-id="bafa3-p105">Процедура 1. Проверка доступности службы ProjectData</span><span class="sxs-lookup"><span data-stu-id="bafa3-p105">Procedure 1. To verify that the ProjectData service is accessible</span></span>


1. <span data-ttu-id="bafa3-p106">Чтобы разрешить браузеру напрямую отображать XML-данные из запроса REST, отключите вид чтения канала. Дополнительные сведения о том, как это сделать в Internet Explorer, см. в процедуру 1, шаг 4 в статье [Создание запросов веб-каналов OData для данных отчетов Project](http://msdn.microsoft.com/library/3eafda3b-f006-48be-baa6-961b2ed9fe01%28Office.15%29.aspx).</span><span class="sxs-lookup"><span data-stu-id="bafa3-p106">To enable your browser to directly show the XML data from a REST query, turn off the feed reading view. For information about how to do this in Internet Explorer, see Procedure 1, step 4 in [Querying OData feeds for Project reporting data](http://msdn.microsoft.com/library/3eafda3b-f006-48be-baa6-961b2ed9fe01%28Office.15%29.aspx).</span></span>
    
2. <span data-ttu-id="bafa3-121">Отправьте запрос службе  **ProjectData** с помощью веб-обозревателя, используя следующий URL-адрес: **http://ServerName /ProjectServerName /_api/ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-121">Query the  **ProjectData** service by using your browser with the following URL:</span></span> <span data-ttu-id="bafa3-122">Например, если `http://MyServer/pwa` — это экземпляр Project Web App, то в браузере будут показаны следующие результаты:</span><span class="sxs-lookup"><span data-stu-id="bafa3-122">Query the  ProjectData service by using your browser with the following URL: http://ServerName /ProjectServerName /_api/ProjectData. For example, if the Project Web App instance is  `http://MyServer/pwa`, the browser shows the following results:</span></span>
    
    ```xml
    <?xml version="1.0" encoding="utf-8"?>
        <service xml:base="http://myserver/pwa/_api/ProjectData/" 
        xmlns="http://www.w3.org/2007/app" 
        xmlns:atom="http://www.w3.org/2005/Atom">
        <workspace>
            <atom:title>Default</atom:title>
            <collection href="Projects">
                <atom:title>Projects</atom:title>
            </collection>
            <collection href="ProjectBaselines">
                <atom:title>ProjectBaselines</atom:title>
            </collection>
            <!-- ... and 33 more collection elements -->
        </workspace>
        </service>
    ```

3. <span data-ttu-id="bafa3-p108">Вам может потребоваться предоставить свои сетевые учетные данные, чтобы увидеть результаты. Если браузер показывает сообщение "Ошибка 403, доступ запрещен", то либо у вас либо нет разрешений на вход для заданного экземпляра Project Web App, либо имеется проблема сети, требующая помощи администратора.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p108">You may have to provide your network credentials to see the results. If the browser shows "Error 403, Access Denied," either you do not have logon permission for that Project Web App instance, or there is a network problem that requires administrative help.</span></span>
    

## <a name="using-visual-studio-to-create-a-task-pane-add-in-for-project"></a><span data-ttu-id="bafa3-125">Создание надстройки области задач для Project с помощью Visual Studio</span><span class="sxs-lookup"><span data-stu-id="bafa3-125">Using Visual Studio to create a task pane add-in for Project</span></span>

<span data-ttu-id="bafa3-p109">Инструменты разработчика Office для Visual Studio включает шаблон надстроек области задач для Project 2013. Если вы создаете решение с именем **HelloProjectOData**, оно содержит следующие два проекта Visual Studio:</span><span class="sxs-lookup"><span data-stu-id="bafa3-p109">Office Developer Tools for Visual Studio includes a template for task pane add-ins for Project 2013. If you create a solution named  **HelloProjectOData**, the solution contains the following two Visual Studio projects:</span></span>


- <span data-ttu-id="bafa3-p110">Проект надстройки получает имя решения. Оно включает в себя XML-файл манифеста для приложения и настраивается на целевую платформу .NET Framework 4.5. В процедуре 3 показаны шаги по изменению манифеста надстройки **HelloProjectOData**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p110">The add-in project takes the name of the solution. It includes the XML manifest file for the add-in and targets the .NET Framework 4.5. Procedure 3 shows the steps to modify the manifest for the  **HelloProjectOData** add-in.</span></span>
    
- <span data-ttu-id="bafa3-p111">Веб-проект получает имя **HelloProjectODataWeb**. Оно содержит файлы JavaScript веб-страниц, файлы CSS, рисунки, ссылки и файлы конфигурации для веб-контента в области задач. Веб-проект настраивается на конечную платформу .NET Framework 4. В процедуре 4 и процедуре 5 показано, как изменить эти файлы в веб-проекте, чтобы создать функциональность надстройки **HelloProjectOData**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p111">The web project is named  **HelloProjectODataWeb**. It includes the webpages, JavaScript files, CSS files, images, references, and configuration files for the web content in the task pane. The web project targets the .NET Framework 4. Procedure 4 and Procedure 5 show how to modify the files in the web project to create the functionality of the  **HelloProjectOData** add-in.</span></span>
    

### <a name="procedure-2-to-create-the-helloprojectodata-add-in-for-project"></a><span data-ttu-id="bafa3-p112">Процедура 2. Создание надстройки HelloProjectOData для Project</span><span class="sxs-lookup"><span data-stu-id="bafa3-p112">Procedure 2. To create the HelloProjectOData add-in for Project</span></span>


1. <span data-ttu-id="bafa3-137">Запустите Visual Studio 2015 от имени администратора и выберите команду **Создать проект** на начальной странице.</span><span class="sxs-lookup"><span data-stu-id="bafa3-137">Run Visual Studio 2015 as an administrator, and then select  **New Project** on the Start page.</span></span>
    
2. <span data-ttu-id="bafa3-p113">В диалоговом окне **Новый проект** разверните узлы **Шаблоны** > **Visual C#** > **Office/SharePoint** и выберите **Надстройки Office**. Выберите **.NET Framework 4.5.2** в раскрывающемся списке в верхней части центральной панели, а затем выберите **Надстройка Office** (см. следующий снимок экрана).</span><span class="sxs-lookup"><span data-stu-id="bafa3-p113">In the  **New Project** dialog box, expand the **Templates**,  **Visual C#**, and  **Office/SharePoint** nodes, and then select ** Office Add-ins**. Select  **.NET Framework 4.5.2** in the target framework drop-down list at the top of the center pane, and then select **Office Add-in** (see the next screenshot).</span></span>
    
3. <span data-ttu-id="bafa3-140">Чтобы разместить оба проекта Visual Studio в одной папке, выберите **Создать каталог для решения** и найдите требуемое расположение.</span><span class="sxs-lookup"><span data-stu-id="bafa3-140">To place both of the Visual Studio projects in the same directory, select  **Create directory for solution**, and then browse to the location you want.</span></span>
    
4. <span data-ttu-id="bafa3-141">В поле **Имя** введите HelloProjectOData и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-141">In the  **Name** field, typeHelloProjectOData, and then choose  **OK**.</span></span>
    
    <span data-ttu-id="bafa3-142">*Рис. 1. Создание надстройки Office*</span><span class="sxs-lookup"><span data-stu-id="bafa3-142">*Figure 1. Creating an Office Add-in*</span></span>

    ![Создание надстройки Office](../images/pj15-hello-project-o-data-creating-app.png)

5. <span data-ttu-id="bafa3-144">В диалоговом окне **Выбор типа надстройки** выберите пункт **Надстройка области задач** и нажмите кнопку **Далее** (см. следующий снимок экрана).</span><span class="sxs-lookup"><span data-stu-id="bafa3-144">In the  **Choose the add-in type** dialog box, select **Task pane** and choose **Next** (see the next screenshot).</span></span>
    
    <span data-ttu-id="bafa3-145">*Рис. 2. Выбор типа создаваемой надстройки*</span><span class="sxs-lookup"><span data-stu-id="bafa3-145">*Figure 2. Choosing the type of add-in to create*</span></span>

    ![Выбор типа создаваемой надстройки](../images/pj15-hello-project-o-data-choose-project.png)

6. <span data-ttu-id="bafa3-147">В диалоговом окне **Выбор ведущих приложений** снимите все флажки, кроме флажка **Project** (см. следующий снимок экрана), а затем нажмите кнопку **Готово**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-147">In the  **Choose the host applications** dialog box, clear all check boxes except the **Project** check box (see the next screenshot) and choose **Finish**.</span></span>
    
    <span data-ttu-id="bafa3-148">*Рис. 3. Выбор ведущего приложения*</span><span class="sxs-lookup"><span data-stu-id="bafa3-148">*Figure 3. Choosing the host application*</span></span>

    ![Выбор Project в качестве единственного ведущего приложения](../images/create-office-add-in.png)
    
    <span data-ttu-id="bafa3-150">С помощью Visual Studio можно создавать проекты **HelloProjectOdata** и **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-150">Visual Studio creates the  **HelloProjectOdata** project and the **HelloProjectODataWeb** project.</span></span>
    
<span data-ttu-id="bafa3-p114">В папке **AddIn** (см. следующий снимок экрана) находится файл App.css для пользовательских стилей CSS. В дочерней папке **Home** находится файл Home.html, который содержит ссылки на CSS-файлы и файлы JavaScript, которые использует надстройка, и код HTML5 для надстройки. Кроме того, файл Home.js предназначен для пользовательского кода JavaScript. В папке **Scripts** находятся файлы библиотек jQuery. В дочерней папке **Office** находятся библиотеки JavaScript, например office.js и project-15.js, а также языковые библиотеки для стандартных строк в надстройках Office. В папке **Content** находится файл Office.css, который содержит стили по умолчанию для всех надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p114">The  **AddIn** folder (see the next screenshot) contains the App.css file for custom CSS styles. In the **Home** subfolder , the Home.html file contains references to the CSS files and the JavaScript files that the add-in uses, and the HTML5 content for the add-in. Also, the Home.js file is for your custom JavaScript code. The **Scripts** folder includes the jQuery library files. The **Office** subfolder includes the JavaScript libraries such as office.js and project-15.js, plus the language libraries for standard strings in the Office add-ins. In the **Content** folder, the Office.css file contains the default styles for all of the Office add-ins.</span></span>

<span data-ttu-id="bafa3-156">*Рис. 4. Просмотр файлов веб-проекта по умолчанию в обозревателе решений*</span><span class="sxs-lookup"><span data-stu-id="bafa3-156">*Figure 4. Viewing the default web project files in Solution Explorer*</span></span>

![Просмотр файлов веб-проекта в обозревателе решений](../images/pj15-hello-project-o-data-initial-solution-explorer.png)

<span data-ttu-id="bafa3-p115">Манифест проекта **HelloProjectOData** — это файл HelloProjectOData.xml. Его можно изменить при необходимости, чтобы добавить описание надстройки, ссылку на значок, сведения о дополнительных языках и другие параметры. В процедуре 3 изменяется только отображаемое имя надстройки и описание и добавляется значок.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p115">The manifest for the  **HelloProjectOData** project is the HelloProjectOData.xml file. You can optionally modify the manifest to add a description of the add-in, a reference to an icon, information for additional languages, and other settings. Procedure 3 simply modifies the add-in display name and description, and adds an icon.</span></span>

<span data-ttu-id="bafa3-161">Дополнительные сведения о манифесте см. в статьях [XML-манифест надстроек для Office](../develop/add-in-manifests.md) и [Справка по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md#see-also).</span><span class="sxs-lookup"><span data-stu-id="bafa3-161">For more information about the manifest, see [Office Add-ins XML manifest](../develop/add-in-manifests.md) and [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md#see-also).</span></span>

### <a name="procedure-3-to-modify-the-add-in-manifest"></a><span data-ttu-id="bafa3-p116">Процедура 3. Изменение манифеста надстройки</span><span class="sxs-lookup"><span data-stu-id="bafa3-p116">Procedure 3. To modify the add-in manifest</span></span>


1. <span data-ttu-id="bafa3-164">Откройте файл HelloProjectOData.xml в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="bafa3-164">In Visual Studio, open the HelloProjectOData.xml file.</span></span>
    
2. <span data-ttu-id="bafa3-p117">Отображаемое имя по умолчанию — это имя проекта Visual Studio ("HelloProjectOData"). Например, измените значение по умолчанию элемента **DisplayName** на значение"Hello ProjectData".</span><span class="sxs-lookup"><span data-stu-id="bafa3-p117">The default display name is the name of the Visual Studio project ("HelloProjectOData"). For example, change the default value of the  **DisplayName** element to"Hello ProjectData".</span></span>
    
3. <span data-ttu-id="bafa3-p118">Описание по умолчанию — "HelloProjectOData". Например, измените значение по умолчанию элемента Description на "Test REST queries of the ProjectData service" (тестирование запросов REST службы ProjectData).</span><span class="sxs-lookup"><span data-stu-id="bafa3-p118">The default description is also "HelloProjectOData". For example, change the default value of the Description element to "Test REST queries of the ProjectData service".</span></span>
    
4. <span data-ttu-id="bafa3-p119">Добавьте значок для отображения в раскрывающемся списке **Надстройки Office** на вкладке **Проект** ленты. Вы можете добавить файл значка в решении Visual Studio или использовать URL-адрес значка.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p119">Add an icon to show in the  **Office Add-ins** drop-down list on the **PROJECT** tab of the ribbon. You can add an icon file in the Visual Studio solution or use a URL for an icon.</span></span> 

<span data-ttu-id="bafa3-171">Ниже описано, как добавить файл значка в решение Visual Studio:</span><span class="sxs-lookup"><span data-stu-id="bafa3-171">The following steps show how to add an icon file to the Visual Studio solution:</span></span>
    
1. <span data-ttu-id="bafa3-172">В **обозревателе решений** откройте папку Images.</span><span class="sxs-lookup"><span data-stu-id="bafa3-172">In  **Solution Explorer**, go to the folder named Images.</span></span>
    
2. <span data-ttu-id="bafa3-p120">Чтобы отображаться в раскрывающемся списке **Надстройки Office**, значок должен иметь размер 32 x 32 пикселя. Например, установите пакет SDK Project 2013, затем выберите папку **Images** и добавьте следующий файл из пакета SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`</span><span class="sxs-lookup"><span data-stu-id="bafa3-p120">To be displayed in the  **Office Add-ins** drop-down list, the icon must be 32 x 32 pixels. For example, install the Project 2013 SDK, and then choose the **Images** folder and add the following file from the SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`</span></span>
    
    <span data-ttu-id="bafa3-175">Вы можете использовать собственный значок размером 32 x 32 пикселя или скопировать следующее изображение в файл с именем NewIcon.png, а затем добавить этот файл в папку `HelloProjectODataWeb\Images`:</span><span class="sxs-lookup"><span data-stu-id="bafa3-175">Alternately, use your own 32 x 32 icon; or, copy the following image to a file named NewIcon.png, and then add that file to the  `HelloProjectODataWeb\Images` folder:</span></span>
    
    ![Значок для приложения HelloProjectOData](../images/pj15-hello-project-data-new-icon.jpg)

3. <span data-ttu-id="bafa3-p121">В манифесте HelloProjectOData.xml добавьте элемент **IconUrl** под элементом **Description**. Значением URL-адреса значка является относительный путь на файл значка размером 32 x 32. Например, добавьте следующую строку: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**. Теперь файл манифеста HelloProjectOData.xml содержит следующий текст (ваше значение **Id** будет другим):</span><span class="sxs-lookup"><span data-stu-id="bafa3-p121">In the HelloProjectOData.xml manifest, add an  **IconUrl** element below the **Description** element, where the value of the icon URL is the relative path to the 32x32 icon file. For example, add the following line: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**. The HelloProjectOData.xml manifest file now contains the following (your  **Id** value will be different):</span></span>

    ```XML
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1" 
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <Id>c512df8d-a1c5-4d74-8a34-d30f6bbcbd82 </Id>
        <Version>1.0</Version>
        <ProviderName> [Provider name]</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Hello ProjectData" />
        <Description DefaultValue="Test REST queries of the ProjectData service"/>
        <IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />

        <Hosts>
            <Host Name="Project" />
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="~remoteAppUrl/AddIn/Home/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

## <a name="creating-the-html-content-for-the-helloprojectodata-add-in"></a><span data-ttu-id="bafa3-180">Создание HTML-контента для надстройки HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="bafa3-180">Creating the HTML content for the HelloProjectOData add-in</span></span>

<span data-ttu-id="bafa3-p122">Надстройка **HelloProjectOData** — это пример, который содержит сообщения отладки и сообщения об ошибках. Она не предназначена для использования в рабочей среде. Перед началом написания кода HTML-контента разработайте пользовательский интерфейс и алгоритм работы пользователя с надстройкой, а также выделите функции JavaScript, взаимодействующие с HTML-кодом. Дополнительные сведения см. в статье[Рекомендации по проектированию надстроек Office](../design/add-in-design.md).</span><span class="sxs-lookup"><span data-stu-id="bafa3-p122">The  **HelloProjectOData** add-in is a sample that includes debugging and error output; it is not intended for production use. Before you start coding the HTML content, design the UI and user experience for the add-in, and outline the JavaScript functions that interact with the HTML code. For more information, see[Design guidelines for Office Add-ins](../design/add-in-design.md).</span></span> 

<span data-ttu-id="bafa3-p123">В верхней части области задач размещается отображаемое имя надстройки, соответствующее значению элемента **DisplayName** в манифесте. Элемент **body** в файле HelloProjectOData.html содержит другие элементы пользовательского интерфейса:</span><span class="sxs-lookup"><span data-stu-id="bafa3-p123">The task pane shows the add-in display name at the top, which is the value of the  **DisplayName** element in the manifest. The **body** element in the HelloProjectOData.html file contains the other UI elements, as follows:</span></span>

- <span data-ttu-id="bafa3-186">Подзаголовок, указывающий на общую функциональность или тип работы, например: **ODATA REST QUERY**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-186">A subtitle indicates the general functionality or type of operation, for example,  **ODATA REST QUERY**.</span></span>
    
- <span data-ttu-id="bafa3-p124">Кнопка **Get ProjectData Endpoint** вызывает функцию **setOdataUrl** для получения конечной точки службы **ProjectData** и отображения ее в текстовом поле. Если Project не подключен к Project Web App, надстройка вызовет обработчик ошибок для отображения всплывающего сообщения об ошибке.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p124">The  **Get ProjectData Endpoint** button calls the **setOdataUrl** function to get the endpoint of the **ProjectData** service, and display it in a text box. If Project is not connected with Project Web App, the add-in calls an error handler to display a pop-up error message.</span></span>
    
- <span data-ttu-id="bafa3-p125">Кнопка **Compare All Projects** отключена до тех пор, пока надстройка не получит действительную конечную точку OData. Когда пользователь нажимает эту кнопку, она вызывает функцию **retrieveOData**, которая использует запрос REST для получения сведений о материальных и трудовых затратах проекта из службы **ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p125">The  **Compare All Projects** button is disabled until the add-in gets a valid OData endpoint. When you select the button, it calls the **retrieveOData** function, which uses a REST query to get project cost and work data from the **ProjectData** service.</span></span>
    
- <span data-ttu-id="bafa3-p126">Таблица отображает средние значения затрат проекта, фактических затрат, трудозатрат и процент выполнения. В таблице также сравниваются значения текущего активного проекта со средними. Если текущее значение больше среднего по всем проектам, значение отображается красным цветом. Если текущее значение меньше среднего, оно отображается зеленым цветом. Если текущее значение недоступно, в таблице отображается значение **NA** синим цветом.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p126">A table displays the average values for project cost, actual cost, work, and percent complete. The table also compares the current active project values with the average. If the current value is greater than the average for all projects, the value is displayed as red. If the current value is less than the average, the value is displayed as green. If the current value is not available, the table displays a blue  **NA**.</span></span>
    
    <span data-ttu-id="bafa3-196">Функция **retrieveOData** вызывает функцию **parseODataResult**, которая вычисляет и отображает значения таблицы.</span><span class="sxs-lookup"><span data-stu-id="bafa3-196">The  **retrieveOData** function calls the **parseODataResult** function, which calculates and displays values for the table.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="bafa3-p127">В этом примере данные о материальных и трудовых затратах по активному проекту извлекаются из опубликованных значений. Если изменить значения в Project, служба **ProjectData** не будет знать об изменениях до тех пор, пока проект не будет опубликован.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p127">In this example, cost and work data for the active project are derived from the published values. If you change values in Project, the  **ProjectData** service does not have the changes until the project is published.</span></span>


### <a name="procedure-4-to-create-the-html-content"></a><span data-ttu-id="bafa3-p128">Процедура 4. Создание HTML-контента</span><span class="sxs-lookup"><span data-stu-id="bafa3-p128">Procedure 4. To create the HTML content</span></span>

1. <span data-ttu-id="bafa3-p129">В элементе **head** файла Home.html добавьте любые дополнительные элементы **link** для CSS-файлов, используемых в надстройке. Шаблон проекта Visual Studio содержит ссылку на файл App.css, который можно использовать для настраиваемых стилей CSS.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p129">In the  **head** element of the Home.html file, add any additional **link** elements for CSS files that your add-in uses. The Visual Studio project template includes a link for the App.css file that you can use for custom CSS styles.</span></span>
    
2. <span data-ttu-id="bafa3-p130">Добавьте любые дополнительные элементы **script** для библиотек JavaScript, используемых в надстройке. Шаблон проекта содержит ссылки на файлы jQuery- _[версия]_.js, office.js и MicrosoftAjax.js из папки **Scripts**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p130">Add any additional  **script** elements for JavaScript libraries that your add-in uses. The project template includes links for the jQuery- _[version]_.js, office.js, and MicrosoftAjax.js files in the  **Scripts** folder.</span></span>
    
    > [!NOTE]
    > <span data-ttu-id="bafa3-p131">Перед развертыванием надстройки измените ссылку office.js и ссылку jQuery на ссылку сети доставки содержимого (CDN). Ссылка CDN предоставляет самую последнюю версию и обеспечивает оптимальную производительность.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p131">Before you deploy the add-in, change the office.js reference and the jQuery reference to the content delivery network (CDN) reference. The CDN reference provides the most recent version and better performance.</span></span>

    <span data-ttu-id="bafa3-p132">Надстройка **HelloProjectOData** также использует файл SurfaceErrors.js, с помощью которого во всплывающих сообщениях отображаются ошибки. Можно скопировать код из раздела _Надежное программирование_ статьи [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md), а затем добавить файл SurfaceErrors.js в папку **Scripts\Office** проекта **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p132">The  **HelloProjectOData** add-in also uses the SurfaceErrors.js file, which displays errors in a pop-up message. You can copy the code from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md), and then add a SurfaceErrors.js file in the  **Scripts\Office** folder of the **HelloProjectODataWeb** project.</span></span>
    
    <span data-ttu-id="bafa3-209">Ниже приведен обновленный HTML-код элемента **head** с дополнительной строкой для файла SurfaceErrors.js.</span><span class="sxs-lookup"><span data-stu-id="bafa3-209">Following is the updated HTML code for the  **head** element, with the additional line for the SurfaceErrors.js file:</span></span>
    
    ```HTML
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Test ProjectData Service</title>
    
    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />
    
    <!-- Add your CSS styles to the following file -->
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />
    
    <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
    <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
    <script src="../Scripts/jquery-1.7.1.js"></script>
    
    <!-- Use the CDN reference to office.js when deploying your add-in. -->
    <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->
    
    <!-- Use the local script references for Office.js to enable offline debugging -->
    <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/1.0/Office.js"></script>
    
    <!-- Add your JavaScript to the following files -->
    <script src="../Scripts/HelloProjectOData.js"></script>
    <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
    <!-- See the code in Step 3. -->
    </body>
    </html>
    ```

3. <span data-ttu-id="bafa3-p133">В элементе **body** удалите имеющийся код из шаблона, а затем добавьте код для пользовательского интерфейса. Если элемент требуется заполнить данными или изменить с помощью оператора jQuery, то он должен содержать уникальный атрибут **id**. В приведенном ниже коде атрибуты **id** элементов **button**, **span** и **td** (определение ячейки таблицы), используемых функциями jQuery, выделены полужирным шрифтом.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p133">In the **body** element, delete the existing code from the template, and then add the code for the user interface. If an element is to be filled with data or manipulated by a jQuery statement, the element must include a unique **id** attribute. In the following code, the **id** attributes for the **button**,  **span**, and  **td** (table cell definition) elements that jQuery functions use are shown in bold font.</span></span>
    
   <span data-ttu-id="bafa3-p134">С помощью приведенного ниже HTML-кода можно добавить графическое изображение (например, логотип компании). Можно использовать логотип на свой выбор или же скопировать файл NewLogo.png из скачанного пакета SDK для Project 2013, а затем с помощью **обозревателя решений** добавить файл в папку `HelloProjectODataWeb\Images`.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p134">The following HTML adds a graphic image, which could be a company logo. You can use a logo of your choice, or copy the NewLogo.png file from the Project 2013 SDK download, and then use  **Solution Explorer** to add the file to the `HelloProjectODataWeb\Images` folder.</span></span>
    
    ```HTML
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br /><br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
            <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
            <table class="infoTable" aria-readonly="True" style="width: 100%;">
                <tr>
                    <td class="heading_leftCol"></td>
                    <td class="heading_midCol"><strong>Average</strong></td>
                    <td class="heading_rightCol"><strong>Current</strong></td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
                    <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project Work</strong></td>
                    <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
                </tr>
                <tr>
                    <td class="row_leftCol"><strong>Project % Complete</strong></td>
                    <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
                    <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
                </tr>
            </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
    ```


## <a name="creating-the-javascript-code-for-the-add-in"></a><span data-ttu-id="bafa3-215">Создание кода JavaScript для надстройки</span><span class="sxs-lookup"><span data-stu-id="bafa3-215">Creating the JavaScript code for the add-in</span></span>

<span data-ttu-id="bafa3-p135">Шаблон надстройки области задач для Project содержит код инициализации по умолчанию, который предназначен для демонстрации базовых действий получения и записи данных в документе для типичных приложений Office 2013. Так как Project 2013 не поддерживает действия записи в активный проект, а надстройка **HelloProjectOData** не использует метод **getSelectedDataAsync**, то можно удалить скрипт в функции **Office.initialize** и удалить функцию **setData** и функцию **getData** в файле HelloProjectOData.js по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p135">The template for a Project task pane add-in includes default initialization code that is designed to demonstrate basic get and set actions for data in a document for a typical Office 2013 add-in. Because Project 2013 does not support actions that write to the active project, and the  **HelloProjectOData** add-in does not use the **getSelectedDataAsync** method, you can delete the script within the **Office.initialize** function, and delete the **setData** function and **getData** function in the default HelloProjectOData.js file.</span></span>

<span data-ttu-id="bafa3-p136">В JavaScript содержатся глобальные константы для запроса REST и глобальные переменные, используемые в нескольких функциях. Кнопка **Get ProjectData Endpoint (получить конечную точку ProjectData)** вызывает функцию **setOdataUrl**, инициализирующую глобальные переменные и определяющую, подключен ли Project к Project Web App.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p136">The JavaScript includes global constants for the REST query and global variables that are used in several functions. The  **Get ProjectData Endpoint** button calls the **setOdataUrl** function, which initializes the global variables and determines whether Project is connected with Project Web App.</span></span>

<span data-ttu-id="bafa3-220">Оставшаяся часть файла HelloProjectOData.js содержит две функции: parseODataResult и retrieveOData. Функция **retrieveOData** вызывается когда пользователь выбирает команду **Compare All Projects (сравнить все проекты)**. Функция **parseODataResult** вычисляет средние значения, а затем заполняет таблицу сравнения значениями, отформатированными в соответствии с цветом и единицами измерения.</span><span class="sxs-lookup"><span data-stu-id="bafa3-220">The remainder of the HelloProjectOData.js file includes two functions: the  **retrieveOData** function is called when the user selects **Compare All Projects**; and the  **parseODataResult** function calculates averages and then populates the comparison table with values that are formatted for color and units.</span></span>

### <a name="procedure-5-to-create-the-javascript-code"></a><span data-ttu-id="bafa3-p137">Процедура 5. Создание кода JavaScript</span><span class="sxs-lookup"><span data-stu-id="bafa3-p137">Procedure 5. To create the JavaScript code</span></span>

1. <span data-ttu-id="bafa3-p138">Удалите весь код в файле HelloProjectOData.js по умолчанию и затем добавьте глобальные переменные и функцию **Office.initialize**. Имена переменных, написанные полностью заглавными буквами подразумевают, что они являются константами; они позже будут использоваться с переменной **_pwa** для создания запроса REST в этом примере.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p138">Delete all code in the default HelloProjectOData.js file, and then add the global variables and  **Office.initialize** function. Variable names that are all capitals imply that they are constants; they are later used with the **_pwa** variable to create the REST query in this example.</span></span>
    
    ```js
    var PROJDATA = "/_api/ProjectData";
    var PROJQUERY = "/Projects?";
    var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
    var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
    var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
    var _pwa;           // URL of Project Web App.
    var _projectUid;    // GUID of the active project.
    var _docUrl;        // Path of the project document.
    var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData
    
    // The initialize function is required for all add-ins.
    Office.initialize = function (reason) {
        // Checks for the DOM to load using the jQuery ready function.
        $(document).ready(function () {
            // After the DOM is loaded, app-specific code can run.
        });
    }
    ```

2. <span data-ttu-id="bafa3-p139">Добавьте функцию **setOdataUrl** и связанные функции. Функция **setOdataUrl** вызывает **getProjectGuid** и **getDocumentUrl** для инициализации глобальных переменных. В [методе getProjectFieldAsync](https://dev.office.com/reference/add-ins/shared/projectdocument.getprojectfieldasync) анонимная функция для параметра _callback_ включает кнопку **Compare All Projects** (сравнить все проекты) с помощью метода **removeAttr** из библиотеки jQuery, а затем отображает URL-адрес службы **ProjectData**. Если Project не подключен к Project Web App, функция вызывает ошибку, которая отображает всплывающее сообщение об ошибке. Файл SurfaceErrors.js содержит метод **throwError**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p139">Add  **setOdataUrl** and related functions. The **setOdataUrl** function calls **getProjectGuid** and **getDocumentUrl** to initialize the global variables. In the [getProjectFieldAsync method](https://dev.office.com/reference/add-ins/shared/projectdocument.getprojectfieldasync), the anonymous function for the  _callback_ parameter enables the **Compare All Projects** button by using the **removeAttr** method in the jQuery library, and then displays the URL of the **ProjectData** service. If Project is not connected with Project Web App, the function throws an error, which displays a pop-up error message. The SurfaceErrors.js file includes the **throwError** method.</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="bafa3-p140">Если вы работаете в Visual Studio на компьютере с Project Server, раскомментируйте код после строки, отвечающей за инициализацию глобальной переменной **_pwa**, чтобы можно было выполнять его отладки с помощью клавиши **F5**. Чтобы использовать метод jQuery **ajax** во время отладки на компьютере с Project Server, следует задать значение **localhost** для URL-адреса PWA. При работе в Visual Studio на удаленном компьютере URL-адрес **localhost** не требуется. Перед развертыванием надстройки закомментируйте этот код.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p140">If you run Visual Studio on the Project Server computer, to use  **F5** debugging, uncomment the code after the line that initializes the **_pwa** global variable. To enable using the jQuery **ajax** method when debugging on the Project Server computer, you must set the **localhost** value for the PWA URL.If you run Visual Studio on a remote computer, the  **localhost** URL is not required. Before you deploy the add-in, comment out that code.</span></span>

    ```js
    function setOdataUrl() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.ProjectServerUrl,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _pwa = String(asyncResult.value.fieldValue);
    
                    // If you debug with Visual Studio on a local Project Server computer, 
                    // uncomment the following lines to use the localhost URL.
                    //var localhost = location.host.split(":", 1);
                    //var pwaStartPosition = _pwa.lastIndexOf("/");
                    //var pwaLength = _pwa.length - pwaStartPosition;
                    //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                    //_pwa = location.protocol + "//" + localhost + pwaName;
    
                    if (_pwa.substring(0, 4) == "http") {
                        _odataUrl = _pwa + PROJDATA;
                        $("#compareProjects").removeAttr("disabled");
                        getProjectGuid();
                    }
                    else {
                        _odataUrl = "No connection!";
                        throwError(_odataUrl, "You are not connected to Project Web App.");
                    }
                    getDocumentUrl();
                    $("#projectDataEndPoint").text(_odataUrl);
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }

    // Get the GUID of the active project.
    function getProjectGuid() {
        Office.context.document.getProjectFieldAsync(
            Office.ProjectProjectFields.GUID,
            function (asyncResult) {
                if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                    _projectUid = asyncResult.value.fieldValue;
                }
                else {
                    throwError(asyncResult.error.name, asyncResult.error.message);
                }
            }
        );
    }
    
    // Get the path of the project in Project web app, which is in the form <>\ProjectName .
    function getDocumentUrl() {
        _docUrl = "Document path:\r\n" + Office.context.document.url;
    }
    ```

3. <span data-ttu-id="bafa3-p141">Добавьте функцию **retrieveOData**, которая объединяет значения для запроса REST и затем вызывает функцию **ajax** в jQuery для получения запрошенных данных из службы **ProjectData**. Переменная **support.cors** позволяет производить межплатформенный обмен ресурсами (CORS) с функцией **ajax**. Если оператор **support.cors** пропущен или имеет значение **false**, функция **ajax** возвращает ошибку **No transport (нет передачи)**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p141">Add the  **retrieveOData** function, which concatenates values for the REST query and then calls the **ajax** function in jQuery to get the requested data from the **ProjectData** service. The **support.cors** variable enables cross-origin resource sharing (CORS) with the **ajax** function. If the **support.cors** statement is missing or is set to **false**, the  **ajax** function returns a **No transport** error.</span></span>
    
   > [!NOTE]
   > <span data-ttu-id="bafa3-p142">Приведенный ниже код подходит для локального сервера Project Server 2013. В Project Online можно использовать OAuth для проверки подлинности на основе токенов. Дополнительные сведения см. в статье [Обход ограничений, связанных с принципом одинакового источника, в надстройках Office](../develop/addressing-same-origin-policy-limitations.md).</span><span class="sxs-lookup"><span data-stu-id="bafa3-p142">The following code works with an on-premises installation of Project Server 2013. For Project Online, you can use OAuth for token-based authentication. For more information, see [Addressing same-origin policy limitations in Office Add-ins](../develop/addressing-same-origin-policy-limitations.md).</span></span>

   <span data-ttu-id="bafa3-p143">Для вызова **ajax** можно использовать параметр _headers_ или _beforeSend_. Параметр _complete_ — анонимная функция, поэтому находится в той же области, что и переменные в **retrieveOData**. Функция для параметра _complete_ выводит результаты в элементе управления **odataText**, а также вызывает метод **parseODataResult** для анализа и отображения отклика JSON. Параметр _error_ указывает именованную функцию **getProjectDataErrorHandler**, которая записывает сообщение об ошибке в элемент управления **odataText**, а также выводит всплывающее сообщение с помощью метода **throwError**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p143">In the **ajax** call, you can use either the _headers_ parameter or the _beforeSend_ parameter. The _complete_ parameter is an anonymous function so that it is in the same scope as the variables in **retrieveOData**. The function for the  _complete_ parameter displays results in the **odataText** control and also calls the **parseODataResult** method to parse and display the JSON response. The _error_ parameter specifies the named **getProjectDataErrorHandler** function, which writes an error message to the **odataText** control and also uses the **throwError** method to display a pop-up message.</span></span>

    ```js
    /****************************************************************
    * Functions to get and parse the Project Server reporting data.
    *****************************************************************/
    
    // Get data about all projects on Project Server, 
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();
    
        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project Online.
        $.support.cors = true;
    
        $.ajax({
            url: restUrl,
            type: "GET",
            contentType: "application/json",
            data: "",      // Empty string for the optional data.
            //headers: { "Accept": accept },
            beforeSend: function (xhr) {
                xhr.setRequestHeader("ACCEPT", accept);
            },
            complete: function (xhr, textStatus) {
                // Create a message to display in the text box.
                var message = "\r\ntextStatus: " + textStatus +
                    "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                    "\r\nStatus: " + xhr.status +
                    "\r\nResponseText:\r\n" + xhr.responseText;
    
                // xhr.responseText is the result from an XmlHttpRequest, which 
                // contains the JSON response from the OData service.
                parseODataResult(xhr.responseText, _projectUid);
    
                // Write the document name, response header, status, and JSON to the odataText control.
                $("#odataText").text(_docUrl);
                $("#odataText").append("\r\nREST query:\r\n" + restUrl);
                $("#odataText").append(message);
    
                if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                    $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
                }
            },
            error: getProjectDataErrorHandler
        });
    }
    
    function getProjectDataErrorHandler(data, errorCode, errorMessage) {
        $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
        throwError(errorCode, errorMessage);
    }
    ```

4. <span data-ttu-id="bafa3-p144">Добавьте метод **parseODataResult**, который десериализует и обрабатывает отклик JSON из службы OData. Метод **parseODataResult** вычисляет средние значения материальных и трудовых затрат с точностью до одного или двух десятичных знаков, форматирует значения необходимым цветом и добавляет единицу измерения (**$**, **hrs** или **%**), а затем выводит значения в заданных ячейках таблицы.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p144">Add the **parseODataResult** method, which deserializes and processes the JSON response from the OData service. The **parseODataResult** method calculates average values of the cost and work data to an accuracy of one or two decimal places, formats values with the correct color and adds a unit ( **$**,  **hrs**, or  **%**), and then displays the values in specified table cells.</span></span>
    
   <span data-ttu-id="bafa3-p145">Если GUID активного проекта соответствует значению **ProjectId**, переменной **myProjectIndex** присваивается индекс проекта. Если **myProjectIndex** указывает, что активный проект опубликован на сервере Project Server, метод **parseODataResult** форматирует и отображает данные о затратах и работе для этого проекта. Если активный проект не опубликован, значения для него отображаются как **НД** в синем цвете.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p145">If the GUID of the active project matches the  **ProjectId** value, the **myProjectIndex** variable is set to the project index. If **myProjectIndex** indicates the active project is published on Project Server, the **parseODataResult** method formats and displays cost and work data for that project. If the active project is not published, values for the active project are displayed as a blue **NA**.</span></span>

    ```js
    // Calculate the average values of actual cost, cost, work, and percent complete   
    // for all projects, and compare with the values for the current project.
    function parseODataResult(oDataResult, currentProjectGuid) {
        // Deserialize the JSON string into a JavaScript object.
        var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
        var len = res.d.results.length;
        var projActualCost = 0;
        var projCost = 0;
        var projWork = 0;
        var projPercentCompleted = 0;
        var myProjectIndex = -1;
        for (i = 0; i < len; i++) {
            // If the current project GUID matches the GUID from the OData query,  
            // store the project index.
            if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
                myProjectIndex = i;
            }
            projCost += Number(res.d.results[i].ProjectCost);
            projWork += Number(res.d.results[i].ProjectWork);
            projActualCost += Number(res.d.results[i].ProjectActualCost);
            projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);
        }
        var avgProjCost = projCost / len;
        var avgProjWork = projWork / len;
        var avgProjActualCost = projActualCost / len;
        var avgProjPercentCompleted = projPercentCompleted / len;
        
        // Round off cost to two decimal places, and round off other values to one decimal place.
        avgProjCost = avgProjCost.toFixed(2);
        avgProjWork = avgProjWork.toFixed(1);
        avgProjActualCost = avgProjActualCost.toFixed(2);
        avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);
        
        // Display averages in the table, with the correct units. 
        document.getElementById("AverageProjectCost").innerHTML = "$"
            + avgProjCost;
        document.getElementById("AverageProjectActualCost").innerHTML
            = "$" + avgProjActualCost;
        document.getElementById("AverageProjectWork").innerHTML
            = avgProjWork + " hrs";
        document.getElementById("AverageProjectPercentComplete").innerHTML
            = avgProjPercentCompleted + "%";
            
        // Calculate and display values for the current project.
        if (myProjectIndex != -1) {
            var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
            var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
            var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
            var myProjPercentCompleted =
            Number(res.d.results[myProjectIndex].ProjectPercentCompleted);
            
            myProjCost = myProjCost.toFixed(2);
            myProjWork = myProjWork.toFixed(1);
            myProjActualCost = myProjActualCost.toFixed(2);
            myProjPercentCompleted = myProjPercentCompleted.toFixed(1);
            
            document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;
            
            if (Number(myProjCost) <= Number(avgProjCost)) {
                document.getElementById("CurrentProjectCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectCost").style.color = "red"
            }
            
            document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;
            
            if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
                document.getElementById("CurrentProjectActualCost").style.color = "green"
            }
            else {
                document.getElementById("CurrentProjectActualCost").style.color = "red"
            }
            
            document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";
            
            if (Number(myProjWork) <= Number(avgProjWork)) {
                document.getElementById("CurrentProjectWork").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectWork").style.color = "green"
            }
            
            document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";
            
            if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
                document.getElementById("CurrentProjectPercentComplete").style.color = "red"
            }
            else {
                document.getElementById("CurrentProjectPercentComplete").style.color = "green"
            }
        }
        else {
            document.getElementById("CurrentProjectCost").innerHTML = "NA";
            document.getElementById("CurrentProjectCost").style.color = "blue"
            
            document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
            document.getElementById("CurrentProjectActualCost").style.color = "blue"
            
            document.getElementById("CurrentProjectWork").innerHTML = "NA";
            document.getElementById("CurrentProjectWork").style.color = "blue"
            
            document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
            document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
        }
    }
    ```


## <a name="testing-the-helloprojectodata-add-in"></a><span data-ttu-id="bafa3-248">Тестирование надстройки HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="bafa3-248">Testing the HelloProjectOData add-in</span></span>

<span data-ttu-id="bafa3-p146">Для тестирования и отладки надстройки **HelloProjectOData** с помощью Visual Studio 2015 на компьютере разработки должен быть установлен Project профессиональный 2013. Для работы с различными тестовыми сценариями убедитесь, что можно выбрать открытие файлов Project на локальном компьютере или подключение к Project Web App. Например, выполните следующие действия.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p146">To test and debug the  **HelloProjectOData** add-in with Visual Studio 2015, Project Professional 2013 must be installed on the development computer. To enable different test scenarios, ensure that you can choose whether Project opens for files on the local computer or connects with Project Web App. For example, do the following steps:</span></span>

1. <span data-ttu-id="bafa3-252">Во вкладке **ФАЙЛ** на ленте выберите вкладку **Сведения** в представлении Backstage, а затем выберите **Управление учетными записями**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-252">On the  **FILE** tab on the ribbon, choose the **Info** tab in the Backstage view, and then choose **Manage Accounts**.</span></span>
    
2. <span data-ttu-id="bafa3-p147">В диалоговом окне **Учетные записи Project Web App** список **Доступные учетные записи** может содержать несколько учетных записей Project Web App помимо локальной учетной записи **Компьютер**. В разделе **Во время запуска** выберите **Выбрать учетную запись**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p147">In the  **Project web app Accounts** dialog box, the **Available accounts** list can have multiple Project Web App accounts in addition to the local **Computer** account. In the **When starting** section, select **Choose an account**.</span></span>
    
3. <span data-ttu-id="bafa3-255">Закройте Project, чтобы среда Visual Studio могла запустить его для отладки надстройки.</span><span class="sxs-lookup"><span data-stu-id="bafa3-255">Close Project so that Visual Studio can start it for debugging the add-in.</span></span>
    
<span data-ttu-id="bafa3-256">Базовые тесты должны быть следующие:</span><span class="sxs-lookup"><span data-stu-id="bafa3-256">Basic tests should include the following:</span></span>

- <span data-ttu-id="bafa3-p148">Запустите приложение в Visual Studio и откройте опубликованный проект из Project Web App, содержащего данные о материальных и трудовых затратах. Убедитесь, что надстройка отображает конечную точку **ProjectData** и правильно отображает данные о материальных и трудовых затратах в таблице. Можно использовать выходные данные в элементе управления **odataText** для проверки запроса REST и других сведений.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p148">Run the add-in from Visual Studio, and then open a published project from Project Web App that contains cost and work data. Verify that the add-in displays the  **ProjectData** endpoint and correctly displays the cost and work data in the table. You can use the output in the **odataText** control to check the REST query and other information.</span></span>
    
- <span data-ttu-id="bafa3-p149">Запустите надстройку еще раз и выберите профиль локального компьютера с помощью диалогового окна **Вход** во время запуска Project. Откройте локальный MPP-файл и протестируйте надстройку. Убедитесь, что она отображает сообщение об ошибке при попытке получить конечную точку **ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p149">Run the add-in again, where you choose the local computer profile in the  **Login** dialog box when Project starts. Open a local .mpp file, and then test the add-in. Verify that the add-in displays an error message when you try to get the **ProjectData** endpoint.</span></span>
    
- <span data-ttu-id="bafa3-p150">Запустите надстройку еще раз и создайте проект, содержащий задачи с данными о материальных и трудовых затратах. Этот проект можно сохранить в Project Web App, но не публиковать. Убедитесь, что надстройка отображает данные с Project Server, но показывает **NA** для текущего проекта.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p150">Run the add-in again, where you create a project that has tasks with cost and work data. You can save the project to Project Web App, but don't publish it. Verify that the add-in displays data from Project Server, but  **NA** for the current project.</span></span>
    

### <a name="procedure-6-to-test-the-add-in"></a><span data-ttu-id="bafa3-p151">Процедура 6. Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="bafa3-p151">Procedure 6. To test the add-in</span></span>

1. <span data-ttu-id="bafa3-p152">Запустите Project профессиональный 2013, подключитесь к Project Web App и создайте тестовый проект. Назначьте задачи локальным ресурсам или ресурсам предприятия, настройте различные значения процента выполнения для некоторых задач и затем опубликуйте проект. Закройте Project, что позволит Visual Studio запустить Project для отладки надстройки.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p152">Run Project Professional 2013, connect with Project Web App, and then create a test project. Assign tasks to local resources or to enterprise resources, set various values of percent complete on some tasks, and then publish the project. Quit Project, which enables Visual Studio to start Project for debugging the add-in.</span></span>
    
2. <span data-ttu-id="bafa3-p153">В Visual Studio нажмите клавишу **F5**. Войдите в Project Web App и затем откройте проект, созданный на предыдущем шаге. Проект можно открыть в режиме чтения или в режиме редактирования.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p153">In Visual Studio, press  **F5**. Log on to Project Web App, and then open the project that you created in the previous step. You can open the project in read-only mode or in edit mode.</span></span>
    
3. <span data-ttu-id="bafa3-p154">На вкладке **Проект** ленты в раскрывающемся списке **Надстройки Office** выберите **Hello ProjectData** (см. рис. 5). Кнопка **Compare All Projects** (Сравнить все проекты) должна быть отключена.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p154">On the  **PROJECT** tab of the ribbon, in the **Office Add-ins** drop-down list, select **Hello ProjectData** (see Figure 5). The **Compare All Projects** button should be disabled.</span></span>
    
    <span data-ttu-id="bafa3-276">*Рис. 5. Запуск надстройки HelloProjectOData*</span><span class="sxs-lookup"><span data-stu-id="bafa3-276">*Figure 5. Starting the HelloProjectOData add-in*</span></span>

    ![Тестирование приложения HelloProjectOData](../images/pj15-hello-project-data-test-the-app.png)

4. <span data-ttu-id="bafa3-p155">В области задач **Hello ProjectData** нажмите кнопку **Get ProjectData Endpoint** (Получить конечную точку ProjectData). В строке **projectDataEndPoint** должен отображаться URL-адрес службы **ProjectData**, а кнопка **Compare All Projects** (Сравнить все проекты) должна быть включена (см. рис. 6).</span><span class="sxs-lookup"><span data-stu-id="bafa3-p155">In the  **Hello ProjectData** task pane, select **Get ProjectData Endpoint**. The  **projectDataEndPoint** line should show the URL of the **ProjectData** service, and the **Compare All Projects** button should be enabled (see Figure 6).</span></span>
    
5. <span data-ttu-id="bafa3-p156">Нажмите кнопку **Compare All Projects**. Надстройка может приостановить работу на время получения данных из службы **ProjectData**, а затем она должна отобразить отформатированные средние и текущие значения в таблице.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p156">Select  **Compare All Projects**. The add-in may pause while it retrieves data from the  **ProjectData** service, and then it should display the formatted average and current values in the table.</span></span>
    
    <span data-ttu-id="bafa3-282">*Рис. 6. Просмотр результатов запроса REST*</span><span class="sxs-lookup"><span data-stu-id="bafa3-282">*Figure 6. Viewing results of the REST query*</span></span>

    ![Просмотр результатов запроса REST](../images/pj15-hello-project-data-rest-results.png)

6. <span data-ttu-id="bafa3-p157">Проверьте выходные данные в текстовом поле. Они должны показывать путь к документу, запрос REST, сведения о состоянии и результаты JSON от вызовов **ajax** и **parseODataResult**. Выходные данные помогают понять, создать и отладить код в методе **parseODataResult**, такой как `projCost += Number(res.d.results[i].ProjectCost);`.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p157">Examine output in the text box. It should show the document path, REST query, status information, and JSON results from the calls to  **ajax** and **parseODataResult**. The output helps to understand, create, and debug code in the  **parseODataResult** method such as `projCost += Number(res.d.results[i].ProjectCost);`.</span></span>
    
    <span data-ttu-id="bafa3-287">Ниже приведен пример выходных данных для трех проектов в экземпляре Project Web App с разрывами строки и пробелами, добавленными для ясности.</span><span class="sxs-lookup"><span data-stu-id="bafa3-287">Following is an example of the output with line breaks and spaces added to the text for clarity, for three projects in a Project Web App instance:</span></span>

    ```json
    Document path: <>\WinProj test1

    REST query:
    http://sphvm-37189/pwa/_api/ProjectData/Projects?$filter=ProjectName ne 'Timesheet Administrative Work Items'
        &amp;$select=ProjectId, ProjectName, ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost
    
    textStatus: success
    ContentType: application/json;odata=verbose;charset=utf-8
    Status: 200
    
    ResponseText:
    {"d":{"results":[
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'ce3d0d65-3904-e211-96cd-00155d157123')",
        "type":"ReportingData.Project"},
        "ProjectId":"ce3d0d65-3904-e211-96cd-00155d157123",
        "ProjectActualCost":"0.000000",
        "ProjectCost":"0.000000",
        "ProjectName":"Task list created in PWA",
        "ProjectPercentCompleted":0,
        "ProjectWork":"16.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'c31023fc-1404-e211-86b2-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"c31023fc-1404-e211-86b2-3c075433b7bd",
        "ProjectActualCost":"700.000000",
        "ProjectCost":"2400.000000",
        "ProjectName":"WinProj test 2",
        "ProjectPercentCompleted":29,
        "ProjectWork":"48.000000"},
    {"__metadata":
        {"id":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "uri":"http://sphvm-37189/pwa/_api/ProjectData/Projects(guid'dc81fbb2-b801-e211-9d2a-3c075433b7bd')",
        "type":"ReportingData.Project"},
        "ProjectId":"dc81fbb2-b801-e211-9d2a-3c075433b7bd",
        "ProjectActualCost":"1900.000000",
        "ProjectCost":"5200.000000",
        "ProjectName":"WinProj test1",
        "ProjectPercentCompleted":37,
        "ProjectWork":"104.000000"}
    ]}}
    ```

7. <span data-ttu-id="bafa3-p158">Остановите отладку (нажмите клавиши **SHIFT+F5**), а затем еще раз нажмите клавишу **F5**, чтобы запустить новый экземпляр Project. В диалоговом окне **Вход** выберите локальный профиль **Компьютер**, а не Project Web App. Создайте или откройте локальный MPP-файл проекта, откройте область задач **Hello ProjectData** и нажмите кнопку **Get ProjectData Endpoint** (Получить конечную точку ProjectData). В надстройке должна появиться ошибка **No connection!** (см. рис. 7), а кнопка **Compare All Projects** (Сравнить все проекты) должна остаться отключенной.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p158">Stop debugging (press  **Shift + F5**), and then press  **F5** again to run a new instance of Project. In the **Login** dialog box, choose the local **Computer** profile, not Project Web App. Create or open a local project .mpp file, open the **Hello ProjectData** task pane, and then select **Get ProjectData Endpoint**. The add-in should show a  **No connection!** error (see Figure 7), and the **Compare All Projects** button should remain disabled.</span></span>
    
   <span data-ttu-id="bafa3-293">*Рис. 7. Использование надстройки без подключения Project Web App*</span><span class="sxs-lookup"><span data-stu-id="bafa3-293">*Figure 7. Using the add-in without a Project web app connection*</span></span>

   ![Использование приложения без подключения Project Web App](../images/pj15-hello-project-data-no-connection.png)

8. <span data-ttu-id="bafa3-p159">Остановите отладку и нажмите клавишу **F5** снова. Войдите в Project Web App и создайте проект, содержащий данные о материальных и трудовых затратах. Проект можно сохранить, но не публикуйте его.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p159">Stop debugging, and then press  **F5** again. Log on to Project Web App, and then create a project that contains cost and work data. You can save the project, but don't publish it.</span></span>
    
   <span data-ttu-id="bafa3-298">Когда вы нажимаете кнопку **Compare All Projects** (Сравнить все проекты) в области задач **Hello ProjectData**, в полях столбца **Текущее** должны появиться значения **NA**, выделенные синим цветом (см. рис. 8).</span><span class="sxs-lookup"><span data-stu-id="bafa3-298">In the  **Hello ProjectData** task pane, when you select **Compare All Projects**, you should see a blue  **NA** for fields in the **Current** column (see Figure 8).</span></span>
    
   <span data-ttu-id="bafa3-299">*Рис. 8. Сравнение неопубликованного проекта с другими проектами*</span><span class="sxs-lookup"><span data-stu-id="bafa3-299">*Figure 8. Comparing an unpublished project with other projects*</span></span>

   ![Сравнение неопубликованного проекта с другими проектами](../images/pj15-hello-project-data-not-published.png)

<span data-ttu-id="bafa3-p160">Даже если ваша надстройка работала правильно в предыдущих тестах, есть другие тесты, которые необходимо выполнить. Например:</span><span class="sxs-lookup"><span data-stu-id="bafa3-p160">Even if your add-in is working correctly in the previous tests, there are other tests that should be run. For example:</span></span>

- <span data-ttu-id="bafa3-p161">Откройте в Project Web App проект, который не содержит данных о материальных и трудовых затратах для задач. В полях столбца **Current (текущий)** должны отображаться нули.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p161">Open a project from Project Web App that has no cost or work data for the tasks. You should see values of zero in the fields in the  **Current** column.</span></span>
    
- <span data-ttu-id="bafa3-305">Протестируйте проект, не содержащий задачи.</span><span class="sxs-lookup"><span data-stu-id="bafa3-305">Test a project that has no tasks.</span></span>
    
- <span data-ttu-id="bafa3-p162">Если вы измените надстройку и опубликуете ее, необходимо запустить аналогичные тесты снова с опубликованной надстройкой. Другие вопросы см. в разделе [Дальнейшие действия](#next-steps).</span><span class="sxs-lookup"><span data-stu-id="bafa3-p162">If you modify the add-in and publish it, you should run similar tests again with the published add-in. For other considerations, see [Next steps](#next-steps).</span></span>
    

> [!NOTE]
> <span data-ttu-id="bafa3-p163">Имеются ограничения на объем данных, который может быть возвращен в одном запросе службы **ProjectData**. Это значение зависит от конкретной сущности. Например, для набора сущностей **Projects** по умолчанию действует ограничение в 100 проектов на запрос, но для набора сущностей **Risks** — 200. Для установки в рабочей среде код примера **HelloProjectOData** необходимо изменить, чтобы поддерживались запросы, содержащие более 100 проектов. Дополнительные сведения см. в разделе [Дальнейшие действия](#next-steps) и статье [Создание запросов веб-каналов OData для данных отчетов Project](http://msdn.microsoft.com/library/3eafda3b-f006-48be-baa6-961b2ed9fe01%28Office.15%29.aspx).</span><span class="sxs-lookup"><span data-stu-id="bafa3-p163">There are limits to the amount of data that can be returned in one query of the  **ProjectData** service; the amount of data varies by entity. For example, the **Projects** entity set has a default limit of 100 projects per query, but the **Risks** entity set has a default limit of 200. For a production installation, the code in the **HelloProjectOData** example should be modified to enable queries of more than 100 projects. For more information, see [Next steps](#next-steps) and [Querying OData feeds for Project reporting data](http://msdn.microsoft.com/library/3eafda3b-f006-48be-baa6-961b2ed9fe01%28Office.15%29.aspx).</span></span>


## <a name="example-code-for-the-helloprojectodata-add-in"></a><span data-ttu-id="bafa3-312">Пример кода для надстройки HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="bafa3-312">Example code for the HelloProjectOData add-in</span></span>


### <a name="helloprojectodatahtml-file"></a><span data-ttu-id="bafa3-313">Файл HelloProjectOData.html</span><span class="sxs-lookup"><span data-stu-id="bafa3-313">HelloProjectOData.html file</span></span>

<span data-ttu-id="bafa3-314">Приведенный ниже код находится в файле `Pages\HelloProjectOData.html` проекта **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-314">The following code is in the `Pages\HelloProjectOData.html` file of the **HelloProjectODataWeb** project.</span></span>

```HTML
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Test ProjectData Service</title>

        <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

        <!-- Add your CSS styles to the following file -->
        <link rel="stylesheet" type="text/css" href="../Content/App.css" />

        <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
        <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
        <script src="../Scripts/jquery-1.7.1.js"></script>

        <!-- Use the CDN reference to Office.js when deploying your add-in -->
        <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

        <!-- Use the local script references for Office.js to enable offline debugging -->
        <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
        <script src="../Scripts/Office/1.0/Office.js"></script>

        <!-- Add your JavaScript to the following files -->
        <script src="../Scripts/HelloProjectOData.js"></script>
        <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
        <div id="SectionContent">
        <div id="odataQueries">
            ODATA REST QUERY
        </div>
        <div id="odataInfo">
            <button class="button-wide" onclick="setOdataUrl()">Get ProjectData Endpoint</button>
            <br />
            <br />
            <span class="rest" id="projectDataEndPoint">Endpoint of the 
            <strong>ProjectData</strong> service</span>
            <br />
        </div>
        <div id="compareProjectData">
            <button class="button-wide" disabled="disabled" id="compareProjects"
            onclick="retrieveOData()">
            Compare All Projects</button>
            <br />
        </div>
        </div>
        <div id="corpInfo">
        <table class="infoTable" aria-readonly="True" style="width: 100%;">
            <tr>
            <td class="heading_leftCol"></td>
            <td class="heading_midCol"><strong>Average</strong></td>
            <td class="heading_rightCol"><strong>Current</strong></td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Cost</strong></td>
            <td class="row_midCol" id="AverageProjectCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Actual Cost</strong></td>
            <td class="row_midCol" id="AverageProjectActualCost">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectActualCost">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project Work</strong></td>
            <td class="row_midCol" id="AverageProjectWork">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectWork">&amp;nbsp;</td>
            </tr>
            <tr>
            <td class="row_leftCol"><strong>Project % Complete</strong></td>
            <td class="row_midCol" id="AverageProjectPercentComplete">&amp;nbsp;</td>
            <td class="row_rightCol" id="CurrentProjectPercentComplete">&amp;nbsp;</td>
            </tr>
        </table>
        </div>
        <img alt="Corporation" class="logo" src="../../images/NewLogo.png" />
        <br />
        <textarea id="odataText" rows="12" cols="40"></textarea>
    </body>
</html>
```


### <a name="helloprojectodatajs-file"></a><span data-ttu-id="bafa3-315">Файл HelloProjectOData.js</span><span class="sxs-lookup"><span data-stu-id="bafa3-315">HelloProjectOData.js file</span></span>

<span data-ttu-id="bafa3-316">Приведенный ниже код находится в файле `Scripts\Office\HelloProjectOData.js` проекта **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-316">The following code is in the `Scripts\Office\HelloProjectOData.js` file of the **HelloProjectODataWeb** project.</span></span>

```js
/* File: HelloProjectOData.js
* JavaScript functions for the HelloProjectOData example task pane app.
* October 2, 2012
*/

var PROJDATA = "/_api/ProjectData";
var PROJQUERY = "/Projects?";
var QUERY_FILTER = "$filter=ProjectName ne 'Timesheet Administrative Work Items'";
var QUERY_SELECT1 = "&amp;$select=ProjectId, ProjectName";
var QUERY_SELECT2 = ", ProjectCost, ProjectWork, ProjectPercentCompleted, ProjectActualCost";
var _pwa;           // URL of Project Web App.
var _projectUid;    // GUID of the active project.
var _docUrl;        // Path of the project document.
var _odataUrl = ""; // URL of the OData service: http[s]://ServerName /ProjectServerName /_api/ProjectData

// The initialize function is required for all add-ins.
Office.initialize = function (reason) {
    // Checks for the DOM to load using the jQuery ready function.
    $(document).ready(function () {
        // After the DOM is loaded, app-specific code can run.
    });
}

// Set the global variables, enable the Compare All Projects button,
// and display the URL of the ProjectData service.
// Display an error if Project is not connected with Project Web App.
function setOdataUrl() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.ProjectServerUrl,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _pwa = String(asyncResult.value.fieldValue);

                // If you debug with Visual Studio on a local Project Server computer, 
                // uncomment the following lines to use the localhost URL.
                //var localhost = location.host.split(":", 1);
                //var pwaStartPosition = _pwa.lastIndexOf("/");
                //var pwaLength = _pwa.length - pwaStartPosition;
                //var pwaName = _pwa.substr(pwaStartPosition, pwaLength);
                //_pwa = location.protocol + "//" + localhost + pwaName;

                if (_pwa.substring(0, 4) == "http") {
                    _odataUrl = _pwa + PROJDATA;
                    $("#compareProjects").removeAttr("disabled");
                    getProjectGuid();
                }
                else {
                    _odataUrl = "No connection!";
                    throwError(_odataUrl, "You are not connected to Project Web App.");
                }
                getDocumentUrl();
                $("#projectDataEndPoint").text(_odataUrl);
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the GUID of the active project.
function getProjectGuid() {
    Office.context.document.getProjectFieldAsync(
        Office.ProjectProjectFields.GUID,
        function (asyncResult) {
            if (asyncResult.status == Office.AsyncResultStatus.Succeeded) {
                _projectUid = asyncResult.value.fieldValue;
            }
            else {
                throwError(asyncResult.error.name, asyncResult.error.message);
            }
        }
    );
}

// Get the path of the project in Project web app, which is in the form <>\ProjectName .
function getDocumentUrl() {
    _docUrl = "Document path:\r\n" + Office.context.document.url;
}

/****************************************************************
* Functions to get and parse the Project Server reporting data.
*****************************************************************/

// Get data about all projects on Project Server, 
// by using a REST query with the ajax method in jQuery.
function retrieveOData() {
    var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
    var accept = "application/json; odata=verbose";
    accept.toLocaleLowerCase();

    // Enable cross-origin scripting (required by jQuery 1.5 and later).
    // This does not work with Project Online.
    $.support.cors = true;

    $.ajax({
        url: restUrl,
        type: "GET",
        contentType: "application/json",
        data: "",      // Empty string for the optional data.
        //headers: { "Accept": accept },
        beforeSend: function (xhr) {
            xhr.setRequestHeader("ACCEPT", accept);
        },
        complete: function (xhr, textStatus) {
            // Create a message to display in the text box.
            var message = "\r\ntextStatus: " + textStatus +
                "\r\nContentType: " + xhr.getResponseHeader("Content-Type") +
                "\r\nStatus: " + xhr.status +
                "\r\nResponseText:\r\n" + xhr.responseText;

            // xhr.responseText is the result from an XmlHttpRequest, which 
            // contains the JSON response from the OData service.
            parseODataResult(xhr.responseText, _projectUid);

            // Write the document name, response header, status, and JSON to the odataText control.
            $("#odataText").text(_docUrl);
            $("#odataText").append("\r\nREST query:\r\n" + restUrl);
            $("#odataText").append(message);

            if (xhr.status != 200 &amp;&amp; xhr.status != 1223 &amp;&amp; xhr.status != 201) {
                $("#odataInfo").append("<div>" + htmlEncode(restUrl) + "</div>");
            }
        },
        error: getProjectDataErrorHandler
    });
}

function getProjectDataErrorHandler(data, errorCode, errorMessage) {
    $("#odataText").text("Error code: " + errorCode + "\r\nError message: \r\n"
        + errorMessage);
    throwError(errorCode, errorMessage);
}

// Calculate the average values of actual cost, cost, work, and percent complete   
// for all projects, and compare with the values for the current project.
function parseODataResult(oDataResult, currentProjectGuid) {
    // Deserialize the JSON string into a JavaScript object.
    var res = Sys.Serialization.JavaScriptSerializer.deserialize(oDataResult);
    var len = res.d.results.length;
    var projActualCost = 0;
    var projCost = 0;
    var projWork = 0;
    var projPercentCompleted = 0;
    var myProjectIndex = -1;

    for (i = 0; i < len; i++) {
        // If the current project GUID matches the GUID from the OData query,  
        // then store the project index.
        if (currentProjectGuid.toLocaleLowerCase() == res.d.results[i].ProjectId) {
            myProjectIndex = i;
        }
        projCost += Number(res.d.results[i].ProjectCost);
        projWork += Number(res.d.results[i].ProjectWork);
        projActualCost += Number(res.d.results[i].ProjectActualCost);
        projPercentCompleted += Number(res.d.results[i].ProjectPercentCompleted);

    }
    var avgProjCost = projCost / len;
    var avgProjWork = projWork / len;
    var avgProjActualCost = projActualCost / len;
    var avgProjPercentCompleted = projPercentCompleted / len;

    // Round off cost to two decimal places, and round off other values to one decimal place.
    avgProjCost = avgProjCost.toFixed(2);
    avgProjWork = avgProjWork.toFixed(1);
    avgProjActualCost = avgProjActualCost.toFixed(2);
    avgProjPercentCompleted = avgProjPercentCompleted.toFixed(1);

    // Display averages in the table, with the correct units. 
    document.getElementById("AverageProjectCost").innerHTML = "$"
        + avgProjCost;
    document.getElementById("AverageProjectActualCost").innerHTML
        = "$" + avgProjActualCost;
    document.getElementById("AverageProjectWork").innerHTML
        = avgProjWork + " hrs";
    document.getElementById("AverageProjectPercentComplete").innerHTML
        = avgProjPercentCompleted + "%";

    // Calculate and display values for the current project.
    if (myProjectIndex != -1) {

        var myProjCost = Number(res.d.results[myProjectIndex].ProjectCost);
        var myProjWork = Number(res.d.results[myProjectIndex].ProjectWork);
        var myProjActualCost = Number(res.d.results[myProjectIndex].ProjectActualCost);
        var myProjPercentCompleted = Number(res.d.results[myProjectIndex].ProjectPercentCompleted);

        myProjCost = myProjCost.toFixed(2);
        myProjWork = myProjWork.toFixed(1);
        myProjActualCost = myProjActualCost.toFixed(2);
        myProjPercentCompleted = myProjPercentCompleted.toFixed(1);

        document.getElementById("CurrentProjectCost").innerHTML = "$" + myProjCost;

        if (Number(myProjCost) <= Number(avgProjCost)) {
            document.getElementById("CurrentProjectCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectCost").style.color = "red"
        }

        document.getElementById("CurrentProjectActualCost").innerHTML = "$" + myProjActualCost;

        if (Number(myProjActualCost) <= Number(avgProjActualCost)) {
            document.getElementById("CurrentProjectActualCost").style.color = "green"
        }
        else {
            document.getElementById("CurrentProjectActualCost").style.color = "red"
        }

        document.getElementById("CurrentProjectWork").innerHTML = myProjWork + " hrs";

        if (Number(myProjWork) <= Number(avgProjWork)) {
            document.getElementById("CurrentProjectWork").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectWork").style.color = "green"
        }

        document.getElementById("CurrentProjectPercentComplete").innerHTML = myProjPercentCompleted + "%";

        if (Number(myProjPercentCompleted) <= Number(avgProjPercentCompleted)) {
            document.getElementById("CurrentProjectPercentComplete").style.color = "red"
        }
        else {
            document.getElementById("CurrentProjectPercentComplete").style.color = "green"
        }
    }
    else {    // The current project is not published.
        document.getElementById("CurrentProjectCost").innerHTML = "NA";
        document.getElementById("CurrentProjectCost").style.color = "blue"

        document.getElementById("CurrentProjectActualCost").innerHTML = "NA";
        document.getElementById("CurrentProjectActualCost").style.color = "blue"

        document.getElementById("CurrentProjectWork").innerHTML = "NA";
        document.getElementById("CurrentProjectWork").style.color = "blue"

        document.getElementById("CurrentProjectPercentComplete").innerHTML = "NA";
        document.getElementById("CurrentProjectPercentComplete").style.color = "blue"
    }
}
```

### <a name="appcss-file"></a><span data-ttu-id="bafa3-317">Файл App.css</span><span class="sxs-lookup"><span data-stu-id="bafa3-317">App.css file</span></span>

<span data-ttu-id="bafa3-318">Приведенный ниже код находится в файле `Content\App.css` проекта **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="bafa3-318">The following code is in the `Content\App.css` file of the **HelloProjectODataWeb** project.</span></span>

```css
/*
*  File: App.css for the HelloProjectOData app.
*  Updated: 10/2/2012
*/
 
body
{
    font-size: 11pt;
}
h1 
{
    font-size: 22pt;
}
h2 
{
    font-size: 16pt;
}

/******************************************************************
Code label class
******************************************************************/

.rest 
{
    font-family: 'Courier New';
    font-size: 0.9em;
}

/******************************************************************
Button classes
******************************************************************/

.button-wide {
    width: 210px;
    margin-top: 2px;
}
.button-narrow 
{
    width: 80px;
    margin-top: 2px;
}

/******************************************************************
Table styles
******************************************************************/

.infoTable
{
    text-align: center; 
    vertical-align: middle
}
.heading_leftCol
{
    width: 20px;
    height: 20px;
}
.heading_midCol
{
    width: 100px;
    height: 20px;
    font-size: medium; 
    font-weight: bold; 
}
.heading_rightCol
{
    width: 101px;
    height: 20px;
    font-size: medium; 
    font-weight: bold; 
}
.row_leftCol
{
    width: 20px;
    font-size: small; 
    font-weight: bold; 
}
.row_midCol
{
    width: 100px;
}
.row_rightCol
{
    width: 101px;
}
.logo
{
    width: 135px;
    height: 53px;
}
```

### <a name="surfaceerrorsjs-file"></a><span data-ttu-id="bafa3-319">Файл SurfaceErrors.js</span><span class="sxs-lookup"><span data-stu-id="bafa3-319">SurfaceErrors.js file</span></span>

<span data-ttu-id="bafa3-320">Вы можете скопировать код для файла SurfaceErrors.js из раздела _Надежное программирование_ статьи [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span><span class="sxs-lookup"><span data-stu-id="bafa3-320">You can copy code for the SurfaceErrors.js file from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>


## <a name="next-steps"></a><span data-ttu-id="bafa3-321">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="bafa3-321">Next steps</span></span>

<span data-ttu-id="bafa3-p164">Если бы надстройка **HelloProjectOData** была рабочей надстройкой, предназначенной для продажи в AppSource или распространения в каталоге надстроек SharePoint, она конструировалась бы по-другому. Например, здесь не было бы выходных данных отладки в текстовом поле и, вероятно, не было бы кнопки для получения конечной точки **ProjectData**. Вам также следовало бы переписать функцию **retireveOData** для поддержки экземпляров Project Web App, содержащих более 100 проектов.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p164">If  **HelloProjectOData** were a production add-in to be sold in AppSource or distributed in a SharePoint add-in catalog, it would be designed differently. For example, there would be no debug output in a text box, and probably no button to get the **ProjectData** endpoint. You would also have to rewrite the **retireveOData** function to handle Project Web App instances that have more than 100 projects.</span></span>

<span data-ttu-id="bafa3-p165">Надстройка должна содержать дополнительные проверки ошибок, а также логику для записи, объяснения или демонстрации пограничных случаев. Например, если экземпляр Project Web App содержит 1000 проектов со средней продолжительностью в пять дней и средними затратами в $2400, а активный проект является единственным с продолжительностью более 20 дней, то сравнение материальных и трудовых затрат может быть перекошено. Это может быть показано с помощью частотной диаграммы. Вам необходимо добавить команды для отображения продолжительности, сравнения проектов с одинаковой продолжительностью или сравнения проектов из одного или разных отделов. Либо добавить возможность пользователю выбирать из списка полей, которые требуется отобразить.</span><span class="sxs-lookup"><span data-stu-id="bafa3-p165">The add-in should contain additional error checks, plus logic to catch and explain or show edge cases. For example, if a Project Web App instance has 1000 projects with an average duration of five days and average cost of $2400, and the active project is the only one that has a duration longer than 20 days, the cost and work comparison would be skewed. That could be shown with a frequency graph. You could add options to display duration, compare similar length projects, or compare projects from the same or different departments. Or, add a way for the user to select from a list of fields to display.</span></span>

<span data-ttu-id="bafa3-p166">Для других запросов службы **ProjectData** имеются ограничения на длину строки запроса, что влияет на число шагов, которые запрос может предпринять для выборки из родительской коллекции в объект в дочерней коллекции. Например, двухшаговый запрос **Projects** в **Tasks** для получения элементов задач работает, но трехшаговый запрос, такой как **Projects** в **Tasks** в **Assignments**, для получения элемента назначения может превысить максимальную длину URL-адреса по умолчанию. Дополнительные сведения см. в разделе [Создание запросов веб-каналов OData для данных отчетов Project](http://msdn.microsoft.com/library/3eafda3b-f006-48be-baa6-961b2ed9fe01%28Office.15%29.aspx).</span><span class="sxs-lookup"><span data-stu-id="bafa3-p166">For other queries of the  **ProjectData** service, there are limits to the length of the query string, which affects the number of steps that a query can take from a parent collection to an object in a child collection. For example, a two-step query of **Projects** to **Tasks** to task item works, but a three-step query such as **Projects** to **Tasks** to **Assignments** to assignment item may exceed the default maximum URL length. For more information, see [Querying OData feeds for Project reporting data](http://msdn.microsoft.com/library/3eafda3b-f006-48be-baa6-961b2ed9fe01%28Office.15%29.aspx).</span></span>

<span data-ttu-id="bafa3-333">Если вы изменяете надстройку **HelloProjectOData** для использования в рабочей среде, выполните следующие действия.</span><span class="sxs-lookup"><span data-stu-id="bafa3-333">If you modify the  **HelloProjectOData** add-in for production use, do the following steps:</span></span>

- <span data-ttu-id="bafa3-334">В файле HelloProjectOData.html для лучшей производительности измените ссылку office.js из локального проекта на ссылку CDN:</span><span class="sxs-lookup"><span data-stu-id="bafa3-334">In the HelloProjectOData.html file, for better performance, change the office.js reference from the local project to the CDN reference:</span></span>
    
    ```HTML
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

- <span data-ttu-id="bafa3-p167">Перепишите функцию **retrieveOData** для разрешения запросов, обрабатывающих более 100 проектов. Например, можно получить число проектов с помощью запроса `~/ProjectData/Projects()/$count` и получать данные проекта с помощью операторов _$skip_ и _$top_ в запросе REST. Запустите несколько запросов в цикле, а затем усредните данные из всех запросов. Каждый запрос данных проекта будет выглядеть так:</span><span class="sxs-lookup"><span data-stu-id="bafa3-p167">Rewrite the  **retrieveOData** function to enable queries of more than 100 projects. For example, you could get the number of projects with a `~/ProjectData/Projects()/$count` query, and use the _$skip_ operator and _$top_ operator in the REST query for project data. Run multiple queries in a loop, and then average the data from each query. Each query for project data would be of the form:</span></span> 

  `~/ProjectData/Projects()?skip= [numSkipped]&amp;$top=100&amp;$filter=[filter]&amp;$select=[field1,field2, ???????]`
    
  <span data-ttu-id="bafa3-p168">Дополнительные сведения см. в статье [OData System Query Options Using the REST Endpoint](http://msdn.microsoft.com/library/8a938b9b-7fdb-45a3-a04c-4d2d5cf2e353.aspx). Также можно использовать команду [Set-SPProjectOdataConfiguration](http://technet.microsoft.com/library/jj219516%28v=office.15%29.aspx) в Windows PowerShell, чтобы переопределить размер страницы по умолчанию для запроса набора сущностей **Projects** (или любого другого из 33 наборов сущностей). См. [ProjectData — Справочник по службе Project OData](http://msdn.microsoft.com/library/1ed14ee9-1a1a-4960-9b66-c24ef92cdf6b%28Office.15%29.aspx).</span><span class="sxs-lookup"><span data-stu-id="bafa3-p168">For more information, see [OData System Query Options Using the REST Endpoint](http://msdn.microsoft.com/library/8a938b9b-7fdb-45a3-a04c-4d2d5cf2e353.aspx). You can also use the [Set-SPProjectOdataConfiguration](http://technet.microsoft.com/library/jj219516%28v=office.15%29.aspx) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](http://msdn.microsoft.com/library/1ed14ee9-1a1a-4960-9b66-c24ef92cdf6b%28Office.15%29.aspx).</span></span>
    
- <span data-ttu-id="bafa3-342">Сведения о развертывании надстройки см. в статье [Публикация надстройки Office](../publish/publish.md).</span><span class="sxs-lookup"><span data-stu-id="bafa3-342">To deploy the add-in, see [Publish your Office Add-in](../publish/publish.md).</span></span>
    

## <a name="see-also"></a><span data-ttu-id="bafa3-343">См. также</span><span class="sxs-lookup"><span data-stu-id="bafa3-343">See also</span></span>

- [<span data-ttu-id="bafa3-344">Надстройки области задач для Project</span><span class="sxs-lookup"><span data-stu-id="bafa3-344">Task pane add-ins for Project</span></span>](project-add-ins.md)
- [<span data-ttu-id="bafa3-345">Создание первой надстройки области задач для Project 2013 с помощью текстового редактора</span><span class="sxs-lookup"><span data-stu-id="bafa3-345">Create your first task pane add-in for Project 2013 by using a text editor</span></span>](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- [<span data-ttu-id="bafa3-346">ProjectData — Справочник по службе Project OData</span><span class="sxs-lookup"><span data-stu-id="bafa3-346">ProjectData - Project OData service reference</span></span>](http://msdn.microsoft.com/library/1ed14ee9-1a1a-4960-9b66-c24ef92cdf6b%28Office.15%29.aspx) 
- [<span data-ttu-id="bafa3-347">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="bafa3-347">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md) 
- [<span data-ttu-id="bafa3-348">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="bafa3-348">Publish your Office Add-in</span></span>](../publish/publish.md)
    
