---
title: Создание надстройки Project, использующей REST с локальной службой OData Project Server
description: Узнайте, как создать надстройку области задач для Project профессиональный 2013 г., которая сравнивает данные о затратах и работе в активном проекте со средними значениями для всех проектов в текущем экземпляре Project Web App.
ms.date: 09/26/2019
localization_priority: Normal
ms.openlocfilehash: c03cd580f9f5d4da654022de811d4a060a99e52d
ms.sourcegitcommit: 883f71d395b19ccfc6874a0d5942a7016eb49e2c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/09/2021
ms.locfileid: "53348813"
---
# <a name="create-a-project-add-in-that-uses-rest-with-an-on-premises-project-server-odata-service"></a><span data-ttu-id="3e6f5-103">Создание надстройки Project, использующей REST с локальной службой OData Project Server</span><span class="sxs-lookup"><span data-stu-id="3e6f5-103">Create a Project add-in that uses REST with an on-premises Project Server OData service</span></span>

<span data-ttu-id="3e6f5-104">В этой статье описывается создание надстройки области задач для Project профессиональный 2013, которая сравнивает данные по материальным и трудовым затратам в активном проекте со средними значениями из всех проектов в текущем экземпляре Project Web App.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-104">This article describes how to build a task pane add-in for Project Professional 2013 that compares cost and work data in the active project with the averages for all projects in the current Project Web App instance.</span></span> <span data-ttu-id="3e6f5-105">Надстройка использует REST с библиотекой jQuery для доступа к службе отчетов **ProjectData** OData в Project Server 2013.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-105">The add-in uses REST with the jQuery library to access the **ProjectData** OData reporting service in Project Server 2013.</span></span>

<span data-ttu-id="3e6f5-106">Код в данной статье основан на примере, разработанном Саурабхом Сангхви (Saurabh Sanghvi) и Эрвиндом Лаиром (Arvind Iyer), сотрудниками корпорации Майкрософт.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-106">The code in this article is based on a sample developed by Saurabh Sanghvi and Arvind Iyer, Microsoft Corporation.</span></span>

## <a name="prerequisites-for-creating-a-task-pane-add-in-that-reads-project-server-reporting-data"></a><span data-ttu-id="3e6f5-107">Необходимые условия для создания надстроек области задач, читающей данные отчетов Project Server</span><span class="sxs-lookup"><span data-stu-id="3e6f5-107">Prerequisites for creating a task pane add-in that reads Project Server reporting data</span></span>

<span data-ttu-id="3e6f5-108">Ниже приводится условие создания надстройки Project области задач, которая читает службу **ProjectData** экземпляра Project Web App в локальной установке Project Server 2013.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-108">The following are the prerequisites for creating a Project task pane add-in that reads the **ProjectData** service of a Project Web App instance in an on-premises installation of Project Server 2013.</span></span>

- <span data-ttu-id="3e6f5-p102">Проверьте, что на локальном компьютере разработчика установлены самые последние пакеты обновления и обновления Windows. Операционной системой может быть Windows 7, Windows 8, Windows Server 2008 или Windows Server 2012.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p102">Ensure that you have installed the most recent service packs and Windows updates on your local development computer. The operating system can be Windows 7, Windows 8, Windows Server 2008, or Windows Server 2012.</span></span>

- <span data-ttu-id="3e6f5-111">Project профессиональный 2013 требуется для подключения к Project Web App.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-111">Project Professional 2013 is required to connect with Project Web App.</span></span> <span data-ttu-id="3e6f5-112">Компьютер разработки должен иметь Project профессиональный 2013, чтобы включить **отладку F5** с Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-112">The development computer must have Project Professional 2013 installed to enable **F5** debugging with Visual Studio.</span></span>

    > [!NOTE]
    > <span data-ttu-id="3e6f5-113">Project стандартный 2013 г. также могут принимать надстройки области задач, но не могут войти в Project Web App.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-113">Project Standard 2013 can also host task pane add-ins, but cannot sign in to Project Web App.</span></span>

- <span data-ttu-id="3e6f5-114">Visual Studio 2015 с Инструменты разработчика Office для Visual Studio содержит шаблоны, позволяющие создавать Надстройки Office и SharePoint. Убедитесь, что у вас установлена самая последняя версия Office Developer Tools. См. раздел _Средства_ статьи [Надстройки Office и скачиваемые файлы для SharePoint](https://developer.microsoft.com/office/docs).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-114">Visual Studio 2015 with Office Developer Tools for Visual Studio includes templates for creating Office and SharePoint Add-ins. Ensure that you have installed the most recent version of Office Developer Tools; see the  _Tools_ section of the [Office Add-ins and SharePoint downloads](https://developer.microsoft.com/office/docs).</span></span>

- <span data-ttu-id="3e6f5-115">В примерах процедур и кода в этой статье можно получить доступ к службе **ProjectData** Project Server 2013 в локальном домене.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-115">The procedures and code examples in this article access the **ProjectData** service of Project Server 2013 in a local domain.</span></span> <span data-ttu-id="3e6f5-116">Методы jQuery в этой статье не работают с Project в Интернете.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-116">The jQuery methods in this article do not work with Project on the web.</span></span>

    <span data-ttu-id="3e6f5-117">Убедитесь, **что служба ProjectData** доступна с компьютера разработки.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-117">Verify that the **ProjectData** service is accessible from your development computer.</span></span>

### <a name="procedure-1-to-verify-that-the-projectdata-service-is-accessible"></a><span data-ttu-id="3e6f5-p105">Процедура 1. Проверка доступности службы ProjectData</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p105">Procedure 1. To verify that the ProjectData service is accessible</span></span>

1. <span data-ttu-id="3e6f5-p106">Чтобы разрешить браузеру напрямую отображать XML-данные из запроса REST, отключите вид чтения канала. Дополнительные сведения о том, как это сделать в Internet Explorer, см. в процедуру 1, шаг 4 в статье [Создание запросов веб-каналов OData для данных отчетов Project](/previous-versions/office/project-odata/jj163048(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p106">To enable your browser to directly show the XML data from a REST query, turn off the feed reading view. For information about how to do this in Internet Explorer, see Procedure 1, step 4 in [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

2. <span data-ttu-id="3e6f5-122">Запрос **службы ProjectData** с помощью браузера со следующим **http://ServerName URL-адресом: /ProjectServerName /_api/ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-122">Query the **ProjectData** service by using your browser with the following URL: **http://ServerName /ProjectServerName /_api/ProjectData**.</span></span> <span data-ttu-id="3e6f5-123">Например, если экземпляр `http://MyServer/pwa` Project Web App, браузер показывает следующие результаты.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-123">For example, if the Project Web App instance is  `http://MyServer/pwa`, the browser shows the following results.</span></span>

    ```xml
    <?xml version="1.0" encoding="utf-8"?>
        <service xml:base="http://myserver/pwa/_api/ProjectData/"
        xmlns="https://www.w3.org/2007/app"
        xmlns:atom="https://www.w3.org/2005/Atom">
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

3. <span data-ttu-id="3e6f5-p108">Вам может потребоваться предоставить свои сетевые учетные данные, чтобы увидеть результаты. Если браузер показывает сообщение "Ошибка 403, доступ запрещен", то либо у вас либо нет разрешений на вход для заданного экземпляра Project Web App, либо имеется проблема сети, требующая помощи администратора.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p108">You may have to provide your network credentials to see the results. If the browser shows "Error 403, Access Denied," either you do not have logon permission for that Project Web App instance, or there is a network problem that requires administrative help.</span></span>

## <a name="using-visual-studio-to-create-a-task-pane-add-in-for-project"></a><span data-ttu-id="3e6f5-126">Создание надстройки области задач для Project с помощью Visual Studio</span><span class="sxs-lookup"><span data-stu-id="3e6f5-126">Using Visual Studio to create a task pane add-in for Project</span></span>

<span data-ttu-id="3e6f5-127">Инструменты разработчика Office для Visual Studio включает шаблон надстроек области задач для Project 2013.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-127">Office Developer Tools for Visual Studio includes a template for task pane add-ins for Project 2013.</span></span> <span data-ttu-id="3e6f5-128">Если вы создаете решение с именем **HelloProjectOData,** решение содержит следующие два Visual Studio:</span><span class="sxs-lookup"><span data-stu-id="3e6f5-128">If you create a solution named **HelloProjectOData**, the solution contains the following two Visual Studio projects:</span></span>

- <span data-ttu-id="3e6f5-129">Проект надстройки получает имя решения.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-129">The add-in project takes the name of the solution.</span></span> <span data-ttu-id="3e6f5-130">Оно включает в себя XML-файл манифеста для приложения и настраивается на целевую платформу .NET Framework 4.5.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-130">It includes the XML manifest file for the add-in and targets the .NET Framework 4.5.</span></span> <span data-ttu-id="3e6f5-131">В процедуре 3 показаны действия по изменению манифеста надстройки **HelloProjectOData.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-131">Procedure 3 shows the steps to modify the manifest for the **HelloProjectOData** add-in.</span></span>

- <span data-ttu-id="3e6f5-132">Веб-проект называется **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-132">The web project is named **HelloProjectODataWeb**.</span></span> <span data-ttu-id="3e6f5-133">Оно содержит файлы JavaScript веб-страниц, файлы CSS, рисунки, ссылки и файлы конфигурации для веб-контента в области задач.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-133">It includes the webpages, JavaScript files, CSS files, images, references, and configuration files for the web content in the task pane.</span></span> <span data-ttu-id="3e6f5-134">Веб-проект настраивается на конечную платформу .NET Framework 4.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-134">The web project targets the .NET Framework 4.</span></span> <span data-ttu-id="3e6f5-135">В процедуре 4 и процедуре 5 показано, как изменить эти файлы в веб-проекте, чтобы создать функциональность надстройки **HelloProjectOData**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-135">Procedure 4 and Procedure 5 show how to modify the files in the web project to create the functionality of the **HelloProjectOData** add-in.</span></span>

### <a name="procedure-2-to-create-the-helloprojectodata-add-in-for-project"></a><span data-ttu-id="3e6f5-p112">Процедура 2. Создание надстройки HelloProjectOData для Project</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p112">Procedure 2. To create the HelloProjectOData add-in for Project</span></span>

1. <span data-ttu-id="3e6f5-138">Запустите Visual Studio 2015 в качестве администратора, а затем выберите **new Project** на странице Начните.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-138">Run Visual Studio 2015 as an administrator, and then select **New Project** on the Start page.</span></span>

2. <span data-ttu-id="3e6f5-139">В **диалоговом окне Project** расширения шаблонов, визуальных C# **и** **Office/SharePoint,** а затем выберите Office надстройки .  Выберите **платформа .NET Framework 4.5.2** в выпадаемом списке целевых рамок в верхней части области центра, а затем выберите Office надстройку **(см.** следующий скриншот).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-139">In the **New Project** dialog box, expand the **Templates**, **Visual C#**, and **Office/SharePoint** nodes, and then select **Office Add-ins**. Select **.NET Framework 4.5.2** in the target framework drop-down list at the top of the center pane, and then select **Office Add-in** (see the next screenshot).</span></span>

3. <span data-ttu-id="3e6f5-140">Чтобы разместить оба проекта Visual Studio в одной папке, выберите **Создать каталог для решения** и найдите требуемое расположение.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-140">To place both of the Visual Studio projects in the same directory, select **Create directory for solution**, and then browse to the location you want.</span></span>

4. <span data-ttu-id="3e6f5-141">В поле **Name** введитеHelloProjectOData и выберите **ОК.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-141">In the **Name** field, typeHelloProjectOData, and then choose **OK**.</span></span>

    <span data-ttu-id="3e6f5-142">*Рис. 1. Создание надстройки Office*</span><span class="sxs-lookup"><span data-stu-id="3e6f5-142">*Figure 1. Creating an Office Add-in*</span></span>

    ![Создание Office надстройки.](../images/pj15-hello-project-o-data-creating-app.png)

5. <span data-ttu-id="3e6f5-144">В диалоговом окне **Выбор типа надстройки** выберите пункт **Надстройка области задач** и нажмите кнопку **Далее** (см. следующий снимок экрана).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-144">In the **Choose the add-in type** dialog box, select **Task pane** and choose **Next** (see the next screenshot).</span></span>

    <span data-ttu-id="3e6f5-145">*Рис. 2. Выбор типа создаваемой надстройки*</span><span class="sxs-lookup"><span data-stu-id="3e6f5-145">*Figure 2. Choosing the type of add-in to create*</span></span>

    ![Выбор типа надстройки для создания.](../images/pj15-hello-project-o-data-choose-project.png)

6. <span data-ttu-id="3e6f5-147">В диалоговом окне **Выбор ведущих приложений** снимите все флажки, кроме флажка **Project** (см. следующий снимок экрана), а затем нажмите кнопку **Готово**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-147">In the **Choose the host applications** dialog box, clear all check boxes except the **Project** check box (see the next screenshot) and choose **Finish**.</span></span>

    <span data-ttu-id="3e6f5-148">*Рис. 3. Выбор ведущего приложения*</span><span class="sxs-lookup"><span data-stu-id="3e6f5-148">*Figure 3. Choosing the host application*</span></span>

    ![Выбор Project как единственное хост-приложение.](../images/create-office-add-in.png)

    <span data-ttu-id="3e6f5-150">Visual Studio проект **HelloProjectOdata** и **проект HelloProjectODataWeb.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-150">Visual Studio creates the **HelloProjectOdata** project and the **HelloProjectODataWeb** project.</span></span>

<span data-ttu-id="3e6f5-151">Папка **AddIn** (см. следующий скриншот) содержит файл App.css для пользовательских стилей CSS.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-151">The **AddIn** folder (see the next screenshot) contains the App.css file for custom CSS styles.</span></span> <span data-ttu-id="3e6f5-152">Во вложенной папке **Home** находится файл Home.html, содержащий ссылки на CSS-файлы и файлы JavaScript, используемые надстройкой, а также содержимое HTML5 для этой надстройки.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-152">In the **Home** subfolder , the Home.html file contains references to the CSS files and the JavaScript files that the add-in uses, and the HTML5 content for the add-in.</span></span> <span data-ttu-id="3e6f5-153">Также в ней располагается файл Home.js, предназначенный для настраиваемого кода JavaScript.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-153">Also, the Home.js file is for your custom JavaScript code.</span></span> <span data-ttu-id="3e6f5-154">Папка **Scripts** содержит файлы библиотеки jQuery.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-154">The **Scripts** folder includes the jQuery library files.</span></span> <span data-ttu-id="3e6f5-155">Во вложенной папке **Office** находятся библиотеки JavaScript, например office.js и project-15.js, а также языковые библиотеки для стандартных строк в надстройках Office. В папке **Content** находится файл Office.css, содержащий стили по умолчанию для всех надстроек Office.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-155">The **Office** subfolder includes the JavaScript libraries such as office.js and project-15.js, plus the language libraries for standard strings in the Office Add-ins. In the **Content** folder, the Office.css file contains the default styles for all of the Office Add-ins.</span></span>

<span data-ttu-id="3e6f5-156">*Рис. 4. Просмотр файлов веб-проекта по умолчанию в обозревателе решений*</span><span class="sxs-lookup"><span data-stu-id="3e6f5-156">*Figure 4. Viewing the default web project files in Solution Explorer*</span></span>

![Просмотр файлов веб-проекта в Expl решения.](../images/pj15-hello-project-o-data-initial-solution-explorer.png)

<span data-ttu-id="3e6f5-158">Манифест проекта **HelloProjectOData** — это файл HelloProjectOData.xml.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-158">The manifest for the **HelloProjectOData** project is the HelloProjectOData.xml file.</span></span> <span data-ttu-id="3e6f5-159">Его можно изменить при необходимости, чтобы добавить описание надстройки, ссылку на значок, сведения о дополнительных языках и другие параметры.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-159">You can optionally modify the manifest to add a description of the add-in, a reference to an icon, information for additional languages, and other settings.</span></span> <span data-ttu-id="3e6f5-160">В процедуре 3 изменяется только отображаемое имя надстройки и описание и добавляется значок.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-160">Procedure 3 simply modifies the add-in display name and description, and adds an icon.</span></span>

<span data-ttu-id="3e6f5-161">Дополнительные сведения о манифесте см. в статьях [XML-манифест надстроек для Office](../develop/add-in-manifests.md) и [Справка по схеме для манифестов надстроек Office (версия 1.1)](../develop/add-in-manifests.md#see-also).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-161">For more information about the manifest, see [Office Add-ins XML manifest](../develop/add-in-manifests.md) and [Schema reference for Office Add-ins manifests (v1.1)](../develop/add-in-manifests.md#see-also).</span></span>

### <a name="procedure-3-to-modify-the-add-in-manifest"></a><span data-ttu-id="3e6f5-p115">Процедура 3. Изменение манифеста надстройки</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p115">Procedure 3. To modify the add-in manifest</span></span>

1. <span data-ttu-id="3e6f5-164">Откройте файл HelloProjectOData.xml в Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-164">In Visual Studio, open the HelloProjectOData.xml file.</span></span>

2. <span data-ttu-id="3e6f5-165">Отображаемое имя по умолчанию — это имя проекта Visual Studio ("HelloProjectOData").</span><span class="sxs-lookup"><span data-stu-id="3e6f5-165">The default display name is the name of the Visual Studio project ("HelloProjectOData").</span></span> <span data-ttu-id="3e6f5-166">Например, измените значение элемента **DisplayName** по умолчанию на "Hello ProjectData".</span><span class="sxs-lookup"><span data-stu-id="3e6f5-166">For example, change the default value of the **DisplayName** element to"Hello ProjectData".</span></span>

3. <span data-ttu-id="3e6f5-p117">Описание по умолчанию — "HelloProjectOData". Например, измените значение по умолчанию элемента Description на "Test REST queries of the ProjectData service" (тестирование запросов REST службы ProjectData).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p117">The default description is also "HelloProjectOData". For example, change the default value of the Description element to "Test REST queries of the ProjectData service".</span></span>

4. <span data-ttu-id="3e6f5-169">Добавьте значок для отображения в раскрывающемся списке **Надстройки Office** на вкладке **PROJECT** ленты.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-169">Add an icon to show in the **Office Add-ins** drop-down list on the **PROJECT** tab of the ribbon.</span></span> <span data-ttu-id="3e6f5-170">Можно добавить файл значка в решении Visual Studio или использовать URL-адрес значка.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-170">You can add an icon file in the Visual Studio solution or use a URL for an icon.</span></span> 

<span data-ttu-id="3e6f5-171">В следующих действиях покажите, как добавить файл значка в Visual Studio решение.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-171">The following steps show how to add an icon file to the Visual Studio solution.</span></span>

1. <span data-ttu-id="3e6f5-172">В **обозревателе решений** перейдите в папку с именем Images.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-172">In **Solution Explorer**, go to the folder named Images.</span></span>

2. <span data-ttu-id="3e6f5-173">Чтобы отображаться в раскрывающемся списке **Надстройки Office**, значок должен иметь размер 32 x 32 пикселя.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-173">To be displayed in the **Office Add-ins** drop-down list, the icon must be 32 x 32 pixels.</span></span> <span data-ttu-id="3e6f5-174">Например, установите пакет SDK Project 2013, затем выберите папку **Images** и добавьте следующий файл из пакета SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`</span><span class="sxs-lookup"><span data-stu-id="3e6f5-174">For example, install the Project 2013 SDK, and then choose the **Images** folder and add the following file from the SDK: `\Samples\Apps\HelloProjectOData\HelloProjectODataWeb\Images\NewIcon.png`</span></span>

    <span data-ttu-id="3e6f5-175">Поочередно используйте собственный значок 32 x 32; или скопируйте следующее изображение в файл с именем NewIcon.png, а затем добавьте этот файл в  `HelloProjectODataWeb\Images` папку.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-175">Alternately, use your own 32 x 32 icon; or, copy the following image to a file named NewIcon.png, and then add that file to the  `HelloProjectODataWeb\Images` folder.</span></span>

    ![Значок для приложения HelloProjectOData.](../images/pj15-hello-project-data-new-icon.jpg)

3. <span data-ttu-id="3e6f5-177">В манифесте HelloProjectOData.xml добавьте элемент **IconUrl**  ниже элемента Description, где значение URL-адреса значка является относительным путем к файлу значка 32x32.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-177">In the HelloProjectOData.xml manifest, add an **IconUrl** element below the **Description** element, where the value of the icon URL is the relative path to the 32x32 icon file.</span></span> <span data-ttu-id="3e6f5-178">Например, добавьте следующую строку: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />** .</span><span class="sxs-lookup"><span data-stu-id="3e6f5-178">For example, add the following line: **<IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />**.</span></span> <span data-ttu-id="3e6f5-179">Файл манифеста HelloProjectOData.xml теперь содержит (ваше значение **Id** будет другим):</span><span class="sxs-lookup"><span data-stu-id="3e6f5-179">The HelloProjectOData.xml manifest file now contains the following (your **Id** value will be different):</span></span>

    ```XML
    <?xml version="1.0" encoding="UTF-8"?>
    <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
            xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance" xsi:type="TaskPaneApp">
        <!--IMPORTANT! Id must be unique for each add-in. If you copy this manifest ensure that you change this id to your own GUID. -->
        <Id>c512df8d-a1c5-4d74-8a34-d30f6bbcbd82</Id>
        <Version>1.0</Version>
        <ProviderName> [Provider name]</ProviderName>
        <DefaultLocale>en-US</DefaultLocale>
        <DisplayName DefaultValue="Hello ProjectData" />
        <Description DefaultValue="Test REST queries of the ProjectData service"/>
        <IconUrl DefaultValue="~remoteAppUrl/Images/NewIcon.png" />
        <SupportUrl DefaultValue="[Insert the URL of a page that provides support information for the app]" />
        <Hosts>
            <Host Name="Project" />
        </Hosts>
        <DefaultSettings>
            <SourceLocation DefaultValue="~remoteAppUrl/AddIn/Home/Home.html" />
        </DefaultSettings>
        <Permissions>ReadWriteDocument</Permissions>
    </OfficeApp>
    ```

## <a name="creating-the-html-content-for-the-helloprojectodata-add-in"></a><span data-ttu-id="3e6f5-180">Создание HTML-контента для надстройки HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="3e6f5-180">Creating the HTML content for the HelloProjectOData add-in</span></span>

<span data-ttu-id="3e6f5-181">**Надстройка HelloProjectOData** — это пример, который включает отладку и выход ошибок; она не предназначена для производственного использования.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-181">The **HelloProjectOData** add-in is a sample that includes debugging and error output; it is not intended for production use.</span></span> <span data-ttu-id="3e6f5-182">Прежде чем приступить к кодированию HTML-контента, разработать пользовательский интерфейс и пользовательский интерфейс для надстройки и наметить функции JavaScript, взаимодействующие с HTML-кодом.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-182">Before you start coding the HTML content, design the UI and user experience for the add-in, and outline the JavaScript functions that interact with the HTML code.</span></span> <span data-ttu-id="3e6f5-183">Дополнительные сведения см. в руководстве по разработке[Office надстройки.](../design/add-in-design.md)</span><span class="sxs-lookup"><span data-stu-id="3e6f5-183">For more information, see[Design guidelines for Office Add-ins](../design/add-in-design.md).</span></span> 

<span data-ttu-id="3e6f5-184">На области задач отображается имя отображения надстройки в верхней части, которое является значением элемента **DisplayName** в манифесте.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-184">The task pane shows the add-in display name at the top, which is the value of the **DisplayName** element in the manifest.</span></span> <span data-ttu-id="3e6f5-185">Элемент **body** в файле HelloProjectOData.html содержит другие элементы пользовательского интерфейса:</span><span class="sxs-lookup"><span data-stu-id="3e6f5-185">The **body** element in the HelloProjectOData.html file contains the other UI elements, as follows:</span></span>

- <span data-ttu-id="3e6f5-186">Подзаголовок, указывающий на общую функциональность или тип работы, например: **ODATA REST QUERY**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-186">A subtitle indicates the general functionality or type of operation, for example, **ODATA REST QUERY**.</span></span>

- <span data-ttu-id="3e6f5-187">Кнопка **Get ProjectData Endpoint** вызывает функцию, чтобы получить конечную точку службы ProjectData и отобразить ее `setOdataUrl` в текстовом окне. </span><span class="sxs-lookup"><span data-stu-id="3e6f5-187">The **Get ProjectData Endpoint** button calls the `setOdataUrl` function to get the endpoint of the **ProjectData** service, and display it in a text box.</span></span> <span data-ttu-id="3e6f5-188">Если Project не подключен к Project Web App, надстройка вызовет обработчик ошибок для отображения всплывающего сообщения об ошибке.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-188">If Project is not connected with Project Web App, the add-in calls an error handler to display a pop-up error message.</span></span>

- <span data-ttu-id="3e6f5-189">Кнопка **Compare All Projects** отключена до тех пор, пока надстройка не получит действительную конечную точку OData.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-189">The **Compare All Projects** button is disabled until the add-in gets a valid OData endpoint.</span></span> <span data-ttu-id="3e6f5-190">При выборе кнопки она вызывает функцию, которая использует запрос REST для получения данных о стоимости проекта и работе из `retrieveOData` **службы ProjectData.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-190">When you select the button, it calls the `retrieveOData` function, which uses a REST query to get project cost and work data from the **ProjectData** service.</span></span>

- <span data-ttu-id="3e6f5-p125">Таблица отображает средние значения затрат проекта, фактических затрат, трудозатрат и процент выполнения. В таблице также сравниваются значения текущего активного проекта со средними. Если текущее значение больше среднего по всем проектам, значение отображается красным цветом. Если текущее значение меньше среднего, оно отображается зеленым цветом. Если текущее значение недоступно, в таблице отображается значение **NA** синим цветом.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p125">A table displays the average values for project cost, actual cost, work, and percent complete. The table also compares the current active project values with the average. If the current value is greater than the average for all projects, the value is displayed as red. If the current value is less than the average, the value is displayed as green. If the current value is not available, the table displays a blue **NA**.</span></span>

    <span data-ttu-id="3e6f5-196">Функция `retrieveOData` вызывает `parseODataResult` функцию, которая вычисляет и отображает значения для таблицы.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-196">The `retrieveOData` function calls the `parseODataResult` function, which calculates and displays values for the table.</span></span>

    > [!NOTE]
    > <span data-ttu-id="3e6f5-197">В данном примере данные о материальных и трудовых затратах по активному проекту извлекаются из опубликованных значений.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-197">In this example, cost and work data for the active project are derived from the published values.</span></span> <span data-ttu-id="3e6f5-198">Если изменить значения в Project, служба **ProjectData** не будет знать об изменениях до тех пор, пока проект не опубликован.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-198">If you change values in Project, the **ProjectData** service does not have the changes until the project is published.</span></span>

### <a name="procedure-4-to-create-the-html-content"></a><span data-ttu-id="3e6f5-p127">Процедура 4. Создание HTML-контента</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p127">Procedure 4. To create the HTML content</span></span>

1. <span data-ttu-id="3e6f5-201">В элементе **head** файла Home.html добавьте любые дополнительные элементы **link** для CSS-файлов, используемых в надстройке.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-201">In the **head** element of the Home.html file, add any additional **link** elements for CSS files that your add-in uses.</span></span> <span data-ttu-id="3e6f5-202">Шаблон проекта Visual Studio содержит ссылку на файл App.css, который можно использовать для настраиваемых стилей CSS.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-202">The Visual Studio project template includes a link for the App.css file that you can use for custom CSS styles.</span></span>

2. <span data-ttu-id="3e6f5-203">Добавьте дополнительные **элементы** скрипта для библиотек JavaScript, которые использует ваша надстройка.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-203">Add any additional **script** elements for JavaScript libraries that your add-in uses.</span></span> <span data-ttu-id="3e6f5-204">Шаблон проекта включает ссылки на файлы _jQuery-[version]_.js, office.js и MicrosoftAjax.js в папке **Scripts.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-204">The project template includes links for the jQuery- _[version]_.js, office.js, and MicrosoftAjax.js files in the **Scripts** folder.</span></span>

    > [!NOTE]
    > <span data-ttu-id="3e6f5-p130">Перед развертыванием надстройки измените ссылку office.js и ссылку jQuery на ссылку сети доставки содержимого (CDN). Ссылка CDN предоставляет самую последнюю версию и обеспечивает оптимальную производительность.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p130">Before you deploy the add-in, change the office.js reference and the jQuery reference to the content delivery network (CDN) reference. The CDN reference provides the most recent version and better performance.</span></span>

    <span data-ttu-id="3e6f5-207">Надстройка **HelloProjectOData** использует файл SurfaceErrors.js, который отображает ошибки и всплывающее сообщение.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-207">The **HelloProjectOData** add-in also uses the SurfaceErrors.js file, which displays errors in a pop-up message.</span></span> <span data-ttu-id="3e6f5-208">Вы можете скопировать  код из раздела Надежное программирование в разделе Создание надстройки области задач для [Project 2013](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)г. с помощью текстового редактора, а затем добавить файл SurfaceErrors.js в папку **Scripts\Office** проекта **HelloProjectODataWeb.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-208">You can copy the code from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md), and then add a SurfaceErrors.js file in the **Scripts\Office** folder of the **HelloProjectODataWeb** project.</span></span>

    <span data-ttu-id="3e6f5-209">Ниже приводится обновленный HTML-код для главного элемента с дополнительной строкой для SurfaceErrors.js файла. </span><span class="sxs-lookup"><span data-stu-id="3e6f5-209">Following is the updated HTML code for the **head** element, with the additional line for the SurfaceErrors.js file.</span></span>

    ```HTML
    <!DOCTYPE html>
    <html>
    <head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <title>Test ProjectData Service</title>

    <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

    <!-- Add your CSS styles to the following file. -->
    <link rel="stylesheet" type="text/css" href="../Content/App.css" />

    <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
    <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
    <script src="../Scripts/jquery-1.7.1.js"></script>

    <!-- Use the CDN reference to office.js when deploying your add-in. -->
    <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

    <!-- Use the local script references for Office.js to enable offline debugging -->
    <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
    <script src="../Scripts/Office/1.0/Office.js"></script>

    <!-- Add your JavaScript to the following files. -->
    <script src="../Scripts/HelloProjectOData.js"></script>
    <script src="../Scripts/SurfaceErrors.js"></script>
    </head>
    <body>
    <!-- See the code in Step 3. -->
    </body>
    </html>
    ```

3. <span data-ttu-id="3e6f5-210">В **элементе body** удалите существующий код из шаблона и добавьте код для пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-210">In the **body** element, delete the existing code from the template, and then add the code for the user interface.</span></span> <span data-ttu-id="3e6f5-211">Если элемент должен заполняться данными или изменяться оператором jQuery, элемент должен содержать уникальный атрибут **id**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-211">If an element is to be filled with data or manipulated by a jQuery statement, the element must include a unique **id** attribute.</span></span> <span data-ttu-id="3e6f5-212">В следующем коде атрибуты **id** для элементов кнопки **,** span **и** **td** (определение ячейки таблицы), которые используют функции jQuery, показаны в шрифте жирным шрифтом.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-212">In the following code, the **id** attributes for the **button**, **span**, and **td** (table cell definition) elements that jQuery functions use are shown in bold font.</span></span>

   <span data-ttu-id="3e6f5-213">Следующий HTML-код добавляет графическое изображение, которое может быть эмблемой компании.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-213">The following HTML adds a graphic image, which could be a company logo.</span></span> <span data-ttu-id="3e6f5-214">Вы можете использовать логотип по вашему выбору или скопировать NewLogo.png из скачивания SDK Project 2013 г., а затем с помощью **обозревателя** решений добавить файл в `HelloProjectODataWeb\Images` папку.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-214">You can use a logo of your choice, or copy the NewLogo.png file from the Project 2013 SDK download, and then use **Solution Explorer** to add the file to the `HelloProjectODataWeb\Images` folder.</span></span>

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

## <a name="creating-the-javascript-code-for-the-add-in"></a><span data-ttu-id="3e6f5-215">Создание кода JavaScript для надстройки</span><span class="sxs-lookup"><span data-stu-id="3e6f5-215">Creating the JavaScript code for the add-in</span></span>

<span data-ttu-id="3e6f5-216">Шаблон надстройки области задач для Project содержит код инициализации по умолчанию, который предназначен для демонстрации базовых действий получения и записи данных в документе для типичных приложений Office 2013.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-216">The template for a Project task pane add-in includes default initialization code that is designed to demonstrate basic get and set actions for data in a document for a typical Office 2013 add-in.</span></span> <span data-ttu-id="3e6f5-217">Поскольку Project 2013 не поддерживает действия, которые записывают в активный проект, а надстройка **HelloProjectOData** не использует метод, можно удалить скрипт в функции и удалить функцию и функцию в файле HelloProjectOData.js `getSelectedDataAsync` `Office.initialize` по `setData` `getData` умолчанию.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-217">Because Project 2013 does not support actions that write to the active project, and the **HelloProjectOData** add-in does not use the `getSelectedDataAsync` method, you can delete the script within the `Office.initialize` function, and delete the `setData` function and `getData` function in the default HelloProjectOData.js file.</span></span>

<span data-ttu-id="3e6f5-218">В JavaScript содержатся глобальные константы для запроса REST и глобальные переменные, используемые в нескольких функциях.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-218">The JavaScript includes global constants for the REST query and global variables that are used in several functions.</span></span> <span data-ttu-id="3e6f5-219">Кнопка **Get ProjectData Endpoint** вызывает функцию, которая инициализирует глобальные переменные и определяет, подключена ли Project к `setOdataUrl` Project Web App.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-219">The **Get ProjectData Endpoint** button calls the `setOdataUrl` function, which initializes the global variables and determines whether Project is connected with Project Web App.</span></span>

<span data-ttu-id="3e6f5-220">Остальная часть файла HelloProjectOData.js включает две функции: функция называется, когда пользователь выбирает Сравнение всех проектов; а функция вычисляет средние значения, а затем заполняет таблицу сравнения значениями, отформатированные для цвета и `retrieveOData`  `parseODataResult` единиц.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-220">The remainder of the HelloProjectOData.js file includes two functions: the `retrieveOData` function is called when the user selects **Compare All Projects**; and the `parseODataResult` function calculates averages and then populates the comparison table with values that are formatted for color and units.</span></span>

### <a name="procedure-5-to-create-the-javascript-code"></a><span data-ttu-id="3e6f5-p136">Процедура 5. Создание кода JavaScript</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p136">Procedure 5. To create the JavaScript code</span></span>

1. <span data-ttu-id="3e6f5-223">Удалите весь код в файле HelloProjectOData.js по умолчанию, а затем добавьте глобальные переменные `**` иOffice.iniфункцию tialize.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-223">Delete all code in the default HelloProjectOData.js file, and then add the global variables and `**`Office.initialize\` function.</span></span> <span data-ttu-id="3e6f5-224">Имена переменных, написанные полностью заглавными буквами подразумевают, что они являются константами; они позже будут использоваться с переменной **_pwa** для создания запроса REST в этом примере.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-224">Variable names that are all capitals imply that they are constants; they are later used with the **_pwa** variable to create the REST query in this example.</span></span>

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

2. <span data-ttu-id="3e6f5-225">Добавление `setOdataUrl` и связанные функции.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-225">Add `setOdataUrl` and related functions.</span></span> <span data-ttu-id="3e6f5-226">Функция `setOdataUrl` вызывает `getProjectGuid` и `getDocumentUrl` инициализирует глобальные переменные.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-226">The `setOdataUrl` function calls `getProjectGuid` and `getDocumentUrl` to initialize the global variables.</span></span> <span data-ttu-id="3e6f5-227">В [методе getProjectFieldAsync](/javascript/api/office/office.document)функция анонимного вызова параметра  _callback_ включает кнопку Сравнение всех проектов с помощью метода в библиотеке jQuery, а затем отображает `removeAttr` URL-адрес **службы ProjectData.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-227">In the [getProjectFieldAsync method](/javascript/api/office/office.document), the anonymous function for the  _callback_ parameter enables the **Compare All Projects** button by using the `removeAttr` method in the jQuery library, and then displays the URL of the **ProjectData** service.</span></span> <span data-ttu-id="3e6f5-228">Если Project не подключен к Project Web App, функция вызывает ошибку, которая отображает всплывающее сообщение об ошибке.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-228">If Project is not connected with Project Web App, the function throws an error, which displays a pop-up error message.</span></span> <span data-ttu-id="3e6f5-229">Файл SurfaceErrors.js включает `throwError` метод.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-229">The SurfaceErrors.js file includes the `throwError` method.</span></span>

   > [!NOTE]
   > <span data-ttu-id="3e6f5-230">Если вы работаете с Visual Studio на компьютере Project Server, то для того, чтобы использовать отладку по клавише **F5**, раскомментируйте код после строки, инициализирующей глобальную переменную **_pwa**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-230">If you run Visual Studio on the Project Server computer, to use **F5** debugging, uncomment the code after the line that initializes the **_pwa** global variable.</span></span> <span data-ttu-id="3e6f5-231">Чтобы включить метод jQuery при отладке на компьютере Project Server, необходимо установить значение для `ajax` `localhost` PWA URL-адреса. Если вы Visual Studio на удаленном компьютере, `localhost` URL-адрес не требуется.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-231">To enable using the jQuery `ajax` method when debugging on the Project Server computer, you must set the `localhost` value for the PWA URL.If you run Visual Studio on a remote computer, the  `localhost` URL is not required.</span></span> <span data-ttu-id="3e6f5-232">Перед развертыванием надстройки закомментируйте этот код.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-232">Before you deploy the add-in, comment out that code.</span></span>

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

3. <span data-ttu-id="3e6f5-233">Добавьте функцию, которая соединять значения для запроса REST, а затем вызывает функцию в jQuery, чтобы получить запрашиваемую информацию из `retrieveOData` `ajax` **службы ProjectData.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-233">Add the `retrieveOData` function, which concatenates values for the REST query and then calls the `ajax` function in jQuery to get the requested data from the **ProjectData** service.</span></span> <span data-ttu-id="3e6f5-234">Переменная **support.cors** позволяет совместное использование ресурсов для разных стран (CORS) с `ajax` функцией.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-234">The **support.cors** variable enables cross-origin resource sharing (CORS) with the `ajax` function.</span></span> <span data-ttu-id="3e6f5-235">Если утверждение **support.cors** отсутствует или настроено на ложное, функция возвращает ошибку `ajax` без **транспорта.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-235">If the **support.cors** statement is missing or is set to **false**, the `ajax` function returns a **No transport** error.</span></span>

   > [!NOTE]
   > <span data-ttu-id="3e6f5-p141">Приведенный ниже код подходит для локального сервера Project Server 2013. В Project в Интернете можно использовать OAuth для проверки подлинности на основе токенов. Дополнительные сведения см. в статье [Обход ограничений, связанных с принципом одинакового источника, в надстройках Office](../develop/addressing-same-origin-policy-limitations.md).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p141">The following code works with an on-premises installation of Project Server 2013. For Project on the web, you can use OAuth for token-based authentication. For more information, see [Addressing same-origin policy limitations in Office Add-ins](../develop/addressing-same-origin-policy-limitations.md).</span></span>

   <span data-ttu-id="3e6f5-239">В `ajax` вызове можно использовать либо параметр _headers,_ либо _параметр beforeSend._</span><span class="sxs-lookup"><span data-stu-id="3e6f5-239">In the `ajax` call, you can use either the _headers_ parameter or the _beforeSend_ parameter.</span></span> <span data-ttu-id="3e6f5-240">Полный _параметр_ — это анонимная функция, так что она находится в той же области, что и переменные `retrieveOData` в .</span><span class="sxs-lookup"><span data-stu-id="3e6f5-240">The _complete_ parameter is an anonymous function so that it is in the same scope as the variables in `retrieveOData`.</span></span> <span data-ttu-id="3e6f5-241">Функция полного  _отображения_ параметра приводит к контролю, а также вызывает метод для анализа и отображения `odataText` ответа `parseODataResult` JSON.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-241">The function for the  _complete_ parameter displays results in the `odataText` control and also calls the `parseODataResult` method to parse and display the JSON response.</span></span> <span data-ttu-id="3e6f5-242">Параметр _ошибки_ указывает именоваемую функцию, которая записывает сообщение об ошибке в управление, а также использует метод для отображения `getProjectDataErrorHandler` `odataText` `throwError` всплывающее сообщение.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-242">The _error_ parameter specifies the named `getProjectDataErrorHandler` function, which writes an error message to the `odataText` control and also uses the `throwError` method to display a pop-up message.</span></span>

    ```js
    // Functions to get and parse the Project Server reporting data./

    // Get data about all projects on Project Server,
    // by using a REST query with the ajax method in jQuery.
    function retrieveOData() {
        var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
        var accept = "application/json; odata=verbose";
        accept.toLocaleLowerCase();

        // Enable cross-origin scripting (required by jQuery 1.5 and later).
        // This does not work with Project on the web.
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

4. <span data-ttu-id="3e6f5-243">Добавьте `parseODataResult` метод, который дезерализует и обрабатывает ответ JSON из службы OData.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-243">Add the `parseODataResult` method, which deserializes and processes the JSON response from the OData service.</span></span> <span data-ttu-id="3e6f5-244">Метод вычисляет средние значения затрат и данных о работе с точностью одного или двух десятичных мест, форматы значения с правильным цветом и добавляет единицу `parseODataResult` **$** **(hrs** или), а затем отображает значения в указанных ячейках **%** таблицы.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-244">The `parseODataResult` method calculates average values of the cost and work data to an accuracy of one or two decimal places, formats values with the correct color and adds a unit ( **$**, **hrs**, or **%**), and then displays the values in specified table cells.</span></span>

   <span data-ttu-id="3e6f5-245">Если GUID активного проекта соответствует значению, переменная `ProjectId` `myProjectIndex` задана индексу проекта.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-245">If the GUID of the active project matches the `ProjectId` value, the `myProjectIndex` variable is set to the project index.</span></span> <span data-ttu-id="3e6f5-246">Если указывает, что активный проект публикуется на Project Server, форматы метода и отображает данные о затратах и работе `myProjectIndex` `parseODataResult` для этого проекта.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-246">If `myProjectIndex` indicates the active project is published on Project Server, the `parseODataResult` method formats and displays cost and work data for that project.</span></span> <span data-ttu-id="3e6f5-247">Если активный проект не опубликован, то для него отображается значение **NA** синим цветом.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-247">If the active project is not published, values for the active project are displayed as a blue **NA**.</span></span>

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

## <a name="testing-the-helloprojectodata-add-in"></a><span data-ttu-id="3e6f5-248">Тестирование надстройки HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="3e6f5-248">Testing the HelloProjectOData add-in</span></span>

<span data-ttu-id="3e6f5-249">Для проверки и отламки надстройки **HelloProjectOData** с Visual Studio 2015 Project профессиональный 2013 года необходимо установить на компьютере разработки.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-249">To test and debug the **HelloProjectOData** add-in with Visual Studio 2015, Project Professional 2013 must be installed on the development computer.</span></span> <span data-ttu-id="3e6f5-250">Для работы с различными тестовыми сценариями убедитесь, что можно выбрать открытие файлов Project на локальном компьютере или подключение к Project Web App.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-250">To enable different test scenarios, ensure that you can choose whether Project opens for files on the local computer or connects with Project Web App.</span></span> <span data-ttu-id="3e6f5-251">Например, сделайте следующие действия.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-251">For example, do the following steps.</span></span>

1. <span data-ttu-id="3e6f5-252">Во вкладке **ФАЙЛ** на ленте выберите вкладку **Сведения** в представлении Backstage, а затем выберите **Управление учетными записями**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-252">On the **FILE** tab on the ribbon, choose the **Info** tab in the Backstage view, and then choose **Manage Accounts**.</span></span>

2. <span data-ttu-id="3e6f5-253">В **диалоговом окне** Project учетных  записей веб-приложения список доступных учетных записей может иметь несколько Project Web App учетных записей в дополнение к локальной **учетной записи Компьютера.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-253">In the **Project web app Accounts** dialog box, the **Available accounts** list can have multiple Project Web App accounts in addition to the local **Computer** account.</span></span> <span data-ttu-id="3e6f5-254">В разделе **Во время запуска** выберите **Выбрать учетную запись**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-254">In the **When starting** section, select **Choose an account**.</span></span>

3. <span data-ttu-id="3e6f5-255">Закройте Project, чтобы среда Visual Studio могла запустить его для отладки надстройки.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-255">Close Project so that Visual Studio can start it for debugging the add-in.</span></span>

<span data-ttu-id="3e6f5-256">Базовые тесты должны быть следующие:</span><span class="sxs-lookup"><span data-stu-id="3e6f5-256">Basic tests should include the following:</span></span>

- <span data-ttu-id="3e6f5-257">Запустите приложение в Visual Studio и откройте опубликованный проект из Project Web App, содержащего данные о материальных и трудовых затратах.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-257">Run the add-in from Visual Studio, and then open a published project from Project Web App that contains cost and work data.</span></span> <span data-ttu-id="3e6f5-258">Убедитесь, что надстройка отображает конечную точку **ProjectData** и правильно отображает данные о затратах и работе в таблице.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-258">Verify that the add-in displays the **ProjectData** endpoint and correctly displays the cost and work data in the table.</span></span> <span data-ttu-id="3e6f5-259">Можно использовать выходные данные в элементе управления **odataText** для проверки запроса REST и других сведений.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-259">You can use the output in the **odataText** control to check the REST query and other information.</span></span>

- <span data-ttu-id="3e6f5-260">Запустите надстройку еще раз и выберите профиль локального компьютера с помощью диалогового окна **Вход** во время запуска Project.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-260">Run the add-in again, where you choose the local computer profile in the **Login** dialog box when Project starts.</span></span> <span data-ttu-id="3e6f5-261">Откройте локальный MPP-файл и протестируйте надстройку.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-261">Open a local .mpp file, and then test the add-in.</span></span> <span data-ttu-id="3e6f5-262">Убедитесь, что она отображает сообщение об ошибке при попытке получить конечную точку **ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-262">Verify that the add-in displays an error message when you try to get the **ProjectData** endpoint.</span></span>

- <span data-ttu-id="3e6f5-263">Запустите надстройку еще раз и создайте проект, содержащий задачи с данными о материальных и трудовых затратах.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-263">Run the add-in again, where you create a project that has tasks with cost and work data.</span></span> <span data-ttu-id="3e6f5-264">Этот проект можно сохранить в Project Web App, но не публиковать.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-264">You can save the project to Project Web App, but don't publish it.</span></span> <span data-ttu-id="3e6f5-265">Убедитесь, что надстройка отображает данные с Project Server, но показывает **NA** для текущего проекта.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-265">Verify that the add-in displays data from Project Server, but **NA** for the current project.</span></span>

### <a name="procedure-6-to-test-the-add-in"></a><span data-ttu-id="3e6f5-p150">Процедура 6. Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p150">Procedure 6. To test the add-in</span></span>

1. <span data-ttu-id="3e6f5-p151">Запустите Project профессиональный 2013, подключитесь к Project Web App и создайте тестовый проект. Назначьте задачи локальным ресурсам или ресурсам предприятия, настройте различные значения процента выполнения для некоторых задач и затем опубликуйте проект. Закройте Project, что позволит Visual Studio запустить Project для отладки надстройки.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p151">Run Project Professional 2013, connect with Project Web App, and then create a test project. Assign tasks to local resources or to enterprise resources, set various values of percent complete on some tasks, and then publish the project. Quit Project, which enables Visual Studio to start Project for debugging the add-in.</span></span>

2. <span data-ttu-id="3e6f5-271">В Visual Studio нажмите клавишу **F5**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-271">In Visual Studio, press **F5**.</span></span> <span data-ttu-id="3e6f5-272">Войдите в Project Web App и затем откройте проект, созданный на предыдущем шаге.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-272">Log on to Project Web App, and then open the project that you created in the previous step.</span></span> <span data-ttu-id="3e6f5-273">Проект можно открыть в режиме чтения или в режиме редактирования.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-273">You can open the project in read-only mode or in edit mode.</span></span>

3. <span data-ttu-id="3e6f5-274">На **вкладке PROJECT** ленты **в** списке Office надстройки выберите **Hello ProjectData** (см. рис. 5).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-274">On the **PROJECT** tab of the ribbon, in the **Office Add-ins** drop-down list, select **Hello ProjectData** (see Figure 5).</span></span> <span data-ttu-id="3e6f5-275">Кнопка **Сравнение всех проектов** должна быть отключена.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-275">The **Compare All Projects** button should be disabled.</span></span>

    <span data-ttu-id="3e6f5-276">*Рис. 5. Запуск надстройки HelloProjectOData*</span><span class="sxs-lookup"><span data-stu-id="3e6f5-276">*Figure 5. Starting the HelloProjectOData add-in*</span></span>

    ![Тестирование приложения HelloProjectOData.](../images/pj15-hello-project-data-test-the-app.png)

4. <span data-ttu-id="3e6f5-278">В области **задач Hello ProjectData** выберите конечную точку **Get ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-278">In the **Hello ProjectData** task pane, select **Get ProjectData Endpoint**.</span></span> <span data-ttu-id="3e6f5-279">В **строке projectDataEndPoint** должен быть указан URL-адрес  службы **ProjectData** и включена кнопка Сравнение всех проектов (см. рис. 6).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-279">The **projectDataEndPoint** line should show the URL of the **ProjectData** service, and the **Compare All Projects** button should be enabled (see Figure 6).</span></span>

5. <span data-ttu-id="3e6f5-280">Нажмите кнопку **Compare All Projects**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-280">Select **Compare All Projects**.</span></span> <span data-ttu-id="3e6f5-281">Надстройка может приостановить работу на время получения данных из службы **ProjectData**, а затем она должна отобразить отформатированные средние и текущие значения в таблице.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-281">The add-in may pause while it retrieves data from the **ProjectData** service, and then it should display the formatted average and current values in the table.</span></span>

    <span data-ttu-id="3e6f5-282">*Рис. 6. Просмотр результатов запроса REST*</span><span class="sxs-lookup"><span data-stu-id="3e6f5-282">*Figure 6. Viewing results of the REST query*</span></span>

    ![Просмотр результатов запроса REST.](../images/pj15-hello-project-data-rest-results.png)

6. <span data-ttu-id="3e6f5-284">Проверьте выходные данные в текстовом поле.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-284">Examine output in the text box.</span></span> <span data-ttu-id="3e6f5-285">Они должны показывать путь к документу, запрос REST, сведения о состоянии и результаты JSON от вызовов **ajax** и **parseODataResult**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-285">It should show the document path, REST query, status information, and JSON results from the calls to **ajax** and **parseODataResult**.</span></span> <span data-ttu-id="3e6f5-286">Вывод помогает понять, создать и отлагировать код в `parseODataResult` таком методе, как `projCost += Number(res.d.results[i].ProjectCost);` .</span><span class="sxs-lookup"><span data-stu-id="3e6f5-286">The output helps to understand, create, and debug code in the `parseODataResult` method such as `projCost += Number(res.d.results[i].ProjectCost);`.</span></span>

    <span data-ttu-id="3e6f5-287">Ниже приводится пример вывода с разрывами строк и пробелами, добавленными в текст для ясности, для трех проектов в Project Web App экземпляре.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-287">Following is an example of the output with line breaks and spaces added to the text for clarity, for three projects in a Project Web App instance.</span></span>

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

7. <span data-ttu-id="3e6f5-288">Остановите отладку **(нажмите кнопку Shift + F5),** а затем нажмите **кнопку F5** снова, чтобы запустить новый экземпляр Project.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-288">Stop debugging (press **Shift + F5**), and then press **F5** again to run a new instance of Project.</span></span> <span data-ttu-id="3e6f5-289">В **диалоговом окне Login** выберите локальный профиль **компьютера,** а не Project Web App.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-289">In the **Login** dialog box, choose the local **Computer** profile, not Project Web App.</span></span> <span data-ttu-id="3e6f5-290">Создайте или откройте локальный файл project .mpp, откройте области **задач Hello ProjectData** и выберите конечную точку **Get ProjectData.**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-290">Create or open a local project .mpp file, open the **Hello ProjectData** task pane, and then select **Get ProjectData Endpoint**.</span></span> <span data-ttu-id="3e6f5-291">Надстройка должна показывать отсутствие **подключения!**</span><span class="sxs-lookup"><span data-stu-id="3e6f5-291">The add-in should show a **No connection!**</span></span> <span data-ttu-id="3e6f5-292">ошибка (см. рис. 7) и кнопка **Сравнение всех** проектов должны оставаться отключенными.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-292">error (see Figure 7), and the **Compare All Projects** button should remain disabled.</span></span>

   <span data-ttu-id="3e6f5-293">*Рис. 7. Использование надстройки без подключения Project Web App*</span><span class="sxs-lookup"><span data-stu-id="3e6f5-293">*Figure 7. Using the add-in without a Project web app connection*</span></span>

   ![Использование приложения без Project Web App подключения.](../images/pj15-hello-project-data-no-connection.png)

8. <span data-ttu-id="3e6f5-295">Остановите отладку и нажмите клавишу **F5** снова.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-295">Stop debugging, and then press **F5** again.</span></span> <span data-ttu-id="3e6f5-296">Войдите в Project Web App и создайте проект, содержащий данные о материальных и трудовых затратах.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-296">Log on to Project Web App, and then create a project that contains cost and work data.</span></span> <span data-ttu-id="3e6f5-297">Проект можно сохранить, но не публикуйте его.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-297">You can save the project, but don't publish it.</span></span>

   <span data-ttu-id="3e6f5-298">В области **задач Hello ProjectData** при выборе сравнения всех проектов следует увидеть синий **NA** для полей в столбце **Current** (см. рис. 8).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-298">In the **Hello ProjectData** task pane, when you select **Compare All Projects**, you should see a blue **NA** for fields in the **Current** column (see Figure 8).</span></span>

   <span data-ttu-id="3e6f5-299">*Рис. 8. Сравнение неопубликованного проекта с другими проектами*</span><span class="sxs-lookup"><span data-stu-id="3e6f5-299">*Figure 8. Comparing an unpublished project with other projects*</span></span>

   ![Сравнение неопубликованного проекта с другими.](../images/pj15-hello-project-data-not-published.png)

<span data-ttu-id="3e6f5-p159">Даже если ваша надстройка работала правильно в предыдущих тестах, есть другие тесты, которые необходимо выполнить. Например:</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p159">Even if your add-in is working correctly in the previous tests, there are other tests that should be run. For example:</span></span>

- <span data-ttu-id="3e6f5-303">Откройте в Project Web App проект, который не содержит данных о материальных и трудовых затратах для задач.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-303">Open a project from Project Web App that has no cost or work data for the tasks.</span></span> <span data-ttu-id="3e6f5-304">В полях столбца **Current (текущий)** должны отображаться нули.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-304">You should see values of zero in the fields in the **Current** column.</span></span>

- <span data-ttu-id="3e6f5-305">Протестируйте проект, не содержащий задачи.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-305">Test a project that has no tasks.</span></span>

- <span data-ttu-id="3e6f5-p161">Если вы измените надстройку и опубликуете ее, необходимо запустить аналогичные тесты снова с опубликованной надстройкой. Другие вопросы см. в разделе [Дальнейшие действия](#next-steps).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p161">If you modify the add-in and publish it, you should run similar tests again with the published add-in. For other considerations, see [Next steps](#next-steps).</span></span>

> [!NOTE]
> <span data-ttu-id="3e6f5-308">Имеются ограничения на объем данных, который может быть возвращен в одном запросе службы **ProjectData**; этот объем данных меняется для разных сущностей.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-308">There are limits to the amount of data that can be returned in one query of the **ProjectData** service; the amount of data varies by entity.</span></span> <span data-ttu-id="3e6f5-309">Например, набор объектов имеет ограничение по умолчанию в 100 проектов на запрос, но набор сущности имеет ограничение по `Projects` `Risks` умолчанию в 200.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-309">For example, the `Projects` entity set has a default limit of 100 projects per query, but the `Risks` entity set has a default limit of 200.</span></span> <span data-ttu-id="3e6f5-310">Для установки в рабочей среде код в примере **HelloProjectOData** необходимо изменить для поддержки запросов, содержащих более 100 проектов.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-310">For a production installation, the code in the **HelloProjectOData** example should be modified to enable queries of more than 100 projects.</span></span> <span data-ttu-id="3e6f5-311">Дополнительные сведения [см.](#next-steps) в следующих действиях и [запросе каналов OData для Project отчетов.](/previous-versions/office/project-odata/jj163048(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="3e6f5-311">For more information, see [Next steps](#next-steps) and [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

## <a name="example-code-for-the-helloprojectodata-add-in"></a><span data-ttu-id="3e6f5-312">Пример кода для надстройки HelloProjectOData</span><span class="sxs-lookup"><span data-stu-id="3e6f5-312">Example code for the HelloProjectOData add-in</span></span>

### <a name="helloprojectodatahtml-file"></a><span data-ttu-id="3e6f5-313">Файл HelloProjectOData.html</span><span class="sxs-lookup"><span data-stu-id="3e6f5-313">HelloProjectOData.html file</span></span>

<span data-ttu-id="3e6f5-314">Приведенный ниже код находится в файле `Pages\HelloProjectOData.html` проекта **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-314">The following code is in the `Pages\HelloProjectOData.html` file of the **HelloProjectODataWeb** project.</span></span>

```HTML
<!DOCTYPE html>
<html>
    <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Test ProjectData Service</title>

        <link rel="stylesheet" type="text/css" href="../Content/Office.css" />

        <!-- Add your CSS styles to the following file. -->
        <link rel="stylesheet" type="text/css" href="../Content/App.css" />

        <!-- Use the CDN reference to the mini-version of jQuery when deploying your add-in. -->
        <!--<script src="http://ajax.aspnetcdn.com/ajax/jquery/jquery-1.9.0.min.js"></script> -->
        <script src="../Scripts/jquery-1.7.1.js"></script>

        <!-- Use the CDN reference to Office.js when deploying your add-in -->
        <!--<script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>-->

        <!-- Use the local script references for Office.js to enable offline debugging -->
        <script src="../Scripts/Office/1.0/MicrosoftAjax.js"></script>
        <script src="../Scripts/Office/1.0/Office.js"></script>

        <!-- Add your JavaScript to the following files. -->
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

### <a name="helloprojectodatajs-file"></a><span data-ttu-id="3e6f5-315">Файл HelloProjectOData.js</span><span class="sxs-lookup"><span data-stu-id="3e6f5-315">HelloProjectOData.js file</span></span>

<span data-ttu-id="3e6f5-316">Приведенный ниже код находится в файле `Scripts\Office\HelloProjectOData.js` проекта **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-316">The following code is in the `Scripts\Office\HelloProjectOData.js` file of the **HelloProjectODataWeb** project.</span></span>

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

//  Functions to get and parse the Project Server reporting data./

// Get data about all projects on Project Server,
// by using a REST query with the ajax method in jQuery.
function retrieveOData() {
    var restUrl = _odataUrl + PROJQUERY + QUERY_FILTER + QUERY_SELECT1 + QUERY_SELECT2;
    var accept = "application/json; odata=verbose";
    accept.toLocaleLowerCase();

    // Enable cross-origin scripting (required by jQuery 1.5 and later).
    // This does not work with Project on the web.
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

### <a name="appcss-file"></a><span data-ttu-id="3e6f5-317">Файл App.css</span><span class="sxs-lookup"><span data-stu-id="3e6f5-317">App.css file</span></span>

<span data-ttu-id="3e6f5-318">Приведенный ниже код находится в файле `Content\App.css` проекта **HelloProjectODataWeb**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-318">The following code is in the `Content\App.css` file of the **HelloProjectODataWeb** project.</span></span>

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

### <a name="surfaceerrorsjs-file"></a><span data-ttu-id="3e6f5-319">Файл SurfaceErrors.js</span><span class="sxs-lookup"><span data-stu-id="3e6f5-319">SurfaceErrors.js file</span></span>

<span data-ttu-id="3e6f5-320">Вы можете скопировать код для файла SurfaceErrors.js из раздела _Надежное программирование_ статьи [Создание первой надстройки области задач для Project 2013 с помощью текстового редактора](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-320">You can copy code for the SurfaceErrors.js file from the _Robust Programming_ section of [Create your first task pane add-in for Project 2013 by using a text editor](../project/create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md).</span></span>

## <a name="next-steps"></a><span data-ttu-id="3e6f5-321">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="3e6f5-321">Next steps</span></span>

<span data-ttu-id="3e6f5-322">Если **HelloProjectOData** — это производственная надстройка, которая будет продаваться в AppSource или распространяться в каталоге SharePoint приложения, она будет разработана по-другому.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-322">If **HelloProjectOData** were a production add-in to be sold in AppSource or distributed in a SharePoint app catalog, it would be designed differently.</span></span> <span data-ttu-id="3e6f5-323">Например, здесь не было бы отладочных выходных данных в текстовом поле и, вероятно, не было бы кнопки для получения конечной точки **ProjectData**.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-323">For example, there would be no debug output in a text box, and probably no button to get the **ProjectData** endpoint.</span></span> <span data-ttu-id="3e6f5-324">Кроме того, необходимо переписать функцию для обработки Project Web App экземпляров с более `retireveOData` чем 100 проектами.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-324">You would also have to rewrite the `retireveOData` function to handle Project Web App instances that have more than 100 projects.</span></span>

<span data-ttu-id="3e6f5-p164">Надстройка должна содержать дополнительные проверки ошибок, а также логику для записи, объяснения или демонстрации пограничных случаев. Например, если экземпляр Project Web App содержит 1000 проектов со средней продолжительностью в пять дней и средними затратами в $2400, а активный проект является единственным с продолжительностью более 20 дней, то сравнение материальных и трудовых затрат может быть перекошено. Это может быть показано с помощью частотной диаграммы. Вам необходимо добавить команды для отображения продолжительности, сравнения проектов с одинаковой продолжительностью или сравнения проектов из одного или разных отделов. Либо добавить возможность пользователю выбирать из списка полей, которые требуется отобразить.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p164">The add-in should contain additional error checks, plus logic to catch and explain or show edge cases. For example, if a Project Web App instance has 1000 projects with an average duration of five days and average cost of $2400, and the active project is the only one that has a duration longer than 20 days, the cost and work comparison would be skewed. That could be shown with a frequency graph. You could add options to display duration, compare similar length projects, or compare projects from the same or different departments. Or, add a way for the user to select from a list of fields to display.</span></span>

<span data-ttu-id="3e6f5-330">Для других запросов службы **ProjectData** имеются ограничения на длину строки запроса, что влияет на число шагов, которые запрос может предпринять для выборки из родительской коллекции в объект в дочерней коллекции.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-330">For other queries of the **ProjectData** service, there are limits to the length of the query string, which affects the number of steps that a query can take from a parent collection to an object in a child collection.</span></span> <span data-ttu-id="3e6f5-331">Например, двухшаговый запрос **Projects** в **Tasks** для получения элементов задач работает, но трехшаговый запрос, такой как **Projects** в **Tasks** в **Assignments**, для получения элемента назначения может превысить максимальную длину URL-адреса по умолчанию.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-331">For example, a two-step query of **Projects** to **Tasks** to task item works, but a three-step query such as **Projects** to **Tasks** to **Assignments** to assignment item may exceed the default maximum URL length.</span></span> <span data-ttu-id="3e6f5-332">Дополнительные сведения см. в [веб-каналах запроса OData для Project отчетов.](/previous-versions/office/project-odata/jj163048(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="3e6f5-332">For more information, see [Querying OData feeds for Project reporting data](/previous-versions/office/project-odata/jj163048(v=office.15)).</span></span>

<span data-ttu-id="3e6f5-333">Если вы измените **надстройки HelloProjectOData** для производственного использования, сделайте следующие действия.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-333">If you modify the **HelloProjectOData** add-in for production use, do the following steps.</span></span>

- <span data-ttu-id="3e6f5-334">В файле HelloProjectOData.html для лучшей производительности измените ссылку office.js из локального проекта на ссылку CDN:</span><span class="sxs-lookup"><span data-stu-id="3e6f5-334">In the HelloProjectOData.html file, for better performance, change the office.js reference from the local project to the CDN reference:</span></span>

    ```HTML
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

- <span data-ttu-id="3e6f5-335">Переписать `retrieveOData` функцию, чтобы включить запросы более 100 проектов.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-335">Rewrite the `retrieveOData` function to enable queries of more than 100 projects.</span></span> <span data-ttu-id="3e6f5-336">Например, можно получить число проектов с помощью запроса `~/ProjectData/Projects()/$count` и использовать оператор _$skip_ и оператор _$top_ в запросе REST для получения данных проекта.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-336">For example, you could get the number of projects with a `~/ProjectData/Projects()/$count` query, and use the _$skip_ operator and _$top_ operator in the REST query for project data.</span></span> <span data-ttu-id="3e6f5-337">Запустите несколько запросов в цикле и затем усредните данные из всех запросов.</span><span class="sxs-lookup"><span data-stu-id="3e6f5-337">Run multiple queries in a loop, and then average the data from each query.</span></span> <span data-ttu-id="3e6f5-338">Каждый запрос для данных проекта будет иметь форму:</span><span class="sxs-lookup"><span data-stu-id="3e6f5-338">Each query for project data would be of the form:</span></span> 

  `~/ProjectData/Projects()?skip= [numSkipped]&amp;$top=100&amp;$filter=[filter]&amp;$select=[field1,field2, ???????]`

  <span data-ttu-id="3e6f5-p167">For more information, see [OData System Query Options Using the REST Endpoint](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps&preserve-view=true) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-p167">For more information, see [OData System Query Options Using the REST Endpoint](/previous-versions/dynamicscrm-2015/developers-guide/gg309461(v=crm.7)). You can also use the [Set-SPProjectOdataConfiguration](/powershell/module/sharepoint-server/Set-SPProjectOdataConfiguration?view=sharepoint-ps&preserve-view=true) command in Windows PowerShell to override the default page size for a query of the **Projects** entity set (or any of the 33 entity sets). See [ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15)).</span></span>

- <span data-ttu-id="3e6f5-342">Сведения о развертывании надстройки см. в статье [Публикация надстройки Office](../publish/publish.md).</span><span class="sxs-lookup"><span data-stu-id="3e6f5-342">To deploy the add-in, see [Publish your Office Add-in](../publish/publish.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3e6f5-343">См. также</span><span class="sxs-lookup"><span data-stu-id="3e6f5-343">See also</span></span>

- [<span data-ttu-id="3e6f5-344">Надстройки области задач для Project</span><span class="sxs-lookup"><span data-stu-id="3e6f5-344">Task pane add-ins for Project</span></span>](project-add-ins.md)
- [<span data-ttu-id="3e6f5-345">Создание первой надстройки области задач для Project 2013 с помощью текстового редактора</span><span class="sxs-lookup"><span data-stu-id="3e6f5-345">Create your first task pane add-in for Project 2013 by using a text editor</span></span>](create-your-first-task-pane-add-in-for-project-by-using-a-text-editor.md)
- <span data-ttu-id="3e6f5-346">[ProjectData — Справочник по службе Project OData](/previous-versions/office/project-odata/jj163015(v=office.15))</span><span class="sxs-lookup"><span data-stu-id="3e6f5-346">[ProjectData - Project OData service reference](/previous-versions/office/project-odata/jj163015(v=office.15))</span></span>
- [<span data-ttu-id="3e6f5-347">XML-манифест надстройки Office</span><span class="sxs-lookup"><span data-stu-id="3e6f5-347">Office Add-ins XML manifest</span></span>](../develop/add-in-manifests.md)
- [<span data-ttu-id="3e6f5-348">Публикация надстройки Office</span><span class="sxs-lookup"><span data-stu-id="3e6f5-348">Publish your Office Add-in</span></span>](../publish/publish.md)
