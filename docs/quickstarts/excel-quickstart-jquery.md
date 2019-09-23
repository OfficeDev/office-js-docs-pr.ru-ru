---
title: Создание первой надстройки области задач Excel
description: Узнайте, как создать простую надстройку области задач Excel, используя API JS для Office.
ms.date: 09/18/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6f5b78e1ffb154eb014bb4bb0ef8cb7135b2012f
ms.sourcegitcommit: a0257feabcfe665061c14b8bdb70cf82f7aca414
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/18/2019
ms.locfileid: "37035324"
---
# <a name="build-an-excel-task-pane-add-in"></a><span data-ttu-id="74e85-103">Создание надстройки области задач Excel</span><span class="sxs-lookup"><span data-stu-id="74e85-103">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="74e85-104">В этой статье вы ознакомитесь с процессом создания надстройки области задач Excel.</span><span class="sxs-lookup"><span data-stu-id="74e85-104">In this article, you'll walk through the process of building an Outlook task pane add-in.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="74e85-105">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="74e85-105">Create the add-in</span></span>

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="yeoman-generatortabyeomangenerator"></a>[<span data-ttu-id="74e85-106">Генератор Yeoman</span><span class="sxs-lookup"><span data-stu-id="74e85-106">Yeoman generator</span></span>](#tab/yeomangenerator)

### <a name="prerequisites"></a><span data-ttu-id="74e85-107">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="74e85-107">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="74e85-108">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="74e85-108">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="74e85-109">**Выберите тип проекта:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="74e85-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="74e85-110">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="74e85-110">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="74e85-111">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="74e85-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="74e85-112">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="74e85-112">**Which Office client application would you like to support?**</span></span> `Excel`

![Генератор Yeoman](../images/yo-office-excel.png)

<span data-ttu-id="74e85-114">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="74e85-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="explore-the-project"></a><span data-ttu-id="74e85-115">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="74e85-115">Explore the project</span></span>

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a><span data-ttu-id="74e85-116">Проверка</span><span class="sxs-lookup"><span data-stu-id="74e85-116">Try it out</span></span>

1. <span data-ttu-id="74e85-117">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="74e85-117">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="74e85-118">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="74e85-118">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="74e85-120">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="74e85-120">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="74e85-121">Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.</span><span class="sxs-lookup"><span data-stu-id="74e85-121">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-3c.png)

# <a name="visual-studiotabvisualstudio"></a>[<span data-ttu-id="74e85-123">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="74e85-123">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="74e85-124">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="74e85-124">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="74e85-125">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="74e85-125">Create the add-in project</span></span>

1. <span data-ttu-id="74e85-126">В строке меню Visual Studio выберите **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="74e85-126">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="74e85-127">В списке типов проекта разверните узел **Visual C#** или **Visual Basic**, разверните **Office/SharePoint**, затем выберите **Надстройки** > **Веб-надстройка Excel**.</span><span class="sxs-lookup"><span data-stu-id="74e85-127">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="74e85-128">Укажите имя проекта и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="74e85-128">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="74e85-129">В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в Excel**, а затем нажмите кнопку **Готово**, чтобы создать проект.</span><span class="sxs-lookup"><span data-stu-id="74e85-129">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="74e85-p101">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="74e85-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="74e85-132">Обзор решения Visual Studio</span><span class="sxs-lookup"><span data-stu-id="74e85-132">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="74e85-133">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="74e85-133">Update the code</span></span>

1. <span data-ttu-id="74e85-p102">Файл **Home.html** содержит HTML-контент, который будет отображаться в области задач надстройки. В файле **Home.html** замените элемент `<body>` на приведенную ниже часть кода и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="74e85-p102">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

    ```html
    <body class="ms-font-m ms-welcome">
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
                <p>Choose the button below to set the color of the selected range to green.</p>
                <br />
                <h3>Try it out</h3>
                <button class="ms-Button" id="set-color">Set color</button>
            </div>
        </div>
    </body>
    ```

2. <span data-ttu-id="74e85-p103">Откройте файл **Home.js** в корневой папке проекта веб-приложения. Этот файл содержит скрипт надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="74e85-p103">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

    ```js
    'use strict';

    (function () {

        Office.onReady(function() {
            // Office is ready
            $(document).ready(function () {
                // The document is ready
                $('#set-color').click(setColor);
            });
        });

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. <span data-ttu-id="74e85-p104">Откройте файл **Home.css** в корневой папке проекта веб-приложения. Этот файл определяет специальные стили надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="74e85-p104">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a><span data-ttu-id="74e85-142">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="74e85-142">Update the manifest</span></span>

1. <span data-ttu-id="74e85-143">Откройте XML-файл манифеста в проекте надстройки.</span><span class="sxs-lookup"><span data-stu-id="74e85-143">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="74e85-144">Этот файл определяет параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="74e85-144">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="74e85-p106">Элемент `ProviderName` содержит заполнитель. Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="74e85-p106">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="74e85-p107">Атрибут `DefaultValue` элемента `DisplayName` содержит заполнитель. Замените его на строку **Моя надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="74e85-p107">The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="74e85-p108">Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его строкой **Надстройка области задач для Excel**.</span><span class="sxs-lookup"><span data-stu-id="74e85-p108">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="74e85-151">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="74e85-151">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="74e85-152">Проверка</span><span class="sxs-lookup"><span data-stu-id="74e85-152">Try it out</span></span>

1. <span data-ttu-id="74e85-p109">Протестируйте новую надстройку Excel в Visual Studio, нажав клавишу **F5** или кнопку **Запустить**, чтобы запустить Excel с кнопкой надстройки **Показать область задач** на ленте. Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="74e85-p109">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="74e85-155">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="74e85-155">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="74e85-157">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="74e85-157">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="74e85-158">В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="74e85-158">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="74e85-160">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="74e85-160">Next steps</span></span>

<span data-ttu-id="74e85-161">Поздравляем! Вы успешно создали надстройку области задач Excel!</span><span class="sxs-lookup"><span data-stu-id="74e85-161">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="74e85-162">Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="74e85-162">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="74e85-163">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="74e85-163">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="74e85-164">См. также</span><span class="sxs-lookup"><span data-stu-id="74e85-164">See also</span></span>

* [<span data-ttu-id="74e85-165">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="74e85-165">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="74e85-166">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="74e85-166">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="74e85-167">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="74e85-167">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="74e85-168">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="74e85-168">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
