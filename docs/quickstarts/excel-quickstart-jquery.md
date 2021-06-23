---
title: Создание первой надстройки области задач Excel
description: Узнайте, как создать простую надстройку области задач Excel, используя API JS для Office.
ms.date: 1/19/2021
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 09abf03c5e345c61a4e98226930d79120c95949b
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076639"
---
# <a name="build-an-excel-task-pane-add-in"></a><span data-ttu-id="ffd2f-103">Создание надстройки области задач Excel</span><span class="sxs-lookup"><span data-stu-id="ffd2f-103">Build an Excel task pane add-in</span></span>

<span data-ttu-id="ffd2f-104">В этой статье вы ознакомитесь с процессом создания надстройки области задач Excel.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-104">In this article, you'll walk through the process of building an Excel task pane add-in.</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="ffd2f-105">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="ffd2f-105">Create the add-in</span></span>

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]
# <a name="yeoman-generator"></a>[<span data-ttu-id="ffd2f-106">Генератор Yeoman</span><span class="sxs-lookup"><span data-stu-id="ffd2f-106">Yeoman generator</span></span>](#tab/yeomangenerator)

[!include[Redirect to the single sign-on (SSO) quick start](../includes/sso-quickstart-reference.md)]

## <a name="prerequisites"></a><span data-ttu-id="ffd2f-107">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="ffd2f-107">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="ffd2f-108">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="ffd2f-108">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="ffd2f-109">**Выберите тип проекта:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="ffd2f-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="ffd2f-110">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="ffd2f-110">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="ffd2f-111">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="ffd2f-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="ffd2f-112">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="ffd2f-112">**Which Office client application would you like to support?**</span></span> `Excel`

![Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office.](../images/yo-office-excel.png)

<span data-ttu-id="ffd2f-114">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-114">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="explore-the-project"></a><span data-ttu-id="ffd2f-115">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="ffd2f-115">Explore the project</span></span>

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a><span data-ttu-id="ffd2f-116">Проверка</span><span class="sxs-lookup"><span data-stu-id="ffd2f-116">Try it out</span></span>

1. <span data-ttu-id="ffd2f-117">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-117">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)]

3. <span data-ttu-id="ffd2f-118">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-118">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: меню "Главная" в Excel с выделенной кнопкой "Показать область задач".](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="ffd2f-120">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-120">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="ffd2f-121">Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-121">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Снимок экрана: Excel с открытой областью задач надстройки и выделенной кнопкой "Запустить".](../images/excel-quickstart-addin-3c.png)

### <a name="next-steps"></a><span data-ttu-id="ffd2f-123">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="ffd2f-123">Next steps</span></span>

<span data-ttu-id="ffd2f-p101">Поздравляем, вы успешно создали надстройку панели задач Excel! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь [руководством по надстройкам Excel](../tutorials/excel-tutorial.md).</span><span class="sxs-lookup"><span data-stu-id="ffd2f-p101">Congratulations, you've successfully created an Excel task pane add-in! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the [Excel add-in tutorial](../tutorials/excel-tutorial.md).</span></span>

# <a name="visual-studio"></a>[<span data-ttu-id="ffd2f-126">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="ffd2f-126">Visual Studio</span></span>](#tab/visualstudio)

### <a name="prerequisites"></a><span data-ttu-id="ffd2f-127">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="ffd2f-127">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="ffd2f-128">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="ffd2f-128">Create the add-in project</span></span>

1. <span data-ttu-id="ffd2f-129">В Visual Studio выберите пункт **Создать проект**.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-129">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="ffd2f-130">Используя поле поиска, введите **надстройка**.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-130">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="ffd2f-131">Выберите вариант **Веб-надстройка Excel** и нажмите кнопку **Далее**.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-131">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="ffd2f-132">Присвойте проекту имя **ExcelWebAddIn1** и выберите **Создать**.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-132">Name your project **ExcelWebAddIn1** and select **Create**.</span></span>

4. <span data-ttu-id="ffd2f-133">В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в Excel**, а затем нажмите кнопку **Готово**, чтобы создать проект.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-133">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="ffd2f-p103">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-p103">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="ffd2f-136">Обзор решения Visual Studio</span><span class="sxs-lookup"><span data-stu-id="ffd2f-136">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="ffd2f-137">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="ffd2f-137">Update the code</span></span>

1. <span data-ttu-id="ffd2f-p104">Файл **Home.html** содержит HTML-контент, который будет отображаться в области задач надстройки. В файле **Home.html** замените элемент `<body>` на приведенную ниже часть кода и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-p104">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

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

2. <span data-ttu-id="ffd2f-p105">Откройте файл **Home.js** в корневой папке проекта веб-приложения. Этот файл содержит скрипт надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-p105">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

3. <span data-ttu-id="ffd2f-p106">Откройте файл **Home.css** в корневой папке проекта веб-приложения. Этот файл определяет специальные стили надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-p106">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="ffd2f-146">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="ffd2f-146">Update the manifest</span></span>

1. <span data-ttu-id="ffd2f-147">Откройте **Обозреватель решений**, перейдите к проекту надстройки **ExcelWebAddIn1**, затем откройте каталог **ExcelWebAddIn1Manifest**.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-147">In **Solution Explorer**, go to the **ExcelWebAddIn1** add-in project and open the **ExcelWebAddIn1Manifest** directory.</span></span> <span data-ttu-id="ffd2f-148">Этот каталог содержит **ExcelWebAddIn1.xml** (ваш файл манифеста).</span><span class="sxs-lookup"><span data-stu-id="ffd2f-148">This directory contains your manifest file, **ExcelWebAddIn1.xml**.</span></span> <span data-ttu-id="ffd2f-149">XML-файл манифеста определяет параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-149">The XML manifest file defines the add-in's settings and capabilities.</span></span> <span data-ttu-id="ffd2f-150">Дополнительные сведения о двух проектах, созданных решением Visual Studio, приведены ранее в разделе [Обзор решения Visual Studio](#explore-the-visual-studio-solution).</span><span class="sxs-lookup"><span data-stu-id="ffd2f-150">See the preceding section [Explore the Visual Studio solution](#explore-the-visual-studio-solution) for more information about the two projects created by your Visual Studio solution.</span></span> 

2. <span data-ttu-id="ffd2f-p108">Элемент `ProviderName` содержит заполнитель. Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-p108">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="ffd2f-p109">Атрибут `DefaultValue` элемента `DisplayName` содержит заполнитель. Замените его на строку **Моя надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-p109">The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="ffd2f-p110">Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его строкой **Надстройка области задач для Excel**.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-p110">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="ffd2f-157">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-157">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="ffd2f-158">Проверка</span><span class="sxs-lookup"><span data-stu-id="ffd2f-158">Try it out</span></span>

1. <span data-ttu-id="ffd2f-159">Протестируйте новую надстройку Excel в Visual Studio, нажав клавишу **F5** или кнопку **Запустить**, чтобы запустить Excel с кнопкой надстройки **Показать область задач** на ленте.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-159">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="ffd2f-160">Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-160">The add-in will be hosted locally on IIS.</span></span> <span data-ttu-id="ffd2f-161">Если вам будет предложено доверять сертификату, согласитесь, чтобы разрешить надстройке подключиться к приложению Office.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-161">If you are asked to trust a certificate, do so to allow the add-in to connect to its Office application.</span></span>

2. <span data-ttu-id="ffd2f-162">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-162">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: меню "Главная" в Excel с выделенной кнопкой "Показать область задач".](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="ffd2f-164">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-164">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="ffd2f-165">В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="ffd2f-165">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Снимок экрана: Excel с открытой областью задач надстройки.](../images/excel-quickstart-addin-2c.png)

[!include[Console tool note](../includes/console-tool-note.md)]

### <a name="next-steps"></a><span data-ttu-id="ffd2f-167">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="ffd2f-167">Next steps</span></span>

<span data-ttu-id="ffd2f-168">Поздравляем! Вы успешно создали надстройку области задач Excel!</span><span class="sxs-lookup"><span data-stu-id="ffd2f-168">Congratulations, you've successfully created an Excel task pane add-in!</span></span> <span data-ttu-id="ffd2f-169">Теперь изучите дополнительные сведения о [разработке надстроек Office с помощью Visual Studio](../develop/develop-add-ins-visual-studio.md).</span><span class="sxs-lookup"><span data-stu-id="ffd2f-169">Next, learn more about [developing Office Add-ins with Visual Studio](../develop/develop-add-ins-visual-studio.md).</span></span>

---

## <a name="see-also"></a><span data-ttu-id="ffd2f-170">См. также</span><span class="sxs-lookup"><span data-stu-id="ffd2f-170">See also</span></span>

* [<span data-ttu-id="ffd2f-171">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="ffd2f-171">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="ffd2f-172">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="ffd2f-172">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="ffd2f-173">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="ffd2f-173">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="ffd2f-174">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="ffd2f-174">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="ffd2f-175">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="ffd2f-175">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
