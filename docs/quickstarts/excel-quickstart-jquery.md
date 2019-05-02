---
title: Создание первой надстройки Excel с помощью jQuery
description: ''
ms.date: 03/19/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: d1783d257b0d46c42f16cf5e1f261407a355a63d
ms.sourcegitcommit: 68872372d181cca5bee37ade73c2250c4a56bab6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/01/2019
ms.locfileid: "33517081"
---
# <a name="build-an-excel-add-in-using-jquery"></a><span data-ttu-id="66cf6-102">Создание надстройки Excel с помощью jQuery</span><span class="sxs-lookup"><span data-stu-id="66cf6-102">Build an Excel add-in using jQuery</span></span>

<span data-ttu-id="66cf6-103">В этой статье мы разберем, как создать надстройку Excel, используя jQuery и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="66cf6-103">In this article, you'll walk through the process of building an Excel add-in by using jQuery and the Excel JavaScript API.</span></span> 

## <a name="create-the-add-in"></a><span data-ttu-id="66cf6-104">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="66cf6-104">Create the add-in</span></span>

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[<span data-ttu-id="66cf6-105">Visual Studio</span><span class="sxs-lookup"><span data-stu-id="66cf6-105">Visual Studio</span></span>](#tab/visual-studio)

### <a name="prerequisites"></a><span data-ttu-id="66cf6-106">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="66cf6-106">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a><span data-ttu-id="66cf6-107">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="66cf6-107">Create the add-in project</span></span>

1. <span data-ttu-id="66cf6-108">В строке меню Visual Studio выберите **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="66cf6-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="66cf6-109">В списке типов проекта разверните узел **Visual C#** или **Visual Basic**, разверните **Office/SharePoint**, затем выберите **Надстройки** > **Веб-надстройка Excel**.</span><span class="sxs-lookup"><span data-stu-id="66cf6-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="66cf6-110">Укажите имя проекта и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="66cf6-110">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="66cf6-111">В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в Excel**, а затем нажмите кнопку **Готово**, чтобы создать проект.</span><span class="sxs-lookup"><span data-stu-id="66cf6-111">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="66cf6-p101">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p101">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="66cf6-114">Обзор решения Visual Studio</span><span class="sxs-lookup"><span data-stu-id="66cf6-114">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a><span data-ttu-id="66cf6-115">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="66cf6-115">Update the code</span></span>

1. <span data-ttu-id="66cf6-p102">Файл **Home.html** содержит HTML-контент, который будет отображаться в области задач надстройки. В файле **Home.html** замените элемент `<body>` на приведенную ниже часть кода и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p102">**Home.html** specifies the HTML that will be rendered in the add-in's task pane. In **Home.html**, replace the `<body>` element with the following markup and save the file.</span></span>

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

2. <span data-ttu-id="66cf6-p103">Откройте файл **Home.js** в корневой папке проекта веб-приложения. Этот файл содержит скрипт надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p103">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

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

3. <span data-ttu-id="66cf6-p104">Откройте файл **Home.css** в корневой папке проекта веб-приложения. Этот файл определяет специальные стили надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p104">Open the file **Home.css** in the root of the web application project. This file specifies the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span> 

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

### <a name="update-the-manifest"></a><span data-ttu-id="66cf6-124">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="66cf6-124">Update the manifest</span></span>

1. <span data-ttu-id="66cf6-125">Откройте XML-файл манифеста в проекте надстройки.</span><span class="sxs-lookup"><span data-stu-id="66cf6-125">Open the XML manifest file in the add-in project.</span></span> <span data-ttu-id="66cf6-126">Этот файл определяет параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="66cf6-126">This file defines the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="66cf6-p106">Элемент `ProviderName` содержит заполнитель. Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p106">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="66cf6-p107">Атрибут `DefaultValue` элемента `DisplayName` содержит заполнитель. Замените его на строку **Моя надстройка Office**.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p107">The `DefaultValue` attribute of the `DisplayName` element has a placeholder. Replace it with **My Office Add-in**.</span></span>

4. <span data-ttu-id="66cf6-p108">Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его строкой **Надстройка области задач для Excel**.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p108">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

5. <span data-ttu-id="66cf6-133">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="66cf6-133">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a><span data-ttu-id="66cf6-134">Проверка</span><span class="sxs-lookup"><span data-stu-id="66cf6-134">Try it out</span></span>

1. <span data-ttu-id="66cf6-p109">Протестируйте новую надстройку Excel в Visual Studio, нажав клавишу **F5** или кнопку **Запустить**, чтобы запустить Excel с кнопкой надстройки **Показать область задач** на ленте. Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p109">Using Visual Studio, test the newly created Excel add-in by pressing **F5** or choosing the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="66cf6-137">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="66cf6-137">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. <span data-ttu-id="66cf6-139">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="66cf6-139">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="66cf6-140">В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="66cf6-140">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

# <a name="any-editortabvisual-studio-code"></a>[<span data-ttu-id="66cf6-142">Любой редактор</span><span class="sxs-lookup"><span data-stu-id="66cf6-142">Any editor</span></span>](#tab/visual-studio-code)

### <a name="prerequisites"></a><span data-ttu-id="66cf6-143">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="66cf6-143">Prerequisites</span></span>

- [<span data-ttu-id="66cf6-144">Node.js</span><span class="sxs-lookup"><span data-stu-id="66cf6-144">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="66cf6-145">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="66cf6-145">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```command&nbsp;line
    npm install -g yo generator-office
    ```

### <a name="create-the-web-app"></a><span data-ttu-id="66cf6-146">Создание веб-приложения</span><span class="sxs-lookup"><span data-stu-id="66cf6-146">Create the web app</span></span>

1. <span data-ttu-id="66cf6-147">Создайте проект надстройки Excel помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="66cf6-147">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="66cf6-148">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="66cf6-148">Run the following command and then answer the prompts as follows:</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="66cf6-149">**Выберите тип проекта:** `Office Add-in project using Jquery framework`</span><span class="sxs-lookup"><span data-stu-id="66cf6-149">**Choose a project type:** `Office Add-in project using Jquery framework`</span></span>
    - <span data-ttu-id="66cf6-150">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="66cf6-150">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="66cf6-151">**Как вы хотите назвать надстройку?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="66cf6-151">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="66cf6-152">**Какое клиентское приложение Office должно поддерживаться?** `Excel`</span><span class="sxs-lookup"><span data-stu-id="66cf6-152">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Генератор Yeoman](../images/yo-office-jquery.png)

    <span data-ttu-id="66cf6-154">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="66cf6-154">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="66cf6-155">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="66cf6-155">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

### <a name="update-the-code"></a><span data-ttu-id="66cf6-156">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="66cf6-156">Update the code</span></span> 

1. <span data-ttu-id="66cf6-p111">В редакторе кода откройте файл **index.html** из корневой папки проекта. Этот файл содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p111">In your code editor, open **index.html** in the root of the project. This file specifies the HTML that will be rendered in the add-in's task pane.</span></span> 

2. <span data-ttu-id="66cf6-159">Замените тег `body` в файле **index.html** приведенной ниже разметкой и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="66cf6-159">Within **index.html**, replace the `body` tag with the following markup and save the file.</span></span>

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
        <script type="text/javascript" src="node_modules/jquery/dist/jquery.js"></script>
        <script type="text/javascript" src="node_modules/office-ui-fabric-js/dist/js/fabric.js"></script>
    </body>
    ```

3. <span data-ttu-id="66cf6-160">Откройте файл **src\index.js**, чтобы указать скрипт для надстройки.</span><span class="sxs-lookup"><span data-stu-id="66cf6-160">Open the file **src\index.js** to specify the script for the add-in.</span></span> <span data-ttu-id="66cf6-161">Замените все его содержимое следующим кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="66cf6-161">Replace the entire contents with the following code and save the file.</span></span>

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

4. <span data-ttu-id="66cf6-p113">Откройте файл **app.css**, чтобы указать собственные стили для надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p113">Open the file **app.css** to specify the custom styles for the add-in. Replace the entire contents with the following code and save the file.</span></span>

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

### <a name="update-the-manifest"></a><span data-ttu-id="66cf6-164">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="66cf6-164">Update the manifest</span></span>

1. <span data-ttu-id="66cf6-165">Откройте файл **manifest.xml**, чтобы определить параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="66cf6-165">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span>

2. <span data-ttu-id="66cf6-p114">Элемент `ProviderName` содержит заполнитель. Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p114">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="66cf6-p115">Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его строкой **Надстройка области задач для Excel**.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p115">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="66cf6-170">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="66cf6-170">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="start-the-dev-server"></a><span data-ttu-id="66cf6-171">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="66cf6-171">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a><span data-ttu-id="66cf6-172">Проверка</span><span class="sxs-lookup"><span data-stu-id="66cf6-172">Try it out</span></span>

1. <span data-ttu-id="66cf6-173">Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="66cf6-173">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="66cf6-174">[Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="66cf6-174">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="66cf6-175">[Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="66cf6-175">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="66cf6-176">[iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="66cf6-176">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="66cf6-177">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="66cf6-177">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="66cf6-179">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="66cf6-179">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="66cf6-180">В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="66cf6-180">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

---

## <a name="next-steps"></a><span data-ttu-id="66cf6-182">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="66cf6-182">Next steps</span></span>

<span data-ttu-id="66cf6-p116">Поздравляем, вы успешно создали надстройку Excel с помощью jQuery! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="66cf6-p116">Congratulations, you've successfully created an Excel add-in using jQuery! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="66cf6-185">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="66cf6-185">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="66cf6-186">См. также</span><span class="sxs-lookup"><span data-stu-id="66cf6-186">See also</span></span>

* [<span data-ttu-id="66cf6-187">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="66cf6-187">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="66cf6-188">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="66cf6-188">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="66cf6-189">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="66cf6-189">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="66cf6-190">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="66cf6-190">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
