# <a name="build-an-excel-add-in-using-angular"></a><span data-ttu-id="d4941-101">Создание надстройки Excel с помощью Angular</span><span class="sxs-lookup"><span data-stu-id="d4941-101">Build an Excel add-in using Angular</span></span>

<span data-ttu-id="d4941-102">В этой статье описывается процесс создания надстройки Excel с помощью Angular и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="d4941-102">In this article, you'll walk you through the process of building an Excel add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="d4941-103">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="d4941-103">Prerequisites</span></span>

- [<span data-ttu-id="d4941-104">Node.js</span><span class="sxs-lookup"><span data-stu-id="d4941-104">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="d4941-105">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="d4941-105">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a><span data-ttu-id="d4941-106">Создание веб-приложения</span><span class="sxs-lookup"><span data-stu-id="d4941-106">Create the web app</span></span>

1. <span data-ttu-id="d4941-107">С помощью генератора Yeoman создайте проект надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="d4941-107">Use the Yeoman generator to create an Outlook add-in project.</span></span> <span data-ttu-id="d4941-108">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже:</span><span class="sxs-lookup"><span data-stu-id="d4941-108">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="d4941-109">**Выберите тип проекта:** `Office Add-in project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="d4941-109">**Choose a project type:** `Office Add-in project using Angular framework`</span></span>
    - <span data-ttu-id="d4941-110">**Выберите тип сценария:** `Typescript`</span><span class="sxs-lookup"><span data-stu-id="d4941-110">**Choose a script type:** `Typescript`</span></span>
    - <span data-ttu-id="d4941-111">**Как вы хотите назвать надстройку?:** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="d4941-111">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="d4941-112">**Какое клиентское приложение Office должно поддерживаться?:** `Excel`</span><span class="sxs-lookup"><span data-stu-id="d4941-112">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Генератор Yeoman](../images/yo-office-excel-angular.png)
    
    <span data-ttu-id="d4941-114">После завершения работы мастера генератор создаст проект и установит поддерживающие компоненты узла.</span><span class="sxs-lookup"><span data-stu-id="d4941-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="d4941-115">Перейдите в корневую папку проекта.</span><span class="sxs-lookup"><span data-stu-id="d4941-115">Navigate to the root folder of the web application project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="d4941-116">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="d4941-116">Update the code</span></span>

1. <span data-ttu-id="d4941-117">В редакторе кода откройте файл **app.css**, добавьте следующие стили в конец файла и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="d4941-117">In your code editor, open the file **app.css**, add the following styles to the end of the file, and save the file.</span></span>

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
        font-family: Arial;
        padding-top: 25px;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
        font-family: Arial;
    }

    .padding {
        padding: 15px;
    }

    .padding-sm {
        padding: 4px;
    }

    .normal-button {
        width: 80px;
        padding: 2px;
    }
    ```

2. <span data-ttu-id="d4941-118">Откройте файл **src/app/app.component.html**, замените все содержимое следующим кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="d4941-118">Open **src/app/app.component.html**, replace file contents with the following HTML, and save the file.</span></span>

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>{{welcomeMessage}}</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <br />
            <div role="button" class="ms-Button" (click)="setColor()">
                <span class="ms-Button-label">Set color</span>
                <span class="ms-Button-icon"><i class="ms-Icon ms-Icon--ChevronRight"></i></span>
            </div>
        </div>
    </div>
    ```

3. <span data-ttu-id="d4941-119">Откройте файл **src/app/app.component.ts**, замените все его содержимое приведенным ниже кодом и сохраните.</span><span class="sxs-lookup"><span data-stu-id="d4941-119">Open **src/app/app.component.ts**, replace file contents with the following code, and save the file.</span></span>

    ```typescript
    import { Component } from '@angular/core';
    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    const template = require('./app.component.html');

    @Component({
        selector: 'app-home',
        template
    })
    export default class AppComponent {
        welcomeMessage = 'Welcome';

        async setColor() {
            try {
                await Excel.run(async context => {
                    const range = context.workbook.getSelectedRange();
                    range.load('address');
                    range.format.fill.color = 'green';
                    await context.sync();
                    console.log(`The range address was ${range.address}.`);
                });
            } catch (error) {
                OfficeHelpers.UI.notify(error);
                OfficeHelpers.Utilities.log(error);
            }
        }

    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="d4941-120">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="d4941-120">Update the manifest</span></span>

1. <span data-ttu-id="d4941-121">Откройте файл **manifest.xml**, чтобы определить параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="d4941-121">Open the file **my-office-add-in-manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="d4941-p102">Элемент `ProviderName` содержит значение заполнителя. Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="d4941-p102">The `ProviderName` element has a placeholder value. Replace it with your name.</span></span>

3. <span data-ttu-id="d4941-p103">Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его на строку **Надстройка области задач для Excel**.</span><span class="sxs-lookup"><span data-stu-id="d4941-p103">The `DefaultValue` attribute of the `Description` element has a placeholder. Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="d4941-126">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="d4941-126">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="d4941-127">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="d4941-127">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="d4941-128">Проверка</span><span class="sxs-lookup"><span data-stu-id="d4941-128">Try it out</span></span>

1. <span data-ttu-id="d4941-129">Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="d4941-129">Follow the instructions for the platform you'll be using to run your add-in and sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="d4941-130">Windows: [загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="d4941-130">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="d4941-131">Excel Online: [загрузка неопубликованных надстроек Office в Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span><span class="sxs-lookup"><span data-stu-id="d4941-131">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)</span></span>
    - <span data-ttu-id="d4941-132">iPad и Mac: [загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="d4941-132">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

   
2. <span data-ttu-id="d4941-133">В Excel перейдите на вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="d4941-133">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="d4941-135">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="d4941-135">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="d4941-136">В области задач нажмите кнопку **Задать цвет**, чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="d4941-136">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="d4941-138">Дальнейшие шаги</span><span class="sxs-lookup"><span data-stu-id="d4941-138">Next steps</span></span>

<span data-ttu-id="d4941-p104">Поздравляем, вы успешно создали надстройку Excel с помощью Angular! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="d4941-p104">Congratulations, you've successfully created an Excel add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="d4941-141">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="d4941-141">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a><span data-ttu-id="d4941-142">См. также</span><span class="sxs-lookup"><span data-stu-id="d4941-142">See also</span></span>

* [<span data-ttu-id="d4941-143">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="d4941-143">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="d4941-144">Основные принципы программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="d4941-144">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="d4941-145">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="d4941-145">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="d4941-146">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="d4941-146">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview?view=office-js)
