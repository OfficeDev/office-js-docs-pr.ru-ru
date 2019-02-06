---
title: Создание надстройки Excel с помощью React
description: ''
ms.date: 10/19/2018
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 02fd62dca59136fe85ff9b29a6b44576f1ceb8e9
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742368"
---
# <a name="build-an-excel-add-in-using-react"></a><span data-ttu-id="44761-102">Создание надстройки Excel с помощью React</span><span class="sxs-lookup"><span data-stu-id="44761-102">Build an Excel add-in using React</span></span>

<span data-ttu-id="44761-103">В этой статье описывается процесс создания надстройки Excel с помощью React и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="44761-103">In this article, you'll walk through the process of building an Excel add-in using React and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="44761-104">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="44761-104">Prerequisites</span></span>

- [<span data-ttu-id="44761-105">Node.js</span><span class="sxs-lookup"><span data-stu-id="44761-105">Node.js</span></span>](https://nodejs.org)

- <span data-ttu-id="44761-106">Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="44761-106">Install the latest version of [Yeoman](https://github.com/yeoman/yo) and the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office) globally.</span></span>
    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a><span data-ttu-id="44761-107">Создание веб-приложения</span><span class="sxs-lookup"><span data-stu-id="44761-107">Create the web app</span></span>

1. <span data-ttu-id="44761-108">Создайте проект надстройки Excel помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="44761-108">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="44761-109">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="44761-109">Run the following command and then answer the prompts as follows:</span></span>

    ```bash
    yo office
    ```

    - <span data-ttu-id="44761-110">**Выберите тип проекта:** `Office Add-in project using React framework`</span><span class="sxs-lookup"><span data-stu-id="44761-110">**Choose a project type:** `Office Add-in project using React framework`</span></span>
    - <span data-ttu-id="44761-111">**Как вы хотите назвать надстройку?** `My Office Add-in`</span><span class="sxs-lookup"><span data-stu-id="44761-111">**What do you want to name your add-in?:** `My Office Add-in`</span></span>
    - <span data-ttu-id="44761-112">**Какое клиентское приложение Office должно поддерживаться?** `Excel`</span><span class="sxs-lookup"><span data-stu-id="44761-112">**Which Office client application would you like to support?:** `Excel`</span></span>

    ![Генератор Yeoman](../images/yo-office-excel-react.png)
    
    <span data-ttu-id="44761-114">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="44761-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

2. <span data-ttu-id="44761-115">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="44761-115">Navigate to the root folder of the project.</span></span>

    ```bash
    cd "My Office Add-in"
    ```

## <a name="update-the-code"></a><span data-ttu-id="44761-116">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="44761-116">Update the code</span></span>

1. <span data-ttu-id="44761-117">В редакторе кода откройте файл **src/styles.less**, добавьте указанные ниже стили в конец файла и сохраните его.</span><span class="sxs-lookup"><span data-stu-id="44761-117">In your code editor, open the file **src/styles.less**, add the following styles to the end of the file, and save the file.</span></span>

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

2. <span data-ttu-id="44761-118">Шаблон проекта, созданный генератором надстройки Office Yeoman включает в себя компонент React, который не является обязательным для быстрого запуска.</span><span class="sxs-lookup"><span data-stu-id="44761-118">The project template that the Office Add-ins Yeoman generator created includes a React component that is not needed for this quick start.</span></span> <span data-ttu-id="44761-119">Удалите файл **src/components/HeroList.tsx**.</span><span class="sxs-lookup"><span data-stu-id="44761-119">Delete the file **src/components/HeroList.tsx**.</span></span>

3. <span data-ttu-id="44761-120">Откройте файл **src/components/Header.tsx**, замените все его содержимое приведенным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="44761-120">Open the file **src/components/Header.tsx**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';

    export interface HeaderProps {
        title: string;
    }

    export class Header extends React.Component<HeaderProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-header'>
                    <div className='padding'>
                        <h1>{this.props.title}</h1>
                    </div>
                </div>
            );
        }
    }
    ```

4. <span data-ttu-id="44761-121">Создайте новый компонент React с именем **Content.tsx** в папке**src/components**, добавьте приведенный ниже код и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="44761-121">Create a new React component named **Content.tsx** in the **src/components** folder, add the following code, and save the file.</span></span>

    ```typescript
    import * as React from 'react';
    import { Button, ButtonType } from 'office-ui-fabric-react';

    export interface ContentProps {
        message: string;
        buttonLabel: string;
        click: any;
    }

    export class Content extends React.Component<ContentProps, any> {
        constructor(props, context) {
            super(props, context);
        }

        render() {
            return (
                <div id='content-main'>
                    <div className='padding'>
                        <p>{this.props.message}</p>
                        <br />
                        <h3>Try it out</h3>
                        <br/>
                        <Button className='normal-button' buttonType={ButtonType.hero} onClick={this.props.click}>{this.props.buttonLabel}</Button>
                    </div>
                </div>
            );
        }
    }
    ```

5. <span data-ttu-id="44761-122">Откройте файл **src/components/App.tsx**, замените все его содержимое приведенным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="44761-122">Open the file **src/components/App.tsx**, replace the entire contents with the following code, and save the file.</span></span>

    ```typescript
    /* global Office, Excel */

    import * as React from 'react';
    import { Header } from './Header';
    import { Content } from './Content';
    import Progress from './Progress';

    import * as OfficeHelpers from '@microsoft/office-js-helpers';

    export interface AppProps {
        title: string;
        isOfficeInitialized: boolean;
    }

    export interface AppState {
    }

    export default class App extends React.Component<AppProps, AppState> {
        constructor(props, context) {
            super(props, context);
        }

        setColor = async () => {
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

        render() {
            const {
                title,
                isOfficeInitialized,
            } = this.props;

            if (!isOfficeInitialized) {
                return (
                    <Progress
                        title={title}
                        logo='assets/logo-filled.png'
                        message='Please sideload your addin to see app body.'
                    />
                );
            }

            return (
                <div className='ms-welcome'>
                    <Header title='Welcome' />
                    <Content message='Choose the button below to set the color of the selected range to green.' buttonLabel='Set color' click={this.setColor} />
                </div>
            );
        }
    }
    ```

## <a name="update-the-manifest"></a><span data-ttu-id="44761-123">Обновление манифеста</span><span class="sxs-lookup"><span data-stu-id="44761-123">Update the manifest</span></span>

1. <span data-ttu-id="44761-124">Откройте файл **manifest.xml**, чтобы определить параметры и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="44761-124">Open the file **manifest.xml** to define the add-in's settings and capabilities.</span></span> 

2. <span data-ttu-id="44761-125">Элемент `ProviderName` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="44761-125">The `ProviderName` element has a placeholder value.</span></span> <span data-ttu-id="44761-126">Замените его на свое имя.</span><span class="sxs-lookup"><span data-stu-id="44761-126">Replace it with your name.</span></span>

3. <span data-ttu-id="44761-127">Атрибут `DefaultValue` элемента `Description` содержит заполнитель.</span><span class="sxs-lookup"><span data-stu-id="44761-127">The `DefaultValue` attribute of the `Description` element has a placeholder.</span></span> <span data-ttu-id="44761-128">Замените его строкой **Надстройка области задач для Excel**.</span><span class="sxs-lookup"><span data-stu-id="44761-128">Replace it with **A task pane add-in for Excel**.</span></span>

4. <span data-ttu-id="44761-129">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="44761-129">Save the file.</span></span>

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

## <a name="start-the-dev-server"></a><span data-ttu-id="44761-130">Запуск сервера разработки</span><span class="sxs-lookup"><span data-stu-id="44761-130">Start the dev server</span></span>

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

## <a name="try-it-out"></a><span data-ttu-id="44761-131">Проверка</span><span class="sxs-lookup"><span data-stu-id="44761-131">Try it out</span></span>

1. <span data-ttu-id="44761-132">Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.</span><span class="sxs-lookup"><span data-stu-id="44761-132">Follow the instructions for the platform you'll use to run your add-in to sideload the add-in within Excel.</span></span>

    - <span data-ttu-id="44761-133">[Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span><span class="sxs-lookup"><span data-stu-id="44761-133">Windows: [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)</span></span>
    - <span data-ttu-id="44761-134">[Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span><span class="sxs-lookup"><span data-stu-id="44761-134">Excel Online: [Sideload Office Add-ins in Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)</span></span>
    - <span data-ttu-id="44761-135">[iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span><span class="sxs-lookup"><span data-stu-id="44761-135">iPad and Mac: [Sideload Office Add-ins on iPad and Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)</span></span>

2. <span data-ttu-id="44761-136">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="44761-136">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2b.png)

3. <span data-ttu-id="44761-138">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="44761-138">Select any range of cells in the worksheet.</span></span>

4. <span data-ttu-id="44761-139">В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.</span><span class="sxs-lookup"><span data-stu-id="44761-139">In the task pane, choose the **Set color** button to set the color of the selected range to green.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a><span data-ttu-id="44761-141">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="44761-141">Next steps</span></span>

<span data-ttu-id="44761-p105">Поздравляем, вы успешно создали надстройку Excel с помощью React! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="44761-p105">Congratulations, you've successfully created an Excel add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="44761-144">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="44761-144">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="44761-145">См. также</span><span class="sxs-lookup"><span data-stu-id="44761-145">See also</span></span>

* [<span data-ttu-id="44761-146">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="44761-146">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="44761-147">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="44761-147">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="44761-148">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="44761-148">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="44761-149">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="44761-149">Excel JavaScript API reference</span></span>](https://docs.microsoft.com/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
