---
title: Создание первой надстройки области задач Excel
description: Узнайте, как создать простую надстройку области задач Excel, используя API JS для Office.
ms.date: 07/12/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: f4bd32e9e38e72b8cfb42db9476f6c180b151909
ms.sourcegitcommit: 9bb790f6264f7206396b32a677a9133ab4854d4e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/15/2022
ms.locfileid: "66797521"
---
# <a name="build-an-excel-task-pane-add-in"></a>Создание надстройки области задач Excel

В этой статье вы ознакомитесь с процессом создания надстройки области задач Excel.

## <a name="create-the-add-in"></a>Создание надстройки

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="yeoman-generator"></a>[Генератор Yeoman](#tab/yeomangenerator)

### <a name="prerequisites"></a>Необходимые условия

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

### <a name="create-the-add-in-project"></a>Создание проекта надстройки

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- **Выберите тип проекта:** `Office Add-in Task Pane project`
- **Выберите тип сценария:** `Javascript`
- **Как вы хотите назвать надстройку?** `My Office Add-in`
- **Какое клиентское приложение Office должно поддерживаться?** `Excel`

![Интерфейс командной строки генератора Yeoman для надстроек Office.](../images/yo-office-excel.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

### <a name="explore-the-project"></a>Знакомство с проектом

[!include[Yeoman generator add-in project components](../includes/yo-task-pane-project-components-js.md)]

### <a name="try-it-out"></a>Проверка

1. Перейдите к корневой папке проекта.

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

1. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)]

1. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Меню "Главная" в Excel с выделенной кнопкой "Показать область задач".](../images/excel-quickstart-addin-3b.png)

1. Выберите любой диапазон ячеек на листе.

1. Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.

    ![Открытая область задач надстройки в Excel с выделенной кнопкой "Выполнить" в области задач надстройки.](../images/excel-quickstart-addin-3c.png)

### <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку панели задач Excel! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь [руководством по надстройкам Excel](../tutorials/excel-tutorial.md).

# <a name="visual-studio"></a>[Visual Studio](#tab/visualstudio)

### <a name="prerequisites"></a>Необходимые компоненты

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a>Создание проекта надстройки

1. В Visual Studio выберите пункт **Создать проект**.

1. Используя поле поиска, введите **надстройка**. Выберите вариант **Веб-надстройка Excel** и нажмите кнопку **Далее**.

1. Присвойте проекту имя **ExcelWebAddIn1** и выберите **Создать**.

1. В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в Excel**, а затем нажмите кнопку **Готово**, чтобы создать проект.

1. Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.

### <a name="explore-the-visual-studio-solution"></a>Обзор решения Visual Studio

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a>Обновление кода

1. Файл **Home.html** содержит HTML-контент, который будет отображаться в области задач надстройки. В файле **Home.html** замените элемент `<body>` на приведенную ниже часть кода и сохраните файл.

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

1. Откройте файл **Home.js** в корневой папке проекта веб-приложения. Этот файл содержит скрипт надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.

    [!include[Information about the use of ES6 JavaScript](../includes/modern-js-note.md)]

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

        async function setColor() {
            await Excel.run(async (context) => {
                const range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                await context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

1. Откройте файл **Home.css** в корневой папке проекта веб-приложения. Этот файл определяет специальные стили надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.

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

### <a name="update-the-manifest"></a>Обновление манифеста

1. Откройте **Обозреватель решений**, перейдите к проекту надстройки **ExcelWebAddIn1**, затем откройте каталог **ExcelWebAddIn1Manifest**. Этот каталог содержит **ExcelWebAddIn1.xml** (ваш файл манифеста). XML-файл манифеста определяет параметры и возможности надстройки. Дополнительные сведения о двух проектах, созданных решением Visual Studio, приведены ранее в разделе [Обзор решения Visual Studio](#explore-the-visual-studio-solution).

1. Элемент `ProviderName` содержит заполнитель. Замените его на свое имя.

1. Атрибут `DefaultValue` элемента `DisplayName` содержит заполнитель. Замените его на строку **Моя надстройка Office**.

1. Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его строкой **Надстройка области задач для Excel**.

1. Сохраните файл.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Excel"/>
    ...
    ```

### <a name="try-it-out"></a>Проверка

1. Протестируйте новую надстройку Excel в Visual Studio, нажав клавишу **F5** или кнопку **Запустить**, чтобы запустить Excel с кнопкой надстройки **Показать область задач** на ленте. Надстройка будет размещена на локальном сервере IIS. Если вам будет предложено доверять сертификату, согласитесь, чтобы разрешить надстройке подключиться к приложению Office.

1. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Меню "Главная" в Excel с выделенной кнопкой "Показать область задач".](../images/excel-quickstart-addin-2a.png)

1. Выберите любой диапазон ячеек на листе.

1. В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.

    ![Открытая область задач надстройки в Excel.](../images/excel-quickstart-addin-2c.png)

[!include[Console tool note](../includes/console-tool-note.md)]

### <a name="next-steps"></a>Дальнейшие действия

Поздравляем! Вы успешно создали надстройку области задач Excel! Теперь изучите дополнительные сведения о [разработке надстроек Office с помощью Visual Studio](../develop/develop-add-ins-visual-studio.md).

---

## <a name="see-also"></a>См. также

- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
- [Разработка надстроек Office](../develop/develop-overview.md)
- [Объектная модель JavaScript для Excel в надстройках Office](../excel/excel-add-ins-core-concepts.md)
- [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
- [Справочник по API JavaScript для Excel](../reference/overview/excel-add-ins-reference-overview.md)
- [Использование Visual Studio Code для публикации](../publish/publish-add-in-vs-code.md#using-visual-studio-code-to-publish)
