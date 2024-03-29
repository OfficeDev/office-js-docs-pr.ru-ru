---
title: Преобразование проекта надстройки Office в Visual Studio в TypeScript
description: Узнайте, как преобразовать проект надстройки Office в Visual Studio для использования TypeScript.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: daa81c3785484083aa49516b04491acad1404884
ms.sourcegitcommit: df7964b6509ee6a807d754fbe895d160bc52c2d3
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/20/2022
ms.locfileid: "66889359"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>Преобразование проекта надстройки Office в Visual Studio в TypeScript

Вы можете использовать шаблон надстройки Office в Visual Studio, чтобы создать надстройку с использованием JavaScript, а затем преобразовать этот проект в TypeScript. В этой статье описан процесс преобразования для надстройки Excel. Таким же образом в Visual Studio можно преобразовывать и другие проекты надстроек Office из JavaScript в TypeScript.

> [!IMPORTANT]
> В этой статье описаны  минимальные действия, необходимые для того, чтобы при нажатии клавиши F5 код автоматически перезагружается в Office. Однако код не очень "TypeScripty". Например, переменные объявляются с ключевым словом `var` `let` `const` или не объявляются с указанным типом. Чтобы воспользоваться всеми преимуществами строгой типизации TypeScript, рассмотрите возможность внесения дополнительных изменений в код.

> [!NOTE]
> Чтобы создать проект надстройки Office на TypeScript без использования Visual Studio, следуйте указаниям из раздела "Генератор Yeoman" любого [5-минутного руководства по началу работы](../index.yml) и выберите `TypeScript` по соответствующему запросу [генератора Yeoman для надстроек Office](yeoman-generator-overview.md).

## <a name="prerequisites"></a>Предварительные требования

- [Visual Studio 2019 или более поздней](https://www.visualstudio.com/vs/) версии с установленной рабочей нагрузкой **разработки Office или SharePoint**

    > [!TIP]
    > Если вы уже установили Visual Studio, [используйте установщик Visual Studio](/visualstudio/install/modify-visual-studio), чтобы убедиться, что также установлена рабочая нагрузка **Разработка надстроек для Office и SharePoint**. Если эта рабочая нагрузка еще не установлена, используйте установщик Visual Studio, чтобы [установить ее](/visualstudio/install/modify-visual-studio#modify-workloads).

- Пакет SDK для TypeScript версии 2.3 или более поздней.

    > [!TIP]
    > В [установщике Visual Studio](/visualstudio/install/modify-visual-studio) выберите вкладку **Отдельные компоненты** и прокрутите вниз до раздела **Пакеты SDK, библиотеки и платформы**. Убедитесь, что в этом разделе выбран хотя бы один из компонентов **Пакет SDK для TypeScript** (версии 2.3 или более поздней). Если ни один из компонентов **пакета SDK для TypeScript** не выбран, выберите последнюю доступную версию пакета SDK и нажмите кнопку **"** Изменить", чтобы установить [этот отдельный компонент](/visualstudio/install/modify-visual-studio?view=vs-2019&preserve-view=true#modify-individual-components).

- Excel 2016 или более поздней версии.

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

1. В Visual Studio выберите пункт **Создать проект**.

1. Используя поле поиска, введите **надстройка**. Выберите вариант **Веб-надстройка Excel** и нажмите кнопку **Далее**.

1. Присвойте проекту имя и нажмите кнопку **Создать**.

1. В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в Excel**, а затем нажмите кнопку **Готово**, чтобы создать проект.

1. Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.

## <a name="convert-the-add-in-project-to-typescript"></a>Преобразование проекта надстройки в TypeScript

1. Найдите файл **Home.js** и переименуйте его в **Home.ts**.

1. Найдите файл **./Functions/FunctionFile.js** и переименуйте его в **FunctionFile.ts**.

1. Найдите файл **./Scripts/MessageBanner.js** и переименуйте его в **MessageBanner.ts**.

1. На вкладке **Средства** выберите **Диспетчер пакетов NuGet** и щелкните пункт **Управление пакетами NuGet для решения...**.

1. Выбрав **вкладку** "Обзор", введите **jquery. TypeScript.DefinitelyTyped**. Установите этот пакет или обновите его, если он уже установлен. Это обеспечит включение определений TypeScript jQuery в проект. Пакеты для jQuery отображаются в файле, созданном Visual Studio, **с** именемpackages.config.

    > [!NOTE]
    > В проекте TypeScript могут быть как файлы TypeScript, так и файлы JavaScript, это не повлияет на компиляцию. Потому что TypeScript — это типизированная расширенная версия языка JavaScript. Код TypeScript компилируется в JavaScript.

1. В **файле Home.ts**`Office.initialize = function (reason) {` найдите строку и добавьте строку сразу после `window.Promise`нее, чтобы выполнить полизаполнение глобального объекта, как показано ниже.

    ```TypeScript
    Office.initialize = function (reason) {
        // Add the following line.
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

1. В **файле Home.ts** найдите функцию `displaySelectedCells` , замените всю функцию приведенным ниже кодом и сохраните файл.

    ```TypeScript
    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            null,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }
    ```

1. В **./Scripts/MessageBanner.ts** найдите строку `_onResize(null);` и замените ее указанным ниже кодом:

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a>Запуск преобразованного проекта надстройки

1. В Visual Studio нажмите клавишу **F5** или кнопку **Запустить**, чтобы запустить Excel с кнопкой **Показать область задач** на ленте. Надстройка будет размещена на локальном сервере IIS.

1. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

1. Выберите на листе девять ячеек, содержащих числа.

1. Нажмите кнопку **Highlight** (Выделить) в области задач, чтобы выделить в выбранном диапазоне ячейку, содержащую самое большое значение.

## <a name="homets-code-file"></a>Файл с кодом Home.ts

Для справки в приведенном ниже фрагменте кода показано содержимое файла **Home.ts** после применения вышеописанных изменений. Этот код включает минимальное количество изменений, необходимое для запуска надстройки.

```typescript
(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        (window as any).Promise = OfficeExtension.Promise;
        $(document).ready(function () {
            // Initialize the notification mechanism and hide it
            var element = document.querySelector('.MessageBanner');
            messageBanner = new components.MessageBanner(element);
            messageBanner.hideBanner();

            // If you're using Excel 2013, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
                $("#template-description").text("This sample will display the value of the cells that you have selected in the spreadsheet.");
                $('#button-text').text("Display!");
                $('#button-desc').text("Display the selection");

                $('#highlight-button').click(displaySelectedCells);
                return;
            }

            $("#template-description").text("This sample highlights the highest value from the cells you have selected in the spreadsheet.");
            $('#button-text').text("Highlight!");
            $('#button-desc').text("Highlights the largest number.");

            loadSampleData();

            // Add a click event handler for the highlight button.
            $('#highlight-button').click(highlightHighestValue);
        });
    };

    function loadSampleData() {
        var values = [
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)],
            [Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000), Math.floor(Math.random() * 1000)]
        ];

        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the active sheet
            var sheet = ctx.workbook.worksheets.getActiveWorksheet();
            // Queue a command to write the sample data to the worksheet
            sheet.getRange("B3:D5").values = values;

            // Run the queued-up commands, and return a promise to indicate task completion
            return ctx.sync();
        })
        .catch(errorHandler);
    }

    function highlightHighestValue() {
        // Run a batch operation against the Excel object model
        Excel.run(function (ctx) {
            // Create a proxy object for the selected range and load its properties
            var sourceRange = ctx.workbook.getSelectedRange().load("values, rowCount, columnCount");

            // Run the queued-up command, and return a promise to indicate task completion
            return ctx.sync()
                .then(function () {
                    var highestRow = 0;
                    var highestCol = 0;
                    var highestValue = sourceRange.values[0][0];

                    // Find the cell to highlight
                    for (var i = 0; i < sourceRange.rowCount; i++) {
                        for (var j = 0; j < sourceRange.columnCount; j++) {
                            if (!isNaN(sourceRange.values[i][j]) && sourceRange.values[i][j] > highestValue) {
                                highestRow = i;
                                highestCol = j;
                                highestValue = sourceRange.values[i][j];
                            }
                        }
                    }

                    cellToHighlight = sourceRange.getCell(highestRow, highestCol);
                    sourceRange.worksheet.getUsedRange().format.fill.clear();
                    sourceRange.worksheet.getUsedRange().format.font.bold = false;

                    // Highlight the cell
                    cellToHighlight.format.fill.color = "orange";
                    cellToHighlight.format.font.bold = true;
                })
                .then(ctx.sync);
        })
        .catch(errorHandler);
    }

    function displaySelectedCells() {
        Office.context.document.getSelectedDataAsync(
            Office.CoercionType.Text,
            null,
            function (result) {
                if (result.status === Office.AsyncResultStatus.Succeeded) {
                    showNotification('The selected text is:', '"' + result.value + '"');
                } else {
                    showNotification('Error', result.error.message);
                }
            });
    }

    // Helper function for treating errors
    function errorHandler(error) {
        // Always be sure to catch any accumulated errors that bubble up from the Excel.run execution
        showNotification("Error", error);
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    }

    // Helper function for displaying notifications
    function showNotification(header, content) {
        $("#notification-header").text(header);
        $("#notification-body").text(content);
        messageBanner.showBanner();
        messageBanner.toggleExpansion();
    }
})();
```

## <a name="see-also"></a>См. также

- [Обсуждение реализации обещаний на сайте StackOverflow](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [Примеры надстроек Office на сайте GitHub](https://github.com/OfficeDev/Office-Add-in-samples)
