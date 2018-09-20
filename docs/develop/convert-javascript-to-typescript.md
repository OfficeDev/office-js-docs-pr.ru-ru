---
title: Преобразование проекта надстройки Office в Visual Studio в TypeScript
description: ''
ms.date: 01/19/2018
ms.openlocfilehash: 015fd9d7e9bf4412c09b76f0de5a97c9946e4d58
ms.sourcegitcommit: 3da2038e827dc3f274d63a01dc1f34c98b04557e
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/19/2018
ms.locfileid: "24016334"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a>Преобразование проекта надстройки Office в Visual Studio в TypeScript

Вы можете использовать шаблон надстройки Office в Visual Studio, чтобы создать надстройку с использованием JavaScript, а затем преобразовать этот проект в TypeScript. Создавая проект надстройки в Visual Studio, вам не придется создавать проект надстройки Office на TypeScript с нуля. 

В этой статье показано, как создать надстройку Excel с помощью Visual Studio, а затем преобразовать проект надстройки из JavaScript в TypeScript. Таким же образом в Visual Studio можно преобразовывать и другие проекты надстроек Office из JavaScript в TypeScript.

> [!NOTE]
> Чтобы создать проект надстройки Office на TypeScript без использования Visual Studio, следуйте указаниям из раздела "Любой редактор" любого [5-минутного руководства по началу работы](../index.yml) и выберите `TypeScript` по соответствующему запросу [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).

## <a name="prerequisites"></a>Необходимые компоненты

- [Visual Studio 2017](https://www.visualstudio.com/vs/) с установленной рабочей нагрузкой **Разработка надстроек для Office и SharePoint**

    > [!NOTE]
    > Если вы уже установили Visual Studio 2017, [используйте установщик Visual Studio](https://docs.microsoft.com/visualstudio/install/modify-visual-studio), чтобы убедиться, что также установлена рабочая нагрузка **Разработка надстроек для Office и SharePoint**. 

- TypeScript 2.3 для Visual Studio 2017

    > [!NOTE]
    > TypeScript должен быть по умолчанию установлен вместе с Visual Studio 2017, но вы можете убедиться в этом с помощью [Visual Studio Installer](https://docs.microsoft.com/visualstudio/install/modify-visual-studio). В Visual Studio Installer выберите вкладку **Отдельные компоненты** и убедитесь, что в разделе **Пакеты SDK, библиотеки и платформы** выбран узел **Пакет SDK для TypeScript 2.3**.

- Excel 2016 или более поздняя версия

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

1. Откройте Visual Studio и в строке меню выберите **Файл** > **Создать** > **Проект**.

2. В списке типов проекта разверните узел **Visual C#** или **Visual Basic**, разверните **Office/SharePoint**, затем выберите **Надстройки** > **Веб-надстройка Excel**. 

3. Укажите имя проекта и нажмите кнопку **ОК**.

4. В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в Excel**, а затем нажмите кнопку **Готово**, чтобы создать проект.

5. Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.

## <a name="convert-the-add-in-project-to-typescript"></a>Преобразование проекта надстройки в TypeScript

1. В **обозревателе решений** измените имя файла**Home.js** на **Home.ts**.

    > [!NOTE]
    > В проекте TypeScript могут быть как файлы TypeScript, так и файлы JavaScript, это не повлияет на компиляцию. Потому что TypeScript — это типизированная расширенная версия языка JavaScript. Код TypeScript компилируется в JavaScript. 

2. Нажмите **Да**, чтобы подтвердить изменение расширения имени файла.

3. Создайте файл с именем **Office.d.ts** в корне проекта веб-приложения.

4. В веб-браузере откройте [файл определений типов для Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts). Скопируйте содержимое этого файла в буфер обмена.

5. В Visual Studio откройте файл **Office.d.ts**, вставьте в него содержимое буфера обмена и сохраните файл.

6. Создайте файл с именем **jQuery.d.ts** в корне проекта веб-приложения.

7. В веб-браузере откройте [файл определений типов для jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts). Скопируйте содержимое этого файла в буфер обмена.

8. В Visual Studio откройте файл **jQuery.d.ts**, вставьте в него содержимое буфера обмена и сохраните файл.

9. В Visual Studio создайте файл с именем **tsconfig.json** в корне проекта веб-приложения.

10. Откройте файл **tsconfig.json**, добавьте в него приведенное ниже содержимое и сохраните файл.

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. Откройте файл **Home.ts** и добавьте в его начале следующее объявление:

    ```javascript
    declare var fabric: any;
    ```

12. В файле **Home.ts** замените **'1.1'** на **1.1** (то есть удалите кавычки) в приведенной ниже строке, а затем сохраните файл.

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a>Запуск преобразованного проекта надстройки

1. В Visual Studio нажмите клавишу F5 или кнопку **Запустить**, чтобы запустить Excel с кнопкой **Show Taskpane** (Показать область задач) на ленте. Надстройка будет размещена на локальном сервере IIS.

2. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

3. Выберите на листе девять ячеек, содержащих числа.

4. Нажмите кнопку **Highlight** (Выделить) в области задач, чтобы выделить в выбранном диапазоне ячейку, содержащую самое большое значение.

## <a name="homets-code-file"></a>Файл с кодом Home.ts

Для справки в приведенном ниже фрагменте кода показано содержимое файла **Home.ts** после применения вышеописанных изменений. Этот код включает минимальное количество изменений, необходимое для запуска надстройки.

```javascript
declare var fabric: any;

(function () {
    "use strict";

    var cellToHighlight;
    var messageBanner;

    // The initialize function must be run each time a new page is loaded.
    Office.initialize = function (reason) {
        $(document).ready(function () {
            // Initialize the FabricUI notification mechanism and hide it
            var element = document.querySelector('.ms-MessageBanner');
            messageBanner = new fabric.MessageBanner(element);
            messageBanner.hideBanner();
            
            // If not using Excel 2016 or later, use fallback logic.
            if (!Office.context.requirements.isSetSupported('ExcelApi', 1.1)) {
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
            $('#highlight-button').click(hightlightHighestValue);
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

    function hightlightHighestValue() {
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
        Office.context.document.getSelectedDataAsync(Office.CoercionType.Text,
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

* [Обсуждение реализации обещаний на сайте StackOverflow](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [Примеры надстроек Office на сайте GitHub](https://github.com/officedev)
