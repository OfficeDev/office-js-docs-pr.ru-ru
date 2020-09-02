---
title: Преобразование проекта надстройки Office в Visual Studio в TypeScript
description: Сведения о том, как преобразовать проект надстройки Office в Visual Studio для использования TypeScript.
ms.date: 09/01/2020
localization_priority: Normal
ms.openlocfilehash: 716222998666b014ab014298a4165678cf747165
ms.sourcegitcommit: 4adfc368a366f00c3f3d7ed387f34aaecb47f17c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/01/2020
ms.locfileid: "47326298"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="77e4e-103">Преобразование проекта надстройки Office в Visual Studio в TypeScript</span><span class="sxs-lookup"><span data-stu-id="77e4e-103">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="77e4e-104">Вы можете использовать шаблон надстройки Office в Visual Studio, чтобы создать надстройку с использованием JavaScript, а затем преобразовать этот проект в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="77e4e-104">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="77e4e-105">В этой статье описан процесс преобразования для надстройки Excel.</span><span class="sxs-lookup"><span data-stu-id="77e4e-105">This article describes this conversion process for an Excel add-in.</span></span> <span data-ttu-id="77e4e-106">Таким же образом в Visual Studio можно преобразовывать и другие проекты надстроек Office из JavaScript в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="77e4e-106">You can use the same process to convert other types of Office Add-in projects from JavaScript to TypeScript in Visual Studio.</span></span>

> [!IMPORTANT]
> <span data-ttu-id="77e4e-107">В этой статье описываются *минимальные* действия, которые необходимо выполнить, чтобы убедиться, что при нажатии клавиши F5 код будет передаваться в JavaScript, который затем неопубликованные автоматически в Office.</span><span class="sxs-lookup"><span data-stu-id="77e4e-107">This article describes the *minimal* steps necessary to ensure that, when you press F5, the code will be transpiled to JavaScript which is then sideloaded automatically into Office.</span></span> <span data-ttu-id="77e4e-108">Однако код не является "TypeScript".</span><span class="sxs-lookup"><span data-stu-id="77e4e-108">However, the code is not very "TypeScripty".</span></span> <span data-ttu-id="77e4e-109">Например, переменные объявляются с помощью `var` ключевого слова, а `let` не объявления с указанным типом.</span><span class="sxs-lookup"><span data-stu-id="77e4e-109">For example, variables are declared with the `var` keyword instead of `let` and they are not declared with a specified type.</span></span> <span data-ttu-id="77e4e-110">Чтобы использовать все преимущества строгой типизации TypeScript, рассмотрите возможность внесения дальнейших изменений в код.</span><span class="sxs-lookup"><span data-stu-id="77e4e-110">To take full advantage of the strong typing of TypeScript, consider making further changes to the code.</span></span> 

> [!NOTE]
> <span data-ttu-id="77e4e-111">Чтобы создать проект надстройки Office на TypeScript без использования Visual Studio, следуйте указаниям из раздела "Генератор Yeoman" любого [5-минутного руководства по началу работы](/office/dev/add-ins/) и выберите `TypeScript` по соответствующему запросу [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="77e4e-111">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Yeoman generator" section of any [5-minute quick start](/office/dev/add-ins/) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="77e4e-112">Предварительные требования</span><span class="sxs-lookup"><span data-stu-id="77e4e-112">Prerequisites</span></span>

- <span data-ttu-id="77e4e-113">[Visual Studio 2019](https://www.visualstudio.com/vs/) с установленной рабочей нагрузкой **Разработка надстроек для Office и SharePoint**</span><span class="sxs-lookup"><span data-stu-id="77e4e-113">[Visual Studio 2019](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!TIP]
    > <span data-ttu-id="77e4e-114">Если вы уже установили Visual Studio 2019, [используйте установщик Visual Studio](/visualstudio/install/modify-visual-studio), чтобы убедиться, что также установлена рабочая нагрузка **Разработка надстроек для Office и SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-114">If you've previously installed Visual Studio 2019, [use the Visual Studio Installer](/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> <span data-ttu-id="77e4e-115">Если эта рабочая нагрузка еще не установлена, используйте установщик Visual Studio, чтобы [установить ее](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span><span class="sxs-lookup"><span data-stu-id="77e4e-115">If this workload is not yet installed, use the Visual Studio Installer to [install it](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-workloads).</span></span>

- <span data-ttu-id="77e4e-116">Пакет SDK для TypeScript версии 2.3 или более поздней (для Visual Studio 2019)</span><span class="sxs-lookup"><span data-stu-id="77e4e-116">TypeScript SDK version 2.3 or later (for Visual Studio 2019)</span></span>

    > [!TIP]
    > <span data-ttu-id="77e4e-117">В [установщике Visual Studio](/visualstudio/install/modify-visual-studio) выберите вкладку **Отдельные компоненты** и прокрутите вниз до раздела **Пакеты SDK, библиотеки и платформы**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-117">In the [Visual Studio Installer](/visualstudio/install/modify-visual-studio), select the **Individual components** tab and then scroll down to the **SDKs, libraries, and frameworks** section.</span></span> <span data-ttu-id="77e4e-118">Убедитесь, что в этом разделе выбран хотя бы один из компонентов **Пакет SDK для TypeScript** (версии 2.3 или более поздней).</span><span class="sxs-lookup"><span data-stu-id="77e4e-118">Within that section, ensure that at least one of the **TypeScript SDK** components (version 2.3 or later) is selected.</span></span> <span data-ttu-id="77e4e-119">Если не выбран ни один из компонентов **Пакет для TypeScript**, выберите последнюю доступную версию пакета SDK и нажмите кнопку **Изменить**, чтобы [установить этот отдельный компонент](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span><span class="sxs-lookup"><span data-stu-id="77e4e-119">If none of the **TypeScript SDK** components are selected, select the latest available version of the SDK and then choose the **Modify** button to [install that individual component](/visualstudio/install/modify-visual-studio?view=vs-2019#modify-individual-components).</span></span> 

- <span data-ttu-id="77e4e-120">Excel 2016 или более поздней версии</span><span class="sxs-lookup"><span data-stu-id="77e4e-120">Excel 2016 or later</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="77e4e-121">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="77e4e-121">Create the add-in project</span></span>

1. <span data-ttu-id="77e4e-122">В Visual Studio выберите пункт **Создать проект**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-122">In Visual Studio, choose **Create a new project**.</span></span>

2. <span data-ttu-id="77e4e-123">Используя поле поиска, введите **надстройка**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-123">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="77e4e-124">Выберите вариант **Веб-надстройка Excel** и нажмите кнопку **Далее**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-124">Choose **Excel Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="77e4e-125">Присвойте проекту имя и нажмите кнопку **Создать**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-125">Name your project and select **Create**.</span></span>

4. <span data-ttu-id="77e4e-126">В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в Excel**, а затем нажмите кнопку **Готово**, чтобы создать проект.</span><span class="sxs-lookup"><span data-stu-id="77e4e-126">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="77e4e-p106">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-p106">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="77e4e-129">Преобразование проекта надстройки в TypeScript</span><span class="sxs-lookup"><span data-stu-id="77e4e-129">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="77e4e-130">Найдите файл **Home.js** и переименуйте его в **Home.ts**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-130">Find the **Home.js** file and rename it to **Home.ts**.</span></span>

2. <span data-ttu-id="77e4e-131">Найдите файл **./Functions/FunctionFile.js** и переименуйте его в **FunctionFile.ts**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-131">Find the **./Functions/FunctionFile.js** file and rename it to **FunctionFile.ts**.</span></span>

3. <span data-ttu-id="77e4e-132">Найдите файл **./Scripts/MessageBanner.js** и переименуйте его в **MessageBanner.ts**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-132">Find the **./Scripts/MessageBanner.js** file and rename it to **MessageBanner.ts**.</span></span>

4. <span data-ttu-id="77e4e-133">На вкладке **Средства** выберите **Диспетчер пакетов NuGet** и щелкните пункт **Управление пакетами NuGet для решения...**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-133">From the **Tools** tab, choose **NuGet Package Manager** and then select **Manage NuGet Packages for Solution...**.</span></span>

5. <span data-ttu-id="77e4e-134">После выбора вкладки **Обзор** введите **jQuery. TypeScript. DefinitelyTyped**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-134">With the **Browse** tab selected, enter **jquery.TypeScript.DefinitelyTyped**.</span></span> <span data-ttu-id="77e4e-135">Установите этот пакет или обновите его, если он уже установлен.</span><span class="sxs-lookup"><span data-stu-id="77e4e-135">Install this package, or update it if it's already installed.</span></span> <span data-ttu-id="77e4e-136">Это обеспечит включение определений TypeScript для jQuery в проект.</span><span class="sxs-lookup"><span data-stu-id="77e4e-136">This will ensure the jQuery TypeScript definitions are included in your project.</span></span> <span data-ttu-id="77e4e-137">Пакеты для jQuery отображаются в файле, созданном Visual Studio, который называется **packages.config**.</span><span class="sxs-lookup"><span data-stu-id="77e4e-137">The packages for jQuery appear in a file generated by Visual Studio, called **packages.config**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="77e4e-p108">В проекте TypeScript могут быть как файлы TypeScript, так и файлы JavaScript, это не повлияет на компиляцию. Потому что TypeScript — это типизированная расширенная версия языка JavaScript. Код TypeScript компилируется в JavaScript.</span><span class="sxs-lookup"><span data-stu-id="77e4e-p108">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span>

6. <span data-ttu-id="77e4e-140">В **Home.ts** найдите строку `Office.initialize = function (reason) {` и добавьте строку сразу после нее для полизаполнения глобального объекта `window.Promise`, как показано здесь:</span><span class="sxs-lookup"><span data-stu-id="77e4e-140">In **Home.ts**, find the line `Office.initialize = function (reason) {` and add a line immediately after it to polyfill the global `window.Promise`, as shown here:</span></span>

    ```TypeScript
    Office.initialize = function (reason) {
        // add the following line
        (window as any).Promise = OfficeExtension.Promise;
        ...
    ```

7. <span data-ttu-id="77e4e-141">В **Home.ts** найдите функцию `displaySelectedCells`, замените всю функцию приведенным ниже кодом и сохраните файл:</span><span class="sxs-lookup"><span data-stu-id="77e4e-141">In **Home.ts**, find the `displaySelectedCells` function, replace the entire function with the following code, and save the file:</span></span>

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

8. <span data-ttu-id="77e4e-142">В **./Scripts/MessageBanner.ts** найдите строку `_onResize(null);` и замените ее указанным ниже кодом:</span><span class="sxs-lookup"><span data-stu-id="77e4e-142">In **./Scripts/MessageBanner.ts**, find the line `_onResize(null);` and replace it with the following:</span></span>

    ```TypeScript
    _onResize();
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="77e4e-143">Запуск преобразованного проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="77e4e-143">Run the converted add-in project</span></span>

1. <span data-ttu-id="77e4e-p109">В Visual Studio нажмите клавишу **F5** или кнопку **Запустить**, чтобы запустить Excel с кнопкой **Показать область задач** на ленте. Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="77e4e-p109">In Visual Studio, press **F5** or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="77e4e-146">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="77e4e-146">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="77e4e-147">Выберите на листе девять ячеек, содержащих числа.</span><span class="sxs-lookup"><span data-stu-id="77e4e-147">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="77e4e-148">Нажмите кнопку **Highlight** (Выделить) в области задач, чтобы выделить в выбранном диапазоне ячейку, содержащую самое большое значение.</span><span class="sxs-lookup"><span data-stu-id="77e4e-148">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="77e4e-149">Файл с кодом Home.ts</span><span class="sxs-lookup"><span data-stu-id="77e4e-149">Home.ts code file</span></span>

<span data-ttu-id="77e4e-p110">Для справки в приведенном ниже фрагменте кода показано содержимое файла **Home.ts** после применения вышеописанных изменений. Этот код включает минимальное количество изменений, необходимое для запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="77e4e-p110">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied. This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="77e4e-152">См. также</span><span class="sxs-lookup"><span data-stu-id="77e4e-152">See also</span></span>

- [<span data-ttu-id="77e4e-153">Обсуждение реализации обещаний на сайте StackOverflow</span><span class="sxs-lookup"><span data-stu-id="77e4e-153">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
- [<span data-ttu-id="77e4e-154">Примеры надстроек Office на сайте GitHub</span><span class="sxs-lookup"><span data-stu-id="77e4e-154">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
