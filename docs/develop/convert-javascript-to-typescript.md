---
title: Преобразование проекта надстройки Office в Visual Studio в TypeScript
description: ''
ms.date: 01/19/2018
ms.openlocfilehash: 05e845b9d085b64b0534d28053dcd5ca3c7b403e
ms.sourcegitcommit: 17f60431644b448a4816913039aaebfa328f9b0a
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/25/2018
ms.locfileid: "19476531"
---
# <a name="convert-an-office-add-in-project-in-visual-studio-to-typescript"></a><span data-ttu-id="c2cff-102">Преобразование проекта надстройки Office в Visual Studio в TypeScript</span><span class="sxs-lookup"><span data-stu-id="c2cff-102">Convert an Office Add-in project in Visual Studio to TypeScript</span></span>

<span data-ttu-id="c2cff-103">Вы можете использовать шаблон надстройки Office в Visual Studio, чтобы создать надстройку с использованием JavaScript, а затем преобразовать этот проект в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="c2cff-103">You can use the Office Add-in template in Visual Studio to create an add-in that uses JavaScript, and then convert that add-in project to TypeScript.</span></span> <span data-ttu-id="c2cff-104">Создавая проект надстройки в Visual Studio, вам не придется создавать проект надстройки Office на TypeScript с нуля.</span><span class="sxs-lookup"><span data-stu-id="c2cff-104">By using Visual Studio to create the add-in project, you avoid having to create your Office Add-in TypeScript project from scratch.</span></span> 

<span data-ttu-id="c2cff-105">В этой статье показано, как создать надстройку Excel с помощью Visual Studio, а затем преобразовать проект надстройки из JavaScript в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="c2cff-105">This article shows you how to create an Excel add-in using Visual Studio and then convert the add-in project from JavaScript to TypeScript.</span></span> <span data-ttu-id="c2cff-106">Таким же образом в Visual Studio можно преобразовывать и другие проекты надстроек Office из JavaScript в TypeScript.</span><span class="sxs-lookup"><span data-stu-id="c2cff-106">You can use the same process to convert other types of Office Add-in JavaScript projects to TypeScript in Visual Studio.</span></span>

> [!NOTE]
> <span data-ttu-id="c2cff-107">Чтобы создать проект надстройки Office на TypeScript без использования Visual Studio, следуйте указаниям из раздела "Любой редактор" любого [5-минутного руководства по началу работы](../index.yml) и выберите `TypeScript` по соответствующему запросу [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).</span><span class="sxs-lookup"><span data-stu-id="c2cff-107">To create an Office Add-in TypeScript project without using Visual Studio, follow the instructions in the "Any editor" section of any [5-minute quickstart](../index.yml) and choose `TypeScript` when prompted by the [Yeoman generator for Office Add-ins](https://github.com/OfficeDev/generator-office).</span></span>

## <a name="prerequisites"></a><span data-ttu-id="c2cff-108">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="c2cff-108">Prerequisites</span></span>

- <span data-ttu-id="c2cff-109">[Visual Studio 2017](https://www.visualstudio.com/vs/) с установленной рабочей нагрузкой **Разработка надстроек для Office и SharePoint**</span><span class="sxs-lookup"><span data-stu-id="c2cff-109">[Visual Studio 2017](https://www.visualstudio.com/vs/) with the **Office/SharePoint development** workload installed</span></span>

    > [!NOTE]
    > <span data-ttu-id="c2cff-110">Если вы уже установили Visual Studio 2017, [используйте установщик Visual Studio](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio), чтобы убедиться, что также установлена рабочая нагрузка **Разработка надстроек для Office и SharePoint**.</span><span class="sxs-lookup"><span data-stu-id="c2cff-110">If you've previously installed Visual Studio 2017, [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to ensure that the **Office/SharePoint development** workload is installed.</span></span> 

- <span data-ttu-id="c2cff-111">TypeScript 2.3 для Visual Studio 2017</span><span class="sxs-lookup"><span data-stu-id="c2cff-111">TypeScript 2.3 for Visual Studio 2017</span></span>

    > [!NOTE]
    > <span data-ttu-id="c2cff-112">TypeScript должен быть по умолчанию установлен вместе с Visual Studio 2017, но вы можете убедиться в этом с помощью [Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio).</span><span class="sxs-lookup"><span data-stu-id="c2cff-112">TypeScript should be installed by default with Visual Studio 2017, but you can [use the Visual Studio Installer](https://docs.microsoft.com/en-us/visualstudio/install/modify-visual-studio) to confirm that it is installed.</span></span> <span data-ttu-id="c2cff-113">В Visual Studio Installer выберите вкладку **Отдельные компоненты** и убедитесь, что в разделе **Пакеты SDK, библиотеки и платформы** выбран узел **Пакет SDK для TypeScript 2.3**.</span><span class="sxs-lookup"><span data-stu-id="c2cff-113">In the Visual Studio Installer, select the **Individual components** tab and then verify that **TypeScript 2.3 SDK** is selected under **SDKs, libraries, and frameworks**.</span></span>

- <span data-ttu-id="c2cff-114">Excel 2016</span><span class="sxs-lookup"><span data-stu-id="c2cff-114">Excel 2016</span></span>

## <a name="create-the-add-in-project"></a><span data-ttu-id="c2cff-115">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="c2cff-115">Create the add-in project</span></span>

1. <span data-ttu-id="c2cff-116">Откройте Visual Studio и в строке меню выберите **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="c2cff-116">Open Visual Studio and on the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>

2. <span data-ttu-id="c2cff-117">В списке типов проекта разверните узел **Visual C#** или **Visual Basic**, разверните **Office/SharePoint**, затем выберите **Надстройки** > **Веб-надстройка Excel**.</span><span class="sxs-lookup"><span data-stu-id="c2cff-117">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **Excel Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="c2cff-118">Укажите имя проекта и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="c2cff-118">Name the project, and then choose **OK**.</span></span>

4. <span data-ttu-id="c2cff-119">В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в Excel**, а затем нажмите кнопку **Готово**, чтобы создать проект.</span><span class="sxs-lookup"><span data-stu-id="c2cff-119">In the **Create Office Add-in** dialog window, choose **Add new functionalities to Excel**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="c2cff-p104">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="c2cff-p104">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

## <a name="convert-the-add-in-project-to-typescript"></a><span data-ttu-id="c2cff-122">Преобразование проекта надстройки в TypeScript</span><span class="sxs-lookup"><span data-stu-id="c2cff-122">Convert the add-in project to TypeScript</span></span>

1. <span data-ttu-id="c2cff-123">В **обозревателе решений** измените имя файла**Home.js** на **Home.ts**.</span><span class="sxs-lookup"><span data-stu-id="c2cff-123">In **Solution Explorer**, rename the **Home.js** file to **Home.ts**.</span></span>

    > [!NOTE]
    > <span data-ttu-id="c2cff-p105">В проекте TypeScript могут быть как файлы TypeScript, так и файлы JavaScript, это не повлияет на компиляцию. Потому что TypeScript — это типизированная расширенная версия языка JavaScript. Код TypeScript компилируется в JavaScript.</span><span class="sxs-lookup"><span data-stu-id="c2cff-p105">In your TypeScript project, you can have a mix of TypeScript and JavaScript files and your project will compile. This is because TypeScript is a typed superset of JavaScript that compiles JavaScript.</span></span> 

2. <span data-ttu-id="c2cff-126">Нажмите **Да**, чтобы подтвердить изменение расширения имени файла.</span><span class="sxs-lookup"><span data-stu-id="c2cff-126">Select **Yes** when prompted to confirm that you want to change file name extension.</span></span>

3. <span data-ttu-id="c2cff-127">Создайте файл с именем **Office.d.ts** в корне проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="c2cff-127">Create a new file named **Office.d.ts** in the root of the web application project.</span></span>

4. <span data-ttu-id="c2cff-128">В веб-браузере откройте [файл определений типов для Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="c2cff-128">In a web browser, open the [type definitions file for Office.js](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/office-js/index.d.ts).</span></span> <span data-ttu-id="c2cff-129">Скопируйте содержимое этого файла в буфер обмена.</span><span class="sxs-lookup"><span data-stu-id="c2cff-129">Copy the contents of this file to your clipboard.</span></span>

5. <span data-ttu-id="c2cff-130">В Visual Studio откройте файл **Office.d.ts**, вставьте в него содержимое буфера обмена и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="c2cff-130">In Visual Studio, open the **Office.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

6. <span data-ttu-id="c2cff-131">Создайте файл с именем **jQuery.d.ts** в корне проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="c2cff-131">Create a new file named **jQuery.d.ts** in the root of the web application project.</span></span>

7. <span data-ttu-id="c2cff-132">В веб-браузере откройте [файл определений типов для jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts).</span><span class="sxs-lookup"><span data-stu-id="c2cff-132">In a web browser, open the [type definitions file for jQuery](https://raw.githubusercontent.com/DefinitelyTyped/DefinitelyTyped/master/types/jquery/index.d.ts).</span></span> <span data-ttu-id="c2cff-133">Скопируйте содержимое этого файла в буфер обмена.</span><span class="sxs-lookup"><span data-stu-id="c2cff-133">Copy the contents of this file to your clipboard.</span></span>

8. <span data-ttu-id="c2cff-134">В Visual Studio откройте файл **jQuery.d.ts**, вставьте в него содержимое буфера обмена и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="c2cff-134">In Visual Studio, open the **jQuery.d.ts** file, paste the contents of your clipboard into this file, and save the file.</span></span>

9. <span data-ttu-id="c2cff-135">В Visual Studio создайте файл с именем **tsconfig.json** в корне проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="c2cff-135">In Visual Studio, create a new file named **tsconfig.json** in the root of the web application project.</span></span>

10. <span data-ttu-id="c2cff-136">Откройте файл **tsconfig.json**, добавьте в него приведенное ниже содержимое и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="c2cff-136">Open the **tsconfig.json** file, add the following content to the file, and save the file:</span></span>

    ```javascript
    {
        "compilerOptions": {
            "skipLibCheck": true,
            "lib": [ "es5", "dom", "es2015.promise" ]
        }
    }
    ```

11. <span data-ttu-id="c2cff-137">Откройте файл **Home.ts** и добавьте в его начале следующее объявление:</span><span class="sxs-lookup"><span data-stu-id="c2cff-137">Open the **Home.ts** file and add the following declaration at the top of the file:</span></span>

    ```javascript
    declare var fabric: any;
    ```

12. <span data-ttu-id="c2cff-138">В файле **Home.ts** замените **'1.1'** на **1.1** (то есть удалите кавычки) в приведенной ниже строке, а затем сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="c2cff-138">In the **Home.ts** file, change **'1.1'** to **1.1** (that is, remove the quotation marks) in the following line, and save the file:</span></span>

    ```javascript
    if (!Office.context.requirements.isSetSupported('ExcelApi', '1.1')) {
    ```

## <a name="run-the-converted-add-in-project"></a><span data-ttu-id="c2cff-139">Запуск преобразованного проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="c2cff-139">Run the converted add-in project</span></span>

1. <span data-ttu-id="c2cff-p108">В Visual Studio нажмите клавишу F5 или кнопку **Запустить**, чтобы запустить Excel с кнопкой **Show Taskpane** (Показать область задач) на ленте. Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="c2cff-p108">In Visual Studio, press F5 or choose the **Start** button to launch Excel with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

2. <span data-ttu-id="c2cff-142">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2cff-142">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

3. <span data-ttu-id="c2cff-143">Выберите на листе девять ячеек, содержащих числа.</span><span class="sxs-lookup"><span data-stu-id="c2cff-143">In the worksheet, select the nine cells that contain numbers.</span></span>

4. <span data-ttu-id="c2cff-144">Нажмите кнопку **Highlight** (Выделить) в области задач, чтобы выделить в выбранном диапазоне ячейку, содержащую самое большое значение.</span><span class="sxs-lookup"><span data-stu-id="c2cff-144">Press the **Highlight** button on the task pane to highlight the cell in the selected range that contains the highest value.</span></span>

## <a name="homets-code-file"></a><span data-ttu-id="c2cff-145">Файл с кодом Home.ts</span><span class="sxs-lookup"><span data-stu-id="c2cff-145">Home.ts code file</span></span>

<span data-ttu-id="c2cff-146">Для справки в приведенном ниже фрагменте кода показано содержимое файла **Home.ts** после применения вышеописанных изменений.</span><span class="sxs-lookup"><span data-stu-id="c2cff-146">For your reference, the following code snippet shows the contents of the **Home.ts** file after the previously described changes have been applied.</span></span> <span data-ttu-id="c2cff-147">Этот код включает минимальное количество изменений, необходимое для запуска надстройки.</span><span class="sxs-lookup"><span data-stu-id="c2cff-147">This code includes the minimum number of changes needed in order for your add-in to run.</span></span>

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
            
            // If not using Excel 2016, use fallback logic.
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

## <a name="see-also"></a><span data-ttu-id="c2cff-148">См. также</span><span class="sxs-lookup"><span data-stu-id="c2cff-148">See also</span></span>

* [<span data-ttu-id="c2cff-149">Обсуждение реализации обещаний на сайте StackOverflow</span><span class="sxs-lookup"><span data-stu-id="c2cff-149">Promise implementation discussion on StackOverflow</span></span>](https://stackoverflow.com/questions/44461312/office-addins-file-in-its-typescript-version-doesnt-work)
* [<span data-ttu-id="c2cff-150">Примеры надстроек Office на сайте GitHub</span><span class="sxs-lookup"><span data-stu-id="c2cff-150">Office Add-in samples on GitHub</span></span>](https://github.com/officedev)
