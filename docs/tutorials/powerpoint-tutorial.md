---
title: Руководство по надстройкам PowerPoint
description: В этом руководстве вы создадите надстройку PowerPoint, которая вставляет изображение, вставляет текст, получает метаданные слайда и перемещается между слайдами.
ms.date: 10/14/2020
ms.prod: powerpoint
localization_priority: Priority
ms.openlocfilehash: eed21235a2de78a160bf9df9253dcabe43bac366
ms.sourcegitcommit: 42e6cfe51d99d4f3f05a3245829d764b28c46bbb
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/23/2020
ms.locfileid: "48740912"
---
# <a name="tutorial-create-a-powerpoint-task-pane-add-in"></a><span data-ttu-id="33a65-103">Учебник: Создание надстройки области задач PowerPoint</span><span class="sxs-lookup"><span data-stu-id="33a65-103">Tutorial: Create a PowerPoint task pane add-in</span></span>

<span data-ttu-id="33a65-104">В этом учебнике вы будете использовать Visual Studio для создания надстройки области задачи PowerPoint, которая:</span><span class="sxs-lookup"><span data-stu-id="33a65-104">In this tutorial, you'll use Visual Studio to create an PowerPoint task pane add-in that:</span></span>

> [!div class="checklist"]
> * <span data-ttu-id="33a65-105">Добавляет фотографию дня из [Bing](https://www.bing.com) на слайд</span><span class="sxs-lookup"><span data-stu-id="33a65-105">Adds the [Bing](https://www.bing.com) photo of the day to a slide</span></span>
> * <span data-ttu-id="33a65-106">Добавляет текст на слайд</span><span class="sxs-lookup"><span data-stu-id="33a65-106">Adds text to a slide</span></span>
> * <span data-ttu-id="33a65-107">Получает метаданные слайды</span><span class="sxs-lookup"><span data-stu-id="33a65-107">Gets slide metadata</span></span>
> * <span data-ttu-id="33a65-108">Выполняет переходы между слайдами</span><span class="sxs-lookup"><span data-stu-id="33a65-108">Navigates between slides</span></span>

## <a name="prerequisites"></a><span data-ttu-id="33a65-109">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="33a65-109">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="create-your-add-in-project"></a><span data-ttu-id="33a65-110">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="33a65-110">Create your add-in project</span></span>

<span data-ttu-id="33a65-111">Выполните указанные ниже действия, чтобы создать проект надстройки PowerPoint с помощью Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="33a65-111">Complete the following steps to create a PowerPoint add-in project using Visual Studio.</span></span>

1. <span data-ttu-id="33a65-112">Выберите **Создание нового проекта**.</span><span class="sxs-lookup"><span data-stu-id="33a65-112">Choose **Create a new project**.</span></span>

2. <span data-ttu-id="33a65-113">Используя поле поиска, введите **надстройка**.</span><span class="sxs-lookup"><span data-stu-id="33a65-113">Using the search box, enter **add-in**.</span></span> <span data-ttu-id="33a65-114">Выберите вариант **Веб-надстройка PowerPoint** и нажмите кнопку **Далее**.</span><span class="sxs-lookup"><span data-stu-id="33a65-114">Choose **PowerPoint Web Add-in**, then select **Next**.</span></span>

3. <span data-ttu-id="33a65-115">Назовите проект `HelloWorld` и нажмите кнопку **Создать**.</span><span class="sxs-lookup"><span data-stu-id="33a65-115">Name the project `HelloWorld`, and select **Create**.</span></span>

4. <span data-ttu-id="33a65-116">В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в PowerPoint**, а затем нажмите кнопку **Готово**, чтобы создать проект.</span><span class="sxs-lookup"><span data-stu-id="33a65-116">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="33a65-p102">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="33a65-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

     ![Руководство по PowerPoint: окно обозревателя решений Visual Studio с двумя проектами в решении HelloWorld](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="33a65-120">Обзор решения Visual Studio</span><span class="sxs-lookup"><span data-stu-id="33a65-120">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a><span data-ttu-id="33a65-121">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="33a65-121">Update code</span></span> 

<span data-ttu-id="33a65-122">Измените код надстройки, как указано ниже, чтобы создать платформу для реализации функций надстройки, следуя инструкциям в следующих разделах этого руководства.</span><span class="sxs-lookup"><span data-stu-id="33a65-122">Edit the add-in code as follows to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.</span></span>

1. <span data-ttu-id="33a65-123">Файл **Home.html** содержит HTML-контент, который будет отображаться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-123">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="33a65-124">В файле **Home.html** найдите раздел **div** с `id="content-main"`, замените весь этот раздел приведенным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="33a65-124">In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.</span></span>

    ```html
    <!-- TODO2: Create the content-header div. -->
    <div id="content-main">
        <div class="padding">
            <!-- TODO1: Create the insert-image button. -->
            <!-- TODO3: Create the insert-text button. -->
            <!-- TODO4: Create the get-slide-metadata button. -->
            <!-- TODO5: Create the go-to-slide buttons. -->
        </div>
    </div>
    ```

2. <span data-ttu-id="33a65-p104">Откройте файл **Home.js** в корневой папке проекта веб-приложения. Этот файл содержит скрипт надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="33a65-p104">Open the file **Home.js** in the root of the web application project. This file specifies the script for the add-in. Replace the entire contents with the following code and save the file.</span></span>

    ```js
    (function () {
        "use strict";

        var messageBanner;

        Office.onReady(function () {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.MessageBanner');
                messageBanner = new components.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        });

        // TODO2: Define the insertImage function. 

        // TODO3: Define the insertImageFromBase64String function.

        // TODO5: Define the insertText function.

        // TODO7: Define the getSlideMetadata function.

        // TODO9: Define the navigation functions.

        // Helper function for displaying notifications
        function showNotification(header, content) {
            $("#notification-header").text(header);
            $("#notification-body").text(content);
            messageBanner.showBanner();
            messageBanner.toggleExpansion();
        }
    })();
    ```

## <a name="insert-an-image"></a><span data-ttu-id="33a65-128">Вставка изображения</span><span class="sxs-lookup"><span data-stu-id="33a65-128">Insert an image</span></span>

<span data-ttu-id="33a65-129">Выполните указанные ниже действия, чтобы добавить код, который извлекает фотографию дня в [Bing](https://www.bing.com) и вставляет данное изображение на слайд.</span><span class="sxs-lookup"><span data-stu-id="33a65-129">Complete the following steps to add code that retrieves the [Bing](https://www.bing.com) photo of the day and inserts that image into a slide.</span></span>

1. <span data-ttu-id="33a65-130">Используя обозреватель решений, добавьте новую папку **Controllers** в проект **HelloWorldWeb**.</span><span class="sxs-lookup"><span data-stu-id="33a65-130">Using Solution Explorer, add a new folder named **Controllers** to the **HelloWorldWeb** project.</span></span>

    ![Руководство по PowerPoint: окно обозревателя решений Visual Studio с выделенной папкой Controllers в проекте HelloWorldWeb](../images/powerpoint-tutorial-solution-explorer-controllers.png)

2. <span data-ttu-id="33a65-132">Щелкните правой кнопкой мыши папку **Controllers** и выберите **Добавить > Создать шаблонный элемент**.</span><span class="sxs-lookup"><span data-stu-id="33a65-132">Right-click the **Controllers** folder and select **Add > New Scaffolded Item...**.</span></span>

3. <span data-ttu-id="33a65-133">В диалоговом окне **Добавление шаблона** выберите **Контроллер Web API 2 — пустой** и нажмите кнопку **Добавить**.</span><span class="sxs-lookup"><span data-stu-id="33a65-133">In the **Add Scaffold** dialog window, select **Web API 2 Controller - Empty** and choose the **Add** button.</span></span> 

4. <span data-ttu-id="33a65-p105">В диалоговом окне **Добавление контроллера** введите имя **PhotoController** и нажмите кнопку **Добавить**. Visual Studio создаст и откроет файл **PhotoController.cs**.</span><span class="sxs-lookup"><span data-stu-id="33a65-p105">In the **Add Controller** dialog window, enter **PhotoController** as the controller name and choose the **Add** button. Visual Studio creates and opens the **PhotoController.cs** file.</span></span>

5. <span data-ttu-id="33a65-p106">Замените все содержимое файла **PhotoController.cs** приведенным ниже кодом, который вызывает службу Bing для получения фотографии дня в виде строки в кодировке Base64. Когда для вставки изображения в документ используется API JavaScript для Office, данные изображения должны быть закодированы в формате Base64.</span><span class="sxs-lookup"><span data-stu-id="33a65-p106">Replace the entire contents of the **PhotoController.cs** file with the following code that calls the Bing service to retrieve the photo of the day as a Base64 encoded string. When you use the Office JavaScript API to insert an image into a document, the image data must be specified as a Base64 encoded string.</span></span>

    ```csharp
    using System;
    using System.IO;
    using System.Net;
    using System.Text;
    using System.Web.Http;
    using System.Xml;

    namespace HelloWorldWeb.Controllers
    {
        public class PhotoController : ApiController
        {
            public string Get()
            {
                string url = "http://www.bing.com/HPImageArchive.aspx?format=xml&idx=0&n=1";

                // Create the request.
                HttpWebRequest request = (HttpWebRequest)WebRequest.Create(url);
                WebResponse response = request.GetResponse();

                using (Stream responseStream = response.GetResponseStream())
                {
                    // Process the result.
                    StreamReader reader = new StreamReader(responseStream, Encoding.UTF8);
                    string result = reader.ReadToEnd();

                    // Parse the xml response and to get the URL.
                    XmlDocument doc = new XmlDocument();
                    doc.LoadXml(result);
                    string photoURL = "http://bing.com" + doc.SelectSingleNode("/images/image/url").InnerText;

                    // Fetch the photo and return it as a Base64 encoded string.
                    return getPhotoFromURL(photoURL);
                }
            }

            private string getPhotoFromURL(string imageURL)
            {
                var webClient = new WebClient();
                byte[] imageBytes = webClient.DownloadData(imageURL);
                return Convert.ToBase64String(imageBytes);
            }
        }
    }
    ```

6. <span data-ttu-id="33a65-p107">В файле **Home.html** замените `TODO1` приведенным ниже кодом. Этот код определяет кнопку **Insert Image** (Вставить изображение), которая появится в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-p107">In the **Home.html** file, replace `TODO1` with the following markup. This markup defines the **Insert Image** button that will appear within the add-in's task pane.</span></span>

    ```html
    <button class="Button Button--primary" id="insert-image">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Insert Image</span>
        <span class="Button-description">Gets the photo of the day that shows on the Bing home page and adds it to the slide.</span>
    </button>
    ```

7. <span data-ttu-id="33a65-140">В файле **Home.js** замените `TODO1` приведенным ниже кодом, чтобы назначить обработчик событий для кнопки **Insert Image** (Вставить изображение).</span><span class="sxs-lookup"><span data-stu-id="33a65-140">In the **Home.js** file, replace `TODO1` with the following code to assign the event handler for the **Insert Image** button.</span></span>

    ```js
    $('#insert-image').click(insertImage);
    ```

8. <span data-ttu-id="33a65-p108">В файле **Home.js** замените `TODO2` приведенным ниже кодом, чтобы определить функцию `insertImage`. Эта функция извлекает изображение из веб-службы Bing, а затем вызывает функцию `insertImageFromBase64String`, чтобы вставить его в документ.</span><span class="sxs-lookup"><span data-stu-id="33a65-p108">In the **Home.js** file, replace `TODO2` with the following code to define the `insertImage` function. This function fetches the image from the Bing web service and then calls the `insertImageFromBase64String` function to insert that image into the document.</span></span>

    ```js
    function insertImage() {
        // Get image from from web service (as a Base64 encoded string).
        $.ajax({
            url: "/api/Photo/", success: function (result) {
                insertImageFromBase64String(result);
            }, error: function (xhr, status, error) {
                showNotification("Error", "Oops, something went wrong.");
            }
        });
    }
    ```

9. <span data-ttu-id="33a65-p109">В файле **Home.js** замените `TODO3` приведенным ниже кодом, чтобы определить функцию `insertImageFromBase64String`. Эта функция использует API JavaScript для Office, чтобы вставить изображение в документ. Примечание.</span><span class="sxs-lookup"><span data-stu-id="33a65-p109">In the **Home.js** file, replace `TODO3` with the following code to define the `insertImageFromBase64String` function. This function uses the Office JavaScript API to insert the image into the document. Note:</span></span> 

    - <span data-ttu-id="33a65-146">`coercionType`, второй параметр запроса `setSelectedDataAsyc`, определяет тип вставляемых данных.</span><span class="sxs-lookup"><span data-stu-id="33a65-146">The `coercionType` option that's specified as the second parameter of the `setSelectedDataAsyc` request indicates the type of data being inserted.</span></span> 

    - <span data-ttu-id="33a65-147">Объект `asyncResult` инкапсулирует результат запроса `setSelectedDataAsync`, включая сведения о состоянии и ошибке, если запрос завершился ошибкой.</span><span class="sxs-lookup"><span data-stu-id="33a65-147">The `asyncResult` object encapsulates the result of the `setSelectedDataAsync` request, including status and error information if the request failed.</span></span>

    ```js
    function insertImageFromBase64String(image) {
        // Call Office.js to insert the image into the document.
        Office.context.document.setSelectedDataAsync(image, {
            coercionType: Office.CoercionType.Image
        },
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="33a65-148">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="33a65-148">Test the add-in</span></span>

1. <span data-ttu-id="33a65-p110">Протестируйте новую надстройку PowerPoint с помощью Visual Studio, нажав клавишу **F5** или кнопку **Запустить**, чтобы запустить PowerPoint с кнопкой надстройки **Показать область задач** на ленте. Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="33a65-p110">Using Visual Studio, test the newly created PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Запустить"](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="33a65-152">В PowerPoint нажмите кнопку **Show Taskpane** (Показать область задач) на ленте, чтобы открыть надстройку области задач.</span><span class="sxs-lookup"><span data-stu-id="33a65-152">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Show Taskpane" (Показать область задач) на ленте "Главная"](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="33a65-154">В области задач нажмите кнопку **Insert Image** (Вставить изображение), чтобы добавить фотографию дня Bing на текущий слайд.</span><span class="sxs-lookup"><span data-stu-id="33a65-154">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide.</span></span>

    ![Снимок экрана: надстройка PowerPoint с выделенной кнопкой "Insert Image" (Вставить изображение)](../images/powerpoint-tutorial-insert-image-button.png)

4. <span data-ttu-id="33a65-156">В Visual Studio остановите работу надстройки, нажав клавиши **Shift + F5** или кнопку **Остановить**.</span><span class="sxs-lookup"><span data-stu-id="33a65-156">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="33a65-157">PowerPoint автоматически закроется при остановке надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-157">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Остановить"](../images/powerpoint-tutorial-stop.png)

## <a name="customize-user-interface-ui-elements"></a><span data-ttu-id="33a65-159">Настройка элементов пользовательского интерфейса</span><span class="sxs-lookup"><span data-stu-id="33a65-159">Customize User Interface (UI) elements</span></span>

<span data-ttu-id="33a65-160">Выполните указанные ниже действия, чтобы добавить разметку, которая будет изменять область задач пользовательского интерфейса.</span><span class="sxs-lookup"><span data-stu-id="33a65-160">Complete the following steps to add markup that customizes the task pane UI.</span></span>

1. <span data-ttu-id="33a65-p112">В файле **Home.html** замените `TODO2` приведенным ниже кодом, чтобы добавить раздел верхнего колонтитула и заголовок в область задач. Примечание.</span><span class="sxs-lookup"><span data-stu-id="33a65-p112">In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane. Note:</span></span>

    - <span data-ttu-id="33a65-p113">Стили, которые начинаются с `ms-`, относятся к стилям [Office UI Fabric](../design/office-ui-fabric.md), интерфейсной платформы JavaScript для создания функциональных возможностей Office и Office 365. Файл **Home.html** включает ссылку на таблицу стилей Fabric.</span><span class="sxs-lookup"><span data-stu-id="33a65-p113">The styles that begin with `ms-` are defined by [Office UI Fabric](../design/office-ui-fabric.md), a JavaScript front-end framework for building user experiences for Office and Office 365. The **Home.html** file includes a reference to the Fabric stylesheet.</span></span>

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. <span data-ttu-id="33a65-165">В файле **Home.html** найдите раздел **div** с `class="footer"` и удалите весь раздел **div**, чтобы удалить раздел нижнего колонтитула из области задач.</span><span class="sxs-lookup"><span data-stu-id="33a65-165">In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.</span></span>

### <a name="test-the-add-in"></a><span data-ttu-id="33a65-166">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="33a65-166">Test the add-in</span></span>

1. <span data-ttu-id="33a65-167">Испытайте надстройку PowerPoint с помощью Visual Studio, нажав клавишу **F5** или кнопку **Запустить**, чтобы запустить PowerPoint с кнопкой надстройки **Показать область задач** на ленте.</span><span class="sxs-lookup"><span data-stu-id="33a65-167">Using Visual Studio, test the PowerPoint add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="33a65-168">Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="33a65-168">The add-in will be hosted locally on IIS.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Запустить"](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="33a65-170">В PowerPoint нажмите кнопку **Show Taskpane** (Показать область задач) на ленте, чтобы открыть надстройку области задач.</span><span class="sxs-lookup"><span data-stu-id="33a65-170">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Show Taskpane" (Показать область задач) на ленте "Главная"](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="33a65-172">Обратите внимание на то, что область задач теперь содержит раздел верхнего колонтитула и заголовок и больше не содержит раздел нижнего колонтитула.</span><span class="sxs-lookup"><span data-stu-id="33a65-172">Notice that the task pane now contains a header section and title, and no longer contains a footer section.</span></span>

    ![Снимок экрана: надстройка PowerPoint с кнопкой "Вставить изображение"](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. <span data-ttu-id="33a65-174">В Visual Studio остановите работу надстройки, нажав клавиши **Shift + F5** или кнопку **Остановить**.</span><span class="sxs-lookup"><span data-stu-id="33a65-174">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="33a65-175">PowerPoint автоматически закроется при остановке надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-175">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Остановить"](../images/powerpoint-tutorial-stop.png)

## <a name="insert-text"></a><span data-ttu-id="33a65-177">Вставка текста</span><span class="sxs-lookup"><span data-stu-id="33a65-177">Insert text</span></span>

<span data-ttu-id="33a65-178">Выполните указанные ниже действия, чтобы добавить код, который вставляет текст в слайд, который содержит фотографию дня из [Bing](https://www.bing.com).</span><span class="sxs-lookup"><span data-stu-id="33a65-178">Complete the following steps to add code that inserts text into the title slide which contains the [Bing](https://www.bing.com) photo of the day.</span></span>

1. <span data-ttu-id="33a65-p116">В файле **Home.html** замените `TODO3` приведенным ниже кодом. Этот код определяет кнопку **Insert Text** (Вставить текст), которая появится в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-p116">In the **Home.html** file, replace `TODO3` with the following markup. This markup defines the **Insert Text** button that will appear within the add-in's task pane.</span></span>

    ```html
        <br /><br />
        <button class="Button Button--primary" id="insert-text">
            <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
            <span class="Button-label">Insert Text</span>
            <span class="Button-description">Inserts text into the slide.</span>
        </button>
    ```

2. <span data-ttu-id="33a65-181">В файле **Home.js** замените `TODO4` приведенным ниже кодом, чтобы назначить обработчик событий для кнопки **Insert Text** (Вставить текст).</span><span class="sxs-lookup"><span data-stu-id="33a65-181">In the **Home.js** file, replace `TODO4` with the following code to assign the event handler for the **Insert Text** button.</span></span>

    ```js
    $('#insert-text').click(insertText);
    ```

3. <span data-ttu-id="33a65-p117">В файле **Home.js** замените `TODO5` на следующий код, чтобы определить функцию `insertText`. Эта функция вставляет текст в текущий слайд.</span><span class="sxs-lookup"><span data-stu-id="33a65-p117">In the **Home.js** file, replace `TODO5` with the following code to define the `insertText` function. This function inserts text into the current slide.</span></span>

    ```js
    function insertText() {
        Office.context.document.setSelectedDataAsync('Hello World!',
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="33a65-184">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="33a65-184">Test the add-in</span></span>

1. <span data-ttu-id="33a65-185">Испытайте надстройку с помощью Visual Studio, нажав клавишу **F5** или кнопку **Запустить**, чтобы запустить PowerPoint с кнопкой надстройки **Показать область задач** на ленте.</span><span class="sxs-lookup"><span data-stu-id="33a65-185">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="33a65-186">Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="33a65-186">The add-in will be hosted locally on IIS.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Запустить"](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="33a65-188">В PowerPoint нажмите кнопку **Show Taskpane** (Показать область задач) на ленте, чтобы открыть надстройку области задач.</span><span class="sxs-lookup"><span data-stu-id="33a65-188">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Show Taskpane" (Показать область задач) на ленте "Главная"](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="33a65-190">В области задач нажмите кнопку **Insert Image** (Вставить изображение), чтобы добавить фотографию дня Bing на текущий слайд, и выберите макет слайда с текстовым полем для заголовка.</span><span class="sxs-lookup"><span data-stu-id="33a65-190">In the task pane, choose the **Insert Image** button to add the Bing photo of the day to the current slide and choose a design for the slide that contains a text box for the title.</span></span>

    ![Снимок экрана: PowerPoint с выделенным текущим слайдом и надстройка с выделенной кнопкой "Вставить изображение"](../images/powerpoint-tutorial-insert-image-slide-design.png)

4. <span data-ttu-id="33a65-192">Установите курсор в текстовом поле на заглавном слайде и нажмите кнопку **Insert Text** (Вставить текст) в области задач, чтобы добавить текст.</span><span class="sxs-lookup"><span data-stu-id="33a65-192">Put your cursor in the text box on the title slide and then in the task pane, choose the **Insert Text** button to add text to the slide.</span></span>

    ![Снимок экрана: надстройка PowerPoint с выделенной кнопкой "Insert Text" (Вставить текст)](../images/powerpoint-tutorial-insert-text.png)


5. <span data-ttu-id="33a65-194">В Visual Studio остановите работу надстройки, нажав клавиши **Shift + F5** или кнопку **Остановить**.</span><span class="sxs-lookup"><span data-stu-id="33a65-194">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="33a65-195">PowerPoint автоматически закроется при остановке надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-195">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Остановить"](../images/powerpoint-tutorial-stop.png)

## <a name="get-slide-metadata"></a><span data-ttu-id="33a65-197">Получение метаданных слайда</span><span class="sxs-lookup"><span data-stu-id="33a65-197">Get slide metadata</span></span>

<span data-ttu-id="33a65-198">Выполните указанные ниже действия, чтобы добавить код, который извлекает метаданные для выбранного слайда.</span><span class="sxs-lookup"><span data-stu-id="33a65-198">Complete the following steps to add code that retrieves metadata for the selected slide.</span></span>

1. <span data-ttu-id="33a65-p120">В файле **Home.html** замените `TODO4` приведенным ниже кодом. Этот код определяет кнопку **Get Slide Metadata** (Получить метаданные слайда), которая появится в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-p120">In the **Home.html** file, replace `TODO4` with the following markup. This markup defines the **Get Slide Metadata** button that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="Button Button--primary" id="get-slide-metadata">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Get Slide Metadata</span>
        <span class="Button-description">Gets metadata for the selected slide(s).</span>
    </button>
    ```

2. <span data-ttu-id="33a65-201">В файле **Home.js** замените `TODO6` приведенным ниже кодом, чтобы назначить обработчик событий для кнопки **Get Slide Metadata** (Получить метаданные слайда).</span><span class="sxs-lookup"><span data-stu-id="33a65-201">In the **Home.js** file, replace `TODO6` with the following code to assign the event handler for the **Get Slide Metadata** button.</span></span>

    ```js
    $('#get-slide-metadata').click(getSlideMetadata);
    ```

3. <span data-ttu-id="33a65-p121">В файле **Home.js** замените `TODO7` приведенным ниже кодом, чтобы определить функцию `getSlideMetadata`. Эта функция извлекает метаданные выбранных слайдов и записывает их во всплывающее диалоговое окно в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-p121">In the **Home.js** file, replace `TODO7` with the following code to define the `getSlideMetadata` function. This function retrieves metadata for the selected slide(s) and writes it to a popup dialog window within the add-in task pane.</span></span>

    ```js
    function getSlideMetadata() {
        Office.context.document.getSelectedDataAsync(Office.CoercionType.SlideRange,
            function (asyncResult) {
                if (asyncResult.status === Office.AsyncResultStatus.Failed) {
                    showNotification("Error", asyncResult.error.message);
                } else {
                    showNotification("Metadata for selected slide(s):", JSON.stringify(asyncResult.value), null, 2);
                }
            }
        );
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="33a65-204">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="33a65-204">Test the add-in</span></span>

1. <span data-ttu-id="33a65-205">Испытайте надстройку с помощью Visual Studio, нажав клавишу **F5** или кнопку **Запустить**, чтобы запустить PowerPoint с кнопкой надстройки **Показать область задач** на ленте.</span><span class="sxs-lookup"><span data-stu-id="33a65-205">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="33a65-206">Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="33a65-206">The add-in will be hosted locally on IIS.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Запустить"](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="33a65-208">В PowerPoint нажмите кнопку **Show Taskpane** (Показать область задач) на ленте, чтобы открыть надстройку области задач.</span><span class="sxs-lookup"><span data-stu-id="33a65-208">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Show Taskpane" (Показать область задач) на ленте "Главная"](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="33a65-p123">В области задач нажмите кнопку **Get Slide Metadata** (Получить метаданные слайда), чтобы получить метаданные выбранного слайда. Метаданные слайда записываются во всплывающее диалоговое окно в нижней части области задач. В этом случае массив `slides` в метаданных JSON содержит один объект, в котором указаны свойства `id`, `title` и `index` выбранного слайда. Если при извлечении метаданных будет выбрано несколько слайдов, массив `slides` в метаданных JSON будет содержать один объект для каждого выбранного слайда.</span><span class="sxs-lookup"><span data-stu-id="33a65-p123">In the task pane, choose the **Get Slide Metadata** button to get the metadata for the selected slide. The slide metadata is written to the popup dialog window at the bottom of the task pane. In this case, the `slides` array within the JSON metadata contains one object that specifies the `id`, `title`, and `index` of the selected slide. If multiple slides had been selected when you retrieved slide metadata, the `slides` array within the JSON metadata would contain one object for each selected slide.</span></span>

    ![Снимок экрана: надстройка PowerPoint с выделенной кнопкой "Get Slide Metadata" (Получить метаданные слайда)](../images/powerpoint-tutorial-get-slide-metadata.png)

4. <span data-ttu-id="33a65-215">В Visual Studio остановите работу надстройки, нажав клавиши **Shift + F5** или кнопку **Остановить**.</span><span class="sxs-lookup"><span data-stu-id="33a65-215">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="33a65-216">PowerPoint автоматически закроется при остановке надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-216">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Остановить"](../images/powerpoint-tutorial-stop.png)

## <a name="navigate-between-slides"></a><span data-ttu-id="33a65-218">Переход между слайдами</span><span class="sxs-lookup"><span data-stu-id="33a65-218">Navigate between slides</span></span>

<span data-ttu-id="33a65-219">Выполните указанные ниже действия, чтобы добавить код, который выполняет переход между слайдами документа.</span><span class="sxs-lookup"><span data-stu-id="33a65-219">Complete the following steps to add code that navigates between the slides of a document.</span></span>

1. <span data-ttu-id="33a65-p125">В файле **Home.html** замените `TODO5` приведенным ниже кодом. Этот код определяет четыре кнопки навигации, которые появятся в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-p125">In the **Home.html** file, replace `TODO5` with the following markup. This markup defines the four navigation buttons that will appear within the add-in's task pane.</span></span>

    ```html
    <br /><br />
    <button class="Button Button--primary" id="go-to-first-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to First Slide</span>
        <span class="Button-description">Go to the first slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-next-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Next Slide</span>
        <span class="Button-description">Go to the next slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-previous-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Previous Slide</span>
        <span class="Button-description">Go to the previous slide.</span>
    </button>
    <br /><br />
    <button class="Button Button--primary" id="go-to-last-slide">
        <span class="Button-icon"><i class="ms-Icon ms-Icon--plus"></i></span>
        <span class="Button-label">Go to Last Slide</span>
        <span class="Button-description">Go to the last slide.</span>
    </button>
    ```

2. <span data-ttu-id="33a65-222">В файле **Home.js** замените `TODO8` приведенным ниже кодом, чтобы назначить обработчик событий для четырех кнопок навигации.</span><span class="sxs-lookup"><span data-stu-id="33a65-222">In the **Home.js** file, replace `TODO8` with the following code to assign the event handlers for the four navigation buttons.</span></span>

    ```js
    $('#go-to-first-slide').click(goToFirstSlide);
    $('#go-to-next-slide').click(goToNextSlide);
    $('#go-to-previous-slide').click(goToPreviousSlide);
    $('#go-to-last-slide').click(goToLastSlide);
    ```

3. <span data-ttu-id="33a65-223">В файле **Home.js** замените `TODO9` приведенным ниже кодом, чтобы определить функции навигации.</span><span class="sxs-lookup"><span data-stu-id="33a65-223">In the **Home.js** file, replace `TODO9` with the following code to define the navigation functions.</span></span> <span data-ttu-id="33a65-224">Каждая из этих функций использует функцию `goToByIdAsync` для выбора слайда с учетом его позиции в документе (первый, последний, предыдущий, следующий).</span><span class="sxs-lookup"><span data-stu-id="33a65-224">Each of these functions uses the `goToByIdAsync` function to select a slide based upon its position in the document (first, last, previous, and next).</span></span>

    ```js
    function goToFirstSlide() {
        Office.context.document.goToByIdAsync(Office.Index.First, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToLastSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Last, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToPreviousSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Previous, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }

    function goToNextSlide() {
        Office.context.document.goToByIdAsync(Office.Index.Next, Office.GoToType.Index,
            function (asyncResult) {
                if (asyncResult.status == "failed") {
                    showNotification("Error", asyncResult.error.message);
                }
            });
    }
    ```

### <a name="test-the-add-in"></a><span data-ttu-id="33a65-225">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="33a65-225">Test the add-in</span></span>

1. <span data-ttu-id="33a65-226">Испытайте надстройку с помощью Visual Studio, нажав клавишу **F5** или кнопку **Запустить**, чтобы запустить PowerPoint с кнопкой надстройки **Показать область задач** на ленте.</span><span class="sxs-lookup"><span data-stu-id="33a65-226">Using Visual Studio, test the add-in by pressing **F5** or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon.</span></span> <span data-ttu-id="33a65-227">Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="33a65-227">The add-in will be hosted locally on IIS.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Запустить"](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="33a65-229">В PowerPoint нажмите кнопку **Show Taskpane** (Показать область задач) на ленте, чтобы открыть надстройку области задач.</span><span class="sxs-lookup"><span data-stu-id="33a65-229">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Show Taskpane" (Показать область задач) на ленте "Главная"](../images/powerpoint-tutorial-show-taskpane-button.png)


3. <span data-ttu-id="33a65-231">Нажмите кнопку **Создать слайд** на ленте вкладки **Главная**, чтобы добавить в документ два новых слайда.</span><span class="sxs-lookup"><span data-stu-id="33a65-231">Use the **New Slide** button in the ribbon of the **Home** tab to add two new slides to the document.</span></span> 

4. <span data-ttu-id="33a65-p128">В области задач нажмите кнопку **Go to First Slide** (Перейти к первому слайду). Будет выбран и показан первый слайд в документе.</span><span class="sxs-lookup"><span data-stu-id="33a65-p128">In the task pane, choose the **Go to First Slide** button. The first slide in the document is selected and displayed.</span></span>

    ![Снимок экрана: надстройка PowerPoint с выделенной кнопкой "Go to First Slide" (Перейти к первому слайду)](../images/powerpoint-tutorial-go-to-first-slide.png)

5. <span data-ttu-id="33a65-p129">В области задач нажмите кнопку **Go to Next Slide** (Перейти к следующему слайду). Будет выбран и показан следующий слайд в документе.</span><span class="sxs-lookup"><span data-stu-id="33a65-p129">In the task pane, choose the **Go to Next Slide** button. The next slide in the document is selected and displayed.</span></span>

    ![Снимок экрана: надстройка PowerPoint с выделенной кнопкой "Go to Next Slide" (Перейти к следующему слайду)](../images/powerpoint-tutorial-go-to-next-slide.png)

6. <span data-ttu-id="33a65-p130">В области задач нажмите кнопку **Go to Previous Slide** (Перейти к предыдущему слайду). Будет выбран и показан предыдущий слайд в документе.</span><span class="sxs-lookup"><span data-stu-id="33a65-p130">In the task pane, choose the **Go to Previous Slide** button. The previous slide in the document is selected and displayed.</span></span>

    ![Снимок экрана: надстройка PowerPoint с выделенной кнопкой "Go to Previous Slide" (Перейти к предыдущему слайду)](../images/powerpoint-tutorial-go-to-previous-slide.png)

7. <span data-ttu-id="33a65-p131">В области задач нажмите кнопку **Go to Last Slide** (Перейти к последнему слайду). Будет выбран и показан последний слайд в документе.</span><span class="sxs-lookup"><span data-stu-id="33a65-p131">In the task pane, choose the **Go to Last Slide** button. The last slide in the document is selected and displayed.</span></span>

    ![Снимок экрана: надстройка PowerPoint с выделенной кнопкой "Go to Last Slide" (Перейти к последнему слайду)](../images/powerpoint-tutorial-go-to-last-slide.png)

8. <span data-ttu-id="33a65-244">В Visual Studio остановите работу надстройки, нажав клавиши **Shift + F5** или кнопку **Остановить**.</span><span class="sxs-lookup"><span data-stu-id="33a65-244">In Visual Studio, stop the add-in by pressing **Shift + F5** or choosing the **Stop** button.</span></span> <span data-ttu-id="33a65-245">PowerPoint автоматически закроется при остановке надстройки.</span><span class="sxs-lookup"><span data-stu-id="33a65-245">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Остановить"](../images/powerpoint-tutorial-stop.png)

## <a name="next-steps"></a><span data-ttu-id="33a65-247">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="33a65-247">Next steps</span></span>

<span data-ttu-id="33a65-248">Из этого руководства вы узнали, как создать надстройку PowerPoint, которая вставляет изображение и текст, получает метаданные слайда и выполняет переход между слайдами.</span><span class="sxs-lookup"><span data-stu-id="33a65-248">In this tutorial, you've created a PowerPoint add-in that inserts an image, inserts text, gets slide metadata, and navigates between slides.</span></span> <span data-ttu-id="33a65-249">Чтобы узнать больше о создании надстроек PowerPoint, перейдите к следующей статье:</span><span class="sxs-lookup"><span data-stu-id="33a65-249">To learn more about building PowerPoint add-ins, continue to the following article:</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="33a65-250">Обзор надстроек PowerPoint</span><span class="sxs-lookup"><span data-stu-id="33a65-250">PowerPoint add-ins overview</span></span>](../powerpoint/powerpoint-add-ins.md)

## <a name="see-also"></a><span data-ttu-id="33a65-251">См. также</span><span class="sxs-lookup"><span data-stu-id="33a65-251">See also</span></span>

* [<span data-ttu-id="33a65-252">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="33a65-252">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="33a65-253">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="33a65-253">Develop Office Add-ins</span></span>](../develop/develop-overview.md)

