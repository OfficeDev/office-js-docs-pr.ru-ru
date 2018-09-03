<span data-ttu-id="ab0bf-101">Сначала необходимо настроить проект разработки.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-101">You'll begin this tutorial by setting up your development project.</span></span> 

> [!NOTE]
> <span data-ttu-id="ab0bf-102">Это один из разделов руководства по надстройкам PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="ab0bf-103">Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам PowerPoint](../tutorials/powerpoint-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="ab0bf-104">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="ab0bf-104">Prerequisites</span></span>

[!include[Quick Start prerequisites](../includes/quickstart-vs-prerequisites.md)]

## <a name="setup"></a><span data-ttu-id="ab0bf-105">Установка</span><span class="sxs-lookup"><span data-stu-id="ab0bf-105">Setup</span></span>

<span data-ttu-id="ab0bf-106">Из этого руководства вы узнаете, как создать надстройку, используя Visual Studio.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-106">In this tutorial, you'll create an add-in using Visual Studio.</span></span>

### <a name="create-the-add-in-project"></a><span data-ttu-id="ab0bf-107">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="ab0bf-107">Create the add-in project</span></span>

1. <span data-ttu-id="ab0bf-108">В строке меню Visual Studio выберите **Файл** > **Создать** > **Проект**.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-108">On the Visual Studio menu bar, choose  **File** > **New** > **Project**.</span></span>
    
2. <span data-ttu-id="ab0bf-109">В списке типов проекта разверните узел **Visual C#** или **Visual Basic**, разверните **Office/SharePoint**, затем выберите **Надстройки** > **Веб-надстройка PowerPoint**.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-109">In the list of project types under **Visual C#** or **Visual Basic**, expand  **Office/SharePoint**, choose **Add-ins**, and then choose **PowerPoint Web Add-in** as the project type.</span></span> 

3. <span data-ttu-id="ab0bf-110">Назовите проект **HelloWorld** и нажмите кнопку **ОК**.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-110">Name the project **HelloWorld**, and then choose the **OK** button.</span></span>

4. <span data-ttu-id="ab0bf-111">В диалоговом окне **Создание надстройки Office** выберите **Добавить новые функции в PowerPoint** и нажмите кнопку **Готово**, чтобы создать проект.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-111">In the **Create Office Add-in** dialog window, choose **Add new functionalities to PowerPoint**, and then choose **Finish** to create the project.</span></span>

5. <span data-ttu-id="ab0bf-p102">Visual Studio создаст решение, и в **обозревателе решений** появятся два соответствующих проекта. В Visual Studio откроется файл **Home.html**.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-p102">Visual Studio creates a solution and its two projects appear in **Solution Explorer**. The **Home.html** file opens in Visual Studio.</span></span>

     ![Руководство по PowerPoint: окно обозревателя решений Visual Studio с двумя проектами в решении HelloWorld](../images/powerpoint-tutorial-solution-explorer.png)

### <a name="explore-the-visual-studio-solution"></a><span data-ttu-id="ab0bf-115">Обзор решения Visual Studio</span><span class="sxs-lookup"><span data-stu-id="ab0bf-115">Explore the Visual Studio solution</span></span>

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-code"></a><span data-ttu-id="ab0bf-116">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="ab0bf-116">Update code</span></span> 

<span data-ttu-id="ab0bf-117">Измените код надстройки, как указано ниже, чтобы создать платформу для реализации функций надстройки, следуя инструкциям в следующих разделах этого руководства.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-117">Edit the add-in code as follows, to create the framework that you'll use to implement add-in functionality in subsequent steps of this tutorial.</span></span>

1. <span data-ttu-id="ab0bf-118">Файл **Home.html** содержит HTML-контент, который будет отрисовываться в области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-118">**Home.html** specifies the HTML that will be rendered in the add-in's task pane.</span></span> <span data-ttu-id="ab0bf-119">В файле **Home.html** найдите раздел **div** с `id="content-main"`, замените весь этот раздел приведенным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-119">In **Home.html**, find the **div** with `id="content-main"`, replace that entire **div** with the following markup, and save the file.</span></span>

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

2. <span data-ttu-id="ab0bf-120">Откройте файл **Home.js** в корневой папке проекта веб-приложения.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-120">Open the file **Home.js** in the root of the web application project.</span></span> <span data-ttu-id="ab0bf-121">Этот файл содержит скрипт надстройки.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-121">This file specifies the script for the add-in.</span></span> <span data-ttu-id="ab0bf-122">Замените все его содержимое указанным ниже кодом и сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="ab0bf-122">Replace the entire contents with the following code and save the file.</span></span>

    ```javascript
    (function () {
        "use strict";

        var messageBanner;

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Initialize the FabricUI notification mechanism and hide it
                var element = document.querySelector('.ms-MessageBanner');
                messageBanner = new fabric.MessageBanner(element);
                messageBanner.hideBanner();

                // TODO1: Assign event handler for insert-image button.
                // TODO4: Assign event handler for insert-text button.
                // TODO6: Assign event handler for get-slide-metadata button.
                // TODO8: Assign event handlers for the four navigation buttons.
            });
        };

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
