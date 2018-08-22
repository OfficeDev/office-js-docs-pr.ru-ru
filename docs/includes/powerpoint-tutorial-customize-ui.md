<span data-ttu-id="491d6-101">В этом разделе руководства описывается, как настроить пользовательский интерфейс области задач.</span><span class="sxs-lookup"><span data-stu-id="491d6-101">In this step of the tutorial, you'll customize the task pane user interface (UI).</span></span>

> [!NOTE]
> <span data-ttu-id="491d6-102">Это один из разделов руководства по надстройкам PowerPoint.</span><span class="sxs-lookup"><span data-stu-id="491d6-102">This page describes an individual step of the PowerPoint add-in tutorial.</span></span> <span data-ttu-id="491d6-103">Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам PowerPoint](../tutorials/powerpoint-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="491d6-103">If you’ve arrived at this page via search engine results or other direct link, please go to the [PowerPoint add-in tutorial](../tutorials/powerpoint-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="customize-the-task-pane-ui"></a><span data-ttu-id="491d6-104">Настройка пользовательского интерфейса области задач</span><span class="sxs-lookup"><span data-stu-id="491d6-104">Customize the task pane UI</span></span> 

1. <span data-ttu-id="491d6-105">В файле **Home.html** замените `TODO2` приведенным ниже кодом, чтобы добавить раздел верхнего колонтитула и заголовок в область задач.</span><span class="sxs-lookup"><span data-stu-id="491d6-105">In the **Home.html** file, replace `TODO2` with the following markup to add a header section and title to the task pane.</span></span> <span data-ttu-id="491d6-106">Примечание.</span><span class="sxs-lookup"><span data-stu-id="491d6-106">Note:</span></span>

    - <span data-ttu-id="491d6-107">Стили, которые начинаются с `ms-`, относятся к стилям [Office UI Fabric](../design/office-ui-fabric.md), интерфейсной платформы JavaScript для создания функциональных возможностей Office и Office 365.</span><span class="sxs-lookup"><span data-stu-id="491d6-107">The styles that begin with `ms-` are defined by [Office UI Fabric](../design/office-ui-fabric.md), a JavaScript front-end framework for building user experiences for Office and Office 365.</span></span> <span data-ttu-id="491d6-108">Файл **Home.html** включает ссылку на таблицу стилей Fabric.</span><span class="sxs-lookup"><span data-stu-id="491d6-108">The **Home.html** file includes a reference to the Fabric stylesheet.</span></span>

    ```html
    <div id="content-header">
        <div class="ms-Grid ms-bgColor-neutralPrimary">
            <div class="ms-Grid-row">
                <div class="padding ms-Grid-col ms-u-sm12 ms-u-md12 ms-u-lg12"> <div class="ms-font-xl ms-fontColor-white ms-fontWeight-semibold">My PowerPoint Add-in</div></div>
            </div>
        </div>
    </div>
    ```

2. <span data-ttu-id="491d6-109">В файле **Home.html** найдите раздел **div** с `class="footer"` и удалите весь раздел **div**, чтобы удалить раздел нижнего колонтитула из области задач.</span><span class="sxs-lookup"><span data-stu-id="491d6-109">In the **Home.html** file, find the **div** with `class="footer"` and delete that entire **div** to remove the footer section from the task pane.</span></span>

## <a name="test-the-add-in"></a><span data-ttu-id="491d6-110">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="491d6-110">Test the add-in</span></span>

1. <span data-ttu-id="491d6-p104">Протестируйте надстройку PowerPoint с помощью Visual Studio, нажав клавишу `F5` или кнопку **Запустить**, чтобы запустить PowerPoint с кнопкой надстройки **Show Taskpane** (Показать область задач) на ленте. Надстройка будет размещена на локальном сервере IIS.</span><span class="sxs-lookup"><span data-stu-id="491d6-p104">Using Visual Studio, test the PowerPoint add-in by pressing `F5` or choosing the **Start** button to launch PowerPoint with the **Show Taskpane** add-in button displayed in the ribbon. The add-in will be hosted locally on IIS.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Запустить"](../images/powerpoint-tutorial-start.png)

2. <span data-ttu-id="491d6-114">В PowerPoint нажмите кнопку **Show Taskpane** (Показать область задач) на ленте, чтобы открыть надстройку области задач.</span><span class="sxs-lookup"><span data-stu-id="491d6-114">In PowerPoint, select the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Show Taskpane" (Показать область задач) на ленте "Главная"](../images/powerpoint-tutorial-show-taskpane-button.png)

3. <span data-ttu-id="491d6-116">Обратите внимание на то, что область задач теперь содержит раздел верхнего колонтитула и заголовок и больше не содержит раздел нижнего колонтитула.</span><span class="sxs-lookup"><span data-stu-id="491d6-116">Notice that the task pane now contains a header section and title, and no longer contains a footer section.</span></span>

    ![Снимок экрана: надстройка PowerPoint с выделенной кнопкой "Insert Image" (Вставить изображение)](../images/powerpoint-tutorial-new-task-pane-ui.png)

4. <span data-ttu-id="491d6-118">В Visual Studio остановите работу надстройки, нажав клавиши `Shift + F5` или кнопку **Остановить**.</span><span class="sxs-lookup"><span data-stu-id="491d6-118">In Visual Studio, stop the add-in by pressing `Shift + F5` or choosing the **Stop** button.</span></span> <span data-ttu-id="491d6-119">PowerPoint автоматически закроется.</span><span class="sxs-lookup"><span data-stu-id="491d6-119">PowerPoint will automatically close when the add-in is stopped.</span></span>

    ![Снимок экрана: Visual Studio с выделенной кнопкой "Остановить"](../images/powerpoint-tutorial-stop.png)

