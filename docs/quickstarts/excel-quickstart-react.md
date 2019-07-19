---
title: Создание области задач Excel с помощью React
description: ''
ms.date: 05/02/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: b4c7822d20985ad598d77d128fd3890963c50df3
ms.sourcegitcommit: bb44c9694f88cde32ffbb642689130db44456964
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/17/2019
ms.locfileid: "35771766"
---
# <a name="build-an-excel-task-pane-add-in-using-react"></a><span data-ttu-id="0ad11-102">Создание области задач Excel с помощью React</span><span class="sxs-lookup"><span data-stu-id="0ad11-102">Build an Excel task pane add-in using Vue</span></span>

<span data-ttu-id="0ad11-103">В этой статье описывается процесс создания надстройки в области задач Excel с помощью React и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="0ad11-103">In this article, you'll walk through the process of building an Excel task pane add-in using Vue and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="0ad11-104">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="0ad11-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="0ad11-105">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="0ad11-105">Create the add-in project</span></span>

<span data-ttu-id="0ad11-106">Создайте проект надстройки Excel помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="0ad11-106">Use the Yeoman generator to create an Excel add-in project.</span></span> <span data-ttu-id="0ad11-107">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="0ad11-107">Run the following command and then answer the prompts as follows:</span></span>

```command&nbsp;line
yo office
```

- <span data-ttu-id="0ad11-108">**Выберите тип проекта:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="0ad11-108">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="0ad11-109">**Выберите тип сценария:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="0ad11-109">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="0ad11-110">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="0ad11-110">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="0ad11-111">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="0ad11-111">**Which Office client application would you like to support?**</span></span> `Excel`

![Генератор Yeoman](../images/yo-office-excel-react-2.png)

<span data-ttu-id="0ad11-113">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="0ad11-113">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>

## <a name="explore-the-project"></a><span data-ttu-id="0ad11-114">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="0ad11-114">Explore the project</span></span>

<span data-ttu-id="0ad11-115">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="0ad11-115">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="0ad11-116">Если вы хотите ознакомиться с ключевыми компонентами проекта надстройки, откройте проект в редакторе кода и просмотрите файлы, перечисленные ниже.</span><span class="sxs-lookup"><span data-stu-id="0ad11-116">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="0ad11-117">Когда вы будете готовы попробовать собственную надстройку, перейдите к следующему разделу.</span><span class="sxs-lookup"><span data-stu-id="0ad11-117">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="0ad11-118">Файл **manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="0ad11-118">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="0ad11-119">В файле **./src/taskpane/taskpane.html** определена HTML-инфраструктура области задач, а файлы в папке **./src/taskpane/components** определяют разные части пользовательского интерфейса области задач.</span><span class="sxs-lookup"><span data-stu-id="0ad11-119">The **./src/taskpane/taskpane.html** file defines the HTML framework of the task pane, and the files within the **./src/taskpane/components** folder define the various parts of the task pane UI.</span></span>
- <span data-ttu-id="0ad11-120">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="0ad11-120">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="0ad11-121">Файл **./src/taskpane/components/App.tsx** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и Excel.</span><span class="sxs-lookup"><span data-stu-id="0ad11-121">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="0ad11-122">Проверка</span><span class="sxs-lookup"><span data-stu-id="0ad11-122">Try it out</span></span>

1. <span data-ttu-id="0ad11-123">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="0ad11-123">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="0ad11-124">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="0ad11-124">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="0ad11-126">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="0ad11-126">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="0ad11-127">Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.</span><span class="sxs-lookup"><span data-stu-id="0ad11-127">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="0ad11-129">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="0ad11-129">Next steps</span></span>

<span data-ttu-id="0ad11-130">Поздравляем! Вы успешно создали надстройку области задач Excel с помощью React!</span><span class="sxs-lookup"><span data-stu-id="0ad11-130">Congratulations, you've successfully created an Excel task pane add-in using Vue!</span></span> <span data-ttu-id="0ad11-131">Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="0ad11-131">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="0ad11-132">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="0ad11-132">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="0ad11-133">См. также</span><span class="sxs-lookup"><span data-stu-id="0ad11-133">See also</span></span>

* [<span data-ttu-id="0ad11-134">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="0ad11-134">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial-create-table.md)
* [<span data-ttu-id="0ad11-135">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="0ad11-135">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="0ad11-136">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="0ad11-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="0ad11-137">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="0ad11-137">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)
