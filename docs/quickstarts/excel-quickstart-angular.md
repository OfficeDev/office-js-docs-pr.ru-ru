---
title: Создание области задач Excel с помощью Angular
description: ''
ms.date: 01/16/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 6fd28f3b572b7bca27d5bf08b9799333285064bc
ms.sourcegitcommit: 8bce9c94540ed484d0749f07123dc7c72a6ca126
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/22/2020
ms.locfileid: "41265580"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="444e4-102">Создание области задач Excel с помощью Angular</span><span class="sxs-lookup"><span data-stu-id="444e4-102">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="444e4-103">Из этой статьи вы узнаете, как создать надстройку области Excel, используя Angular и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="444e4-103">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="444e4-104">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="444e4-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="444e4-105">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="444e4-105">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="444e4-106">**Выберите тип проекта:** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="444e4-106">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
- <span data-ttu-id="444e4-107">**Выберите тип сценария:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="444e4-107">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="444e4-108">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="444e4-108">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="444e4-109">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="444e4-109">**Which Office client application would you like to support?**</span></span> `Excel`

![Генератор Yeoman](../images/yo-office-excel-angular-2.png)

<span data-ttu-id="444e4-111">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="444e4-111">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="444e4-112">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="444e4-112">Explore the project</span></span>

<span data-ttu-id="444e4-113">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="444e4-113">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="444e4-114">Если вы хотите ознакомиться с ключевыми компонентами проекта надстройки, откройте проект в редакторе кода и просмотрите файлы, перечисленные ниже.</span><span class="sxs-lookup"><span data-stu-id="444e4-114">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="444e4-115">Когда вы будете готовы попробовать собственную надстройку, перейдите к следующему разделу.</span><span class="sxs-lookup"><span data-stu-id="444e4-115">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="444e4-116">Файл **manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="444e4-116">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="444e4-117">Файл **./src/taskpane/app/app.component.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="444e4-117">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="444e4-118">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="444e4-118">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="444e4-119">Файл **./src/taskpane/app/app.component.ts** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и Excel.</span><span class="sxs-lookup"><span data-stu-id="444e4-119">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="444e4-120">Проверка</span><span class="sxs-lookup"><span data-stu-id="444e4-120">Try it out</span></span>

1. <span data-ttu-id="444e4-121">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="444e4-121">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="444e4-122">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="444e4-122">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="444e4-124">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="444e4-124">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="444e4-125">Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.</span><span class="sxs-lookup"><span data-stu-id="444e4-125">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="444e4-127">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="444e4-127">Next steps</span></span>

<span data-ttu-id="444e4-128">Поздравляем! Вы успешно создали надстройку области задач Excel с помощью Angular!</span><span class="sxs-lookup"><span data-stu-id="444e4-128">Congratulations, you've successfully created an Excel task pane add-in using Angular!</span></span> <span data-ttu-id="444e4-129">Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="444e4-129">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="444e4-130">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="444e4-130">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="444e4-131">См. также</span><span class="sxs-lookup"><span data-stu-id="444e4-131">See also</span></span>

* [<span data-ttu-id="444e4-132">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="444e4-132">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="444e4-133">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="444e4-133">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="444e4-134">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="444e4-134">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="444e4-135">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="444e4-135">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="444e4-136">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="444e4-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="444e4-137">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="444e4-137">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)