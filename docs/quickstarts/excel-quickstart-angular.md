---
title: Создание области задач Excel с помощью Angular
description: ''
ms.date: 12/24/2019
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: c2565c305833335e589a53655397bcb68654d22a
ms.sourcegitcommit: 350f5c6954dec3e9384e2030cd3265aaba7ae904
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 12/23/2019
ms.locfileid: "40851560"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="a0240-102">Создание области задач Excel с помощью Angular</span><span class="sxs-lookup"><span data-stu-id="a0240-102">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="a0240-103">Из этой статьи вы узнаете, как создать надстройку области Excel, используя Angular и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="a0240-103">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="a0240-104">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="a0240-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="a0240-105">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="a0240-105">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="a0240-106">**Выберите тип проекта:** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="a0240-106">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
- <span data-ttu-id="a0240-107">**Выберите тип сценария:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="a0240-107">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="a0240-108">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="a0240-108">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="a0240-109">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="a0240-109">**Which Office client application would you like to support?**</span></span> `Excel`

![Генератор Yeoman](../images/yo-office-excel-angular-2.png)

<span data-ttu-id="a0240-111">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="a0240-111">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="a0240-112">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="a0240-112">Explore the project</span></span>

<span data-ttu-id="a0240-113">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="a0240-113">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="a0240-114">Если вы хотите ознакомиться с ключевыми компонентами проекта надстройки, откройте проект в редакторе кода и просмотрите файлы, перечисленные ниже.</span><span class="sxs-lookup"><span data-stu-id="a0240-114">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="a0240-115">Когда вы будете готовы попробовать собственную надстройку, перейдите к следующему разделу.</span><span class="sxs-lookup"><span data-stu-id="a0240-115">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="a0240-116">Файл **manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="a0240-116">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="a0240-117">Файл **./src/taskpane/app/app.component.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="a0240-117">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="a0240-118">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="a0240-118">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="a0240-119">Файл **./src/taskpane/app/app.component.ts** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и Excel.</span><span class="sxs-lookup"><span data-stu-id="a0240-119">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="a0240-120">Проверка</span><span class="sxs-lookup"><span data-stu-id="a0240-120">Try it out</span></span>

1. <span data-ttu-id="a0240-121">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="a0240-121">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="a0240-122">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="a0240-122">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="a0240-124">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="a0240-124">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="a0240-125">Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.</span><span class="sxs-lookup"><span data-stu-id="a0240-125">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Надстройка Excel](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="a0240-127">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="a0240-127">Next steps</span></span>

<span data-ttu-id="a0240-128">Поздравляем! Вы успешно создали надстройку области задач Excel с помощью Angular!</span><span class="sxs-lookup"><span data-stu-id="a0240-128">Congratulations, you've successfully created an Excel task pane add-in using Angular!</span></span> <span data-ttu-id="a0240-129">Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="a0240-129">Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="a0240-130">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="a0240-130">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="a0240-131">См. также</span><span class="sxs-lookup"><span data-stu-id="a0240-131">See also</span></span>

* [<span data-ttu-id="a0240-132">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a0240-132">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="a0240-133">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="a0240-133">Building Office Add-ins using Office.js book</span></span>](../overview/office-add-ins-fundamentals.md)
* <span data-ttu-id="a0240-134">[Разработка надстроек Office](../develop/develop-overview.md)</span><span class="sxs-lookup"><span data-stu-id="a0240-134">[](../develop/develop-overview.md)Develop Office Add-ins with Angular</span></span>
* [<span data-ttu-id="a0240-135">Основные концепции программирования с помощью API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a0240-135">Fundamental programming concepts with the Excel JavaScript API</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="a0240-136">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="a0240-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="a0240-137">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="a0240-137">Excel JavaScript API reference</span></span>](/office/dev/add-ins/reference/overview/excel-add-ins-reference-overview)