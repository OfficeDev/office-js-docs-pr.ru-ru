---
title: Создание области задач Excel с помощью Angular
description: Узнайте, как создать простую надстройку области задач Excel, используя API JS для Office и Angular.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: d843a74b3542df8dbc462ae2876179de7b42a2d2
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076933"
---
# <a name="build-an-excel-task-pane-add-in-using-angular"></a><span data-ttu-id="d5dff-103">Создание области задач Excel с помощью Angular</span><span class="sxs-lookup"><span data-stu-id="d5dff-103">Build an Excel task pane add-in using Angular</span></span>

<span data-ttu-id="d5dff-104">Из этой статьи вы узнаете, как создать надстройку области Excel, используя Angular и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="d5dff-104">In this article, you'll walk through the process of building an Excel task pane add-in using Angular and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="d5dff-105">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="d5dff-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="d5dff-106">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="d5dff-106">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="d5dff-107">**Выберите тип проекта:** `Office Add-in Task Pane project using Angular framework`</span><span class="sxs-lookup"><span data-stu-id="d5dff-107">**Choose a project type:** `Office Add-in Task Pane project using Angular framework`</span></span>
- <span data-ttu-id="d5dff-108">**Выберите тип сценария:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="d5dff-108">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="d5dff-109">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="d5dff-109">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="d5dff-110">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="d5dff-110">**Which Office client application would you like to support?**</span></span> `Excel`

![Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office, где в качестве типа проекта установлена инфраструктура Angular.](../images/yo-office-excel-angular-2.png)

<span data-ttu-id="d5dff-112">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="d5dff-112">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="d5dff-113">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="d5dff-113">Explore the project</span></span>

<span data-ttu-id="d5dff-114">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="d5dff-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="d5dff-115">Если вы хотите ознакомиться с ключевыми компонентами проекта надстройки, откройте проект в редакторе кода и просмотрите файлы, перечисленные ниже.</span><span class="sxs-lookup"><span data-stu-id="d5dff-115">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="d5dff-116">Когда вы будете готовы попробовать собственную надстройку, перейдите к следующему разделу.</span><span class="sxs-lookup"><span data-stu-id="d5dff-116">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="d5dff-117">Файл **manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="d5dff-117">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="d5dff-118">Файл **./src/taskpane/app/app.component.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="d5dff-118">The **./src/taskpane/app/app.component.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="d5dff-119">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="d5dff-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="d5dff-120">Файл **./src/taskpane/app/app.component.ts** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и Excel.</span><span class="sxs-lookup"><span data-stu-id="d5dff-120">The **./src/taskpane/app/app.component.ts** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="d5dff-121">Проверка</span><span class="sxs-lookup"><span data-stu-id="d5dff-121">Try it out</span></span>

1. <span data-ttu-id="d5dff-122">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="d5dff-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="d5dff-123">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="d5dff-123">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: меню "Главная" в Excel с выделенной кнопкой "Показать область задач".](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="d5dff-125">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="d5dff-125">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="d5dff-126">Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.</span><span class="sxs-lookup"><span data-stu-id="d5dff-126">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Снимок экрана: Excel с открытой областью задач надстройки и выделенной кнопкой "Выполнить".](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="d5dff-128">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="d5dff-128">Next steps</span></span>

<span data-ttu-id="d5dff-p102">Поздравляем, вы успешно создали надстройку панели задач Excel с помощью Angular! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="d5dff-p102">Congratulations, you've successfully created an Excel task pane add-in using Angular! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="d5dff-131">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="d5dff-131">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="d5dff-132">См. также</span><span class="sxs-lookup"><span data-stu-id="d5dff-132">See also</span></span>

* [<span data-ttu-id="d5dff-133">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d5dff-133">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="d5dff-134">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="d5dff-134">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
* [<span data-ttu-id="d5dff-135">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="d5dff-135">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="d5dff-136">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="d5dff-136">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="d5dff-137">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="d5dff-137">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)
