---
title: Создание области задач Excel с помощью React
description: Узнайте, как создать простую надстройку области задач Excel, используя API JS для Office и React.
ms.date: 11/09/2020
ms.prod: excel
localization_priority: Priority
ms.openlocfilehash: 4cb3c56af21f11efcb97fd9fe901a2d0718ae801
ms.sourcegitcommit: ee9e92a968e4ad23f1e371f00d4888e4203ab772
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/23/2021
ms.locfileid: "53076912"
---
# <a name="build-an-excel-task-pane-add-in-using-react"></a><span data-ttu-id="feddc-103">Создание области задач Excel с помощью React</span><span class="sxs-lookup"><span data-stu-id="feddc-103">Build an Excel task pane add-in using React</span></span>

<span data-ttu-id="feddc-104">В этой статье описывается процесс создания надстройки в области задач Excel с помощью React и API JavaScript для Excel.</span><span class="sxs-lookup"><span data-stu-id="feddc-104">In this article, you'll walk through the process of building an Excel task pane add-in using React and the Excel JavaScript API.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="feddc-105">Необходимые условия</span><span class="sxs-lookup"><span data-stu-id="feddc-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="feddc-106">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="feddc-106">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="feddc-107">**Выберите тип проекта:** `Office Add-in Task Pane project using React framework`</span><span class="sxs-lookup"><span data-stu-id="feddc-107">**Choose a project type:** `Office Add-in Task Pane project using React framework`</span></span>
- <span data-ttu-id="feddc-108">**Выберите тип сценария:** `TypeScript`</span><span class="sxs-lookup"><span data-stu-id="feddc-108">**Choose a script type:** `TypeScript`</span></span>
- <span data-ttu-id="feddc-109">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="feddc-109">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="feddc-110">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="feddc-110">**Which Office client application would you like to support?**</span></span> `Excel`

![Снимок экрана: интерфейс командной строки генератора Yeoman надстроек Office, где в качестве типа проекта установлена инфраструктура React.](../images/yo-office-excel-react-2.png)

<span data-ttu-id="feddc-112">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="feddc-112">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="feddc-113">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="feddc-113">Explore the project</span></span>

<span data-ttu-id="feddc-114">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="feddc-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> <span data-ttu-id="feddc-115">Если вы хотите ознакомиться с ключевыми компонентами проекта надстройки, откройте проект в редакторе кода и просмотрите файлы, перечисленные ниже.</span><span class="sxs-lookup"><span data-stu-id="feddc-115">If you'd like to explore the key components of your add-in project, open the project in your code editor and review the files listed below.</span></span> <span data-ttu-id="feddc-116">Когда вы будете готовы попробовать собственную надстройку, перейдите к следующему разделу.</span><span class="sxs-lookup"><span data-stu-id="feddc-116">When you're ready to try out your add-in, proceed to the next section.</span></span>

- <span data-ttu-id="feddc-117">Файл **manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="feddc-117">The **manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="feddc-118">В файле **./src/taskpane/taskpane.html** определена HTML-инфраструктура области задач, а файлы в папке **./src/taskpane/components** определяют разные части пользовательского интерфейса области задач.</span><span class="sxs-lookup"><span data-stu-id="feddc-118">The **./src/taskpane/taskpane.html** file defines the HTML framework of the task pane, and the files within the **./src/taskpane/components** folder define the various parts of the task pane UI.</span></span>
- <span data-ttu-id="feddc-119">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="feddc-119">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="feddc-120">Файл **./src/taskpane/components/App.tsx** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и Excel.</span><span class="sxs-lookup"><span data-stu-id="feddc-120">The **./src/taskpane/components/App.tsx** file contains the Office JavaScript API code that facilitates interaction between the task pane and Excel.</span></span>

## <a name="try-it-out"></a><span data-ttu-id="feddc-121">Проверка</span><span class="sxs-lookup"><span data-stu-id="feddc-121">Try it out</span></span>

1. <span data-ttu-id="feddc-122">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="feddc-122">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. [!include[Start server section](../includes/quickstart-yo-start-server-excel.md)] 

3. <span data-ttu-id="feddc-123">В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="feddc-123">In Excel, choose the **Home** tab, and then choose the **Show Taskpane** button in the ribbon to open the add-in task pane.</span></span>

    ![Снимок экрана: меню "Главная" в Excel с выделенной кнопкой "Показать область задач".](../images/excel-quickstart-addin-3b.png)

4. <span data-ttu-id="feddc-125">Выберите любой диапазон ячеек на листе.</span><span class="sxs-lookup"><span data-stu-id="feddc-125">Select any range of cells in the worksheet.</span></span>

5. <span data-ttu-id="feddc-126">Внизу области задач выберите ссылку **Выполнить**, чтобы задать выбранному диапазону желтый цвет.</span><span class="sxs-lookup"><span data-stu-id="feddc-126">At the bottom of the task pane, choose the **Run** link to set the color of the selected range to yellow.</span></span>

    ![Снимок экрана: Excel с открытой областью задач надстройки и выделенной кнопкой "Запустить".](../images/excel-quickstart-addin-3c.png)

## <a name="next-steps"></a><span data-ttu-id="feddc-128">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="feddc-128">Next steps</span></span>

<span data-ttu-id="feddc-p102">Поздравляем, вы успешно создали надстройку панели задач Excel с помощью React! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="feddc-p102">Congratulations, you've successfully created an Excel task pane add-in using React! Next, learn more about the capabilities of an Excel add-in and build a more complex add-in by following along with the Excel add-in tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="feddc-131">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="feddc-131">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)

## <a name="see-also"></a><span data-ttu-id="feddc-132">См. также</span><span class="sxs-lookup"><span data-stu-id="feddc-132">See also</span></span>

* [<span data-ttu-id="feddc-133">Руководство по надстройкам Excel</span><span class="sxs-lookup"><span data-stu-id="feddc-133">Excel add-in tutorial</span></span>](../tutorials/excel-tutorial.md)
* [<span data-ttu-id="feddc-134">Объектная модель JavaScript для Excel в надстройках Office</span><span class="sxs-lookup"><span data-stu-id="feddc-134">Excel JavaScript object model in Office Add-ins</span></span>](../excel/excel-add-ins-core-concepts.md)
* [<span data-ttu-id="feddc-135">Примеры кода надстроек Excel</span><span class="sxs-lookup"><span data-stu-id="feddc-135">Excel add-in code samples</span></span>](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [<span data-ttu-id="feddc-136">Справочник по API JavaScript для Excel</span><span class="sxs-lookup"><span data-stu-id="feddc-136">Excel JavaScript API reference</span></span>](../reference/overview/excel-add-ins-reference-overview.md)