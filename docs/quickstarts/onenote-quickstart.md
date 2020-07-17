---
title: Создание первой надстройки области задач OneNote
description: Узнайте, как создать простую надстройку для области задач OneNote, используя API JS для Office.
ms.date: 07/07/2020
ms.prod: onenote
localization_priority: Priority
ms.openlocfilehash: 6f057d650451d12e834d8f875f40d9d6d71ee4d7
ms.sourcegitcommit: 7ef14753dce598a5804dad8802df7aaafe046da7
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/10/2020
ms.locfileid: "45094157"
---
# <a name="build-your-first-onenote-task-pane-add-in"></a><span data-ttu-id="46a95-103">Создание первой надстройки области задач OneNote</span><span class="sxs-lookup"><span data-stu-id="46a95-103">Build your first OneNote task pane add-in</span></span>

<span data-ttu-id="46a95-104">В этой статье вы ознакомитесь с процессом создания надстройки для области задач OneNote.</span><span class="sxs-lookup"><span data-stu-id="46a95-104">In this article, you'll walk through the process of building a OneNote task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="46a95-105">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="46a95-105">Prerequisites</span></span>

[!include[Set up requirements](../includes/set-up-dev-environment-beforehand.md)]
[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

## <a name="create-the-add-in-project"></a><span data-ttu-id="46a95-106">Создание проекта надстройки</span><span class="sxs-lookup"><span data-stu-id="46a95-106">Create the add-in project</span></span>

[!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

- <span data-ttu-id="46a95-107">**Выберите тип проекта:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="46a95-107">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
- <span data-ttu-id="46a95-108">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="46a95-108">**Choose a script type:** `Javascript`</span></span>
- <span data-ttu-id="46a95-109">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="46a95-109">**What do you want to name your add-in?**</span></span> `My Office Add-in`
- <span data-ttu-id="46a95-110">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="46a95-110">**Which Office client application would you like to support?**</span></span> `OneNote`

![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-onenote.png)

<span data-ttu-id="46a95-112">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="46a95-112">After you complete the wizard, the generator creates the project and installs supporting Node components.</span></span>

[!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

## <a name="explore-the-project"></a><span data-ttu-id="46a95-113">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="46a95-113">Explore the project</span></span>

<span data-ttu-id="46a95-114">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="46a95-114">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="46a95-115">Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="46a95-115">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="46a95-116">Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="46a95-116">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="46a95-117">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="46a95-117">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="46a95-118">Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и ведущим приложением Office.</span><span class="sxs-lookup"><span data-stu-id="46a95-118">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="46a95-119">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="46a95-119">Update the code</span></span>

<span data-ttu-id="46a95-120">Откройте файл **./src/taskpane/taskpane.js** в редакторе кода и добавьте следующий код в функцию `run`.</span><span class="sxs-lookup"><span data-stu-id="46a95-120">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the `run` function.</span></span> <span data-ttu-id="46a95-121">В этом коде используется API JavaScript для OneNote, чтобы настроить заголовок страницы и добавить контур к тексту страницы.</span><span class="sxs-lookup"><span data-stu-id="46a95-121">This code uses the OneNote JavaScript API to set the page title and add an outline to the body of the page.</span></span>

```js
try {
    await OneNote.run(async context => {

        // Get the current page.
        var page = context.application.getActivePage();

        // Queue a command to set the page title.
        page.title = "Hello World";

        // Queue a command to add an outline to the page.
        var html = "<p><ol><li>Item #1</li><li>Item #2</li></ol></p>";
        page.addOutline(40, 90, html);

        // Run the queued commands, and return a promise to indicate task completion.
        return context.sync();
    });
} catch (error) {
    console.log("Error: " + error);
}
```

## <a name="try-it-out"></a><span data-ttu-id="46a95-122">Проверка</span><span class="sxs-lookup"><span data-stu-id="46a95-122">Try it out</span></span>

1. <span data-ttu-id="46a95-123">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="46a95-123">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

2. <span data-ttu-id="46a95-124">Запустите локальный веб-сервер и загрузите неопубликованную надстройку.</span><span class="sxs-lookup"><span data-stu-id="46a95-124">Start the local web server and sideload your add-in.</span></span>

    > [!NOTE]
    > <span data-ttu-id="46a95-125">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="46a95-125">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="46a95-126">Если вам будет предложено установить сертификат после того, как вы запустите одну из указанных ниже команд, примите предложение установить сертификат, предоставленный генератором Yeoman.</span><span class="sxs-lookup"><span data-stu-id="46a95-126">If you are prompted to install a certificate after you run one of the following commands, accept the prompt to install the certificate that the Yeoman generator provides.</span></span>

    > [!TIP]
    > <span data-ttu-id="46a95-127">Если вы тестируете надстройку на компьютере Mac, перед продолжением выполните указанную ниже команду.</span><span class="sxs-lookup"><span data-stu-id="46a95-127">If you're testing your add-in on Mac, run the following command before proceeding.</span></span> <span data-ttu-id="46a95-128">После выполнения этой команды запустится локальный веб-сервер.</span><span class="sxs-lookup"><span data-stu-id="46a95-128">When you run this command, the local web server starts.</span></span>
    >
    > ```command&nbsp;line
    > npm run dev-server
    > ```

    <span data-ttu-id="46a95-129">Выполните указанную ниже команду в корневом каталоге своего проекта.</span><span class="sxs-lookup"><span data-stu-id="46a95-129">Run the following command in the root directory of your project.</span></span> <span data-ttu-id="46a95-130">После выполнения этой команды запустится локальный веб-сервер (если он еще не запущен).</span><span class="sxs-lookup"><span data-stu-id="46a95-130">When you run this command, the local web server will start (if it's not already running).</span></span>

    ```command&nbsp;line
    npm run start:web
    ```

3. <span data-ttu-id="46a95-131">Откройте записную книжку в [OneNote в Интернете](https://www.onenote.com/notebooks) и создайте страницу.</span><span class="sxs-lookup"><span data-stu-id="46a95-131">In [OneNote on the web](https://www.onenote.com/notebooks), open a notebook and create a new page.</span></span>

4. <span data-ttu-id="46a95-132">Выберите **Вставка > Надстройки Office**. Откроется диалоговое окно "Надстройки Office".</span><span class="sxs-lookup"><span data-stu-id="46a95-132">Choose **Insert > Office Add-ins** to open the Office Add-ins dialog.</span></span>

    - <span data-ttu-id="46a95-133">Если вы вошли с помощью обычной учетной записи, выберите **Отправить надстройку** на вкладке **МОИ НАДСТРОЙКИ**.</span><span class="sxs-lookup"><span data-stu-id="46a95-133">If you're signed in with your consumer account, select the **MY ADD-INS** tab, and then choose **Upload My Add-in**.</span></span>

    - <span data-ttu-id="46a95-134">Если вы вошли с помощью рабочей или учебной учетной записи, выберите **Отправить надстройку** на вкладке **МОЯ ОРГАНИЗАЦИЯ**.</span><span class="sxs-lookup"><span data-stu-id="46a95-134">If you're signed in with your work or education account, select the **MY ORGANIZATION** tab, and then select **Upload My Add-in**.</span></span> 

    <span data-ttu-id="46a95-135">На следующем изображении показана вкладка **МОИ НАДСТРОЙКИ** для обычных записных книжек.</span><span class="sxs-lookup"><span data-stu-id="46a95-135">The following image shows the **MY ADD-INS** tab for consumer notebooks.</span></span>

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

5. <span data-ttu-id="46a95-136">В диалоговом окне "Отправить надстройку" выберите **manifest.xml** в папке проекта и нажмите кнопку **Отправить**.</span><span class="sxs-lookup"><span data-stu-id="46a95-136">In the Upload Add-in dialog, browse to **manifest.xml** in your project folder, and then choose **Upload**.</span></span> 

6. <span data-ttu-id="46a95-137">На вкладке **Главная** ленты нажмите кнопку **Показать область задач**.</span><span class="sxs-lookup"><span data-stu-id="46a95-137">From the **Home** tab, choose the **Show Taskpane** button in the ribbon.</span></span> <span data-ttu-id="46a95-138">Область задач надстройки откроется в iFrame рядом со страницей OneNote.</span><span class="sxs-lookup"><span data-stu-id="46a95-138">The add-in task pane opens in an iFrame next to the OneNote page.</span></span>

7. <span data-ttu-id="46a95-139">В нижней части области задач щелкните ссылку **Выполнить**, чтобы настроить заголовок страницы и добавить контур к тексту страницы.</span><span class="sxs-lookup"><span data-stu-id="46a95-139">At the bottom of the task pane, choose the **Run** link to set the page title and add an outline to the body of the page.</span></span>

    ![Надстройка OneNote, созданная на основе этого руководства](../images/onenote-first-add-in-4.png)

## <a name="next-steps"></a><span data-ttu-id="46a95-141">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="46a95-141">Next steps</span></span>

<span data-ttu-id="46a95-142">Поздравляем! Вы успешно создали надстройку области задач OneNote!</span><span class="sxs-lookup"><span data-stu-id="46a95-142">Congratulations, you've successfully created a OneNote task pane add-in!</span></span> <span data-ttu-id="46a95-143">Следующим шагом узнайте больше об основных понятиях, связанных с созданием надстроек OneNote.</span><span class="sxs-lookup"><span data-stu-id="46a95-143">Next, learn more about the core concepts of building OneNote add-ins.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="46a95-144">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="46a95-144">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a><span data-ttu-id="46a95-145">См. также</span><span class="sxs-lookup"><span data-stu-id="46a95-145">See also</span></span>

* [<span data-ttu-id="46a95-146">Обзор платформы надстроек Office</span><span class="sxs-lookup"><span data-stu-id="46a95-146">Office Add-ins platform overview</span></span>](../overview/office-add-ins.md)
* [<span data-ttu-id="46a95-147">Создание надстроек Office</span><span class="sxs-lookup"><span data-stu-id="46a95-147">Building Office Add-ins</span></span>](../overview/office-add-ins-fundamentals.md)
* [<span data-ttu-id="46a95-148">Разработка надстроек Office</span><span class="sxs-lookup"><span data-stu-id="46a95-148">Develop Office Add-ins</span></span>](../develop/develop-overview.md)
- [<span data-ttu-id="46a95-149">Обзор API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="46a95-149">OneNote JavaScript API programming overview</span></span>](../onenote/onenote-add-ins-programming-overview.md)
- [<span data-ttu-id="46a95-150">Справочник по API JavaScript для OneNote</span><span class="sxs-lookup"><span data-stu-id="46a95-150">OneNote JavaScript API reference</span></span>](../reference/overview/onenote-add-ins-javascript-reference.md)
- [<span data-ttu-id="46a95-151">Пример надстройки Rubric Grader</span><span class="sxs-lookup"><span data-stu-id="46a95-151">Rubric Grader sample</span></span>](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)

