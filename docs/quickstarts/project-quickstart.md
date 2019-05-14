---
title: Создание первой надстройки области задач Project
description: ''
ms.date: 05/08/2019
ms.prod: project
localization_priority: Priority
ms.openlocfilehash: d61f8d83b88dbe69ff0ba9cd4b0afef77a4f03d6
ms.sourcegitcommit: a99be9c4771c45f3e07e781646e0e649aa47213f
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/11/2019
ms.locfileid: "33952254"
---
# <a name="build-your-first-project-task-pane-add-in"></a><span data-ttu-id="71887-102">Создание первой надстройки области задач Project</span><span class="sxs-lookup"><span data-stu-id="71887-102">Build your first PowerPoint task pane add-in</span></span>

<span data-ttu-id="71887-103">В этой статье вы ознакомитесь с процессом создания надстройки для области задач Project.</span><span class="sxs-lookup"><span data-stu-id="71887-103">In this article, you'll walk through the process of building a PowerPoint task pane add-in.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="71887-104">Необходимые компоненты</span><span class="sxs-lookup"><span data-stu-id="71887-104">Prerequisites</span></span>

[!include[Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- <span data-ttu-id="71887-105">Project 2016 или более поздней версии для Windows</span><span class="sxs-lookup"><span data-stu-id="71887-105">Project 2016 or later on Windows</span></span>

## <a name="create-the-add-in"></a><span data-ttu-id="71887-106">Создание надстройки</span><span class="sxs-lookup"><span data-stu-id="71887-106">Create the add-in</span></span>

1. <span data-ttu-id="71887-107">С помощью генератора Yeoman создайте проект надстройки Project.</span><span class="sxs-lookup"><span data-stu-id="71887-107">Use the Yeoman generator to create a Project add-in project.</span></span> <span data-ttu-id="71887-108">Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="71887-108">Run the following command and then answer the prompts as follows:</span></span>

    ```command&nbsp;line
    yo office
    ```

    - <span data-ttu-id="71887-109">**Выберите тип проекта:** `Office Add-in Task Pane project`</span><span class="sxs-lookup"><span data-stu-id="71887-109">**Choose a project type:** `Office Add-in Task Pane project`</span></span>
    - <span data-ttu-id="71887-110">**Выберите тип сценария:** `Javascript`</span><span class="sxs-lookup"><span data-stu-id="71887-110">**Choose a script type:** `Javascript`</span></span>
    - <span data-ttu-id="71887-111">**Как вы хотите назвать надстройку?**</span><span class="sxs-lookup"><span data-stu-id="71887-111">**What do you want to name your add-in?**</span></span> `My Office Add-in`
    - <span data-ttu-id="71887-112">**Какое клиентское приложение Office должно поддерживаться?**</span><span class="sxs-lookup"><span data-stu-id="71887-112">**Which Office client application would you like to support?**</span></span> `Project`

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-project.png)
    
    <span data-ttu-id="71887-114">После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="71887-114">After you complete the wizard, the generator will create the project and install supporting Node components.</span></span>
    
2. <span data-ttu-id="71887-115">Перейдите к корневой папке проекта.</span><span class="sxs-lookup"><span data-stu-id="71887-115">Navigate to the root folder of the project.</span></span>

    ```command&nbsp;line
    cd "My Office Add-in"
    ```

## <a name="explore-the-project"></a><span data-ttu-id="71887-116">Знакомство с проектом</span><span class="sxs-lookup"><span data-stu-id="71887-116">Explore the project</span></span>

<span data-ttu-id="71887-117">Проект надстройки, который вы создали с помощью генератора Yeoman, содержит образец кода для простейшей надстройки области задач.</span><span class="sxs-lookup"><span data-stu-id="71887-117">The add-in project that you've created with the Yeoman generator contains sample code for a very basic task pane add-in.</span></span> 

- <span data-ttu-id="71887-118">Файл **./manifest.xml** в корневом каталоге проекта определяет настройки и возможности надстройки.</span><span class="sxs-lookup"><span data-stu-id="71887-118">The **./manifest.xml** file in the root directory of the project defines the settings and capabilities of the add-in.</span></span>
- <span data-ttu-id="71887-119">Файл **./src/taskpane/taskpane.html** содержит разметку HTML для области задач.</span><span class="sxs-lookup"><span data-stu-id="71887-119">The **./src/taskpane/taskpane.html** file contains the HTML markup for the task pane.</span></span>
- <span data-ttu-id="71887-120">Файл **./src/taskpane/taskpane.css** содержит код CSS, который применяется к содержимому области задач.</span><span class="sxs-lookup"><span data-stu-id="71887-120">The **./src/taskpane/taskpane.css** file contains the CSS that's applied to content in the task pane.</span></span>
- <span data-ttu-id="71887-121">Файл **./src/taskpane/taskpane.js** содержит код API JavaScript для Office, который упрощает взаимодействие между областью задач и ведущим приложением Office.</span><span class="sxs-lookup"><span data-stu-id="71887-121">The **./src/taskpane/taskpane.js** file contains the Office JavaScript API code that facilitates interaction between the task pane and the Office host application.</span></span>

## <a name="update-the-code"></a><span data-ttu-id="71887-122">Обновление кода</span><span class="sxs-lookup"><span data-stu-id="71887-122">Update the code</span></span>

<span data-ttu-id="71887-123">Откройте файл **./src/taskpane/taskpane.js** в редакторе кода и добавьте приведенный ниже код в пределах функции **run**.</span><span class="sxs-lookup"><span data-stu-id="71887-123">In your code editor, open the file **./src/taskpane/taskpane.js** and add the following code within the **run** function.</span></span> <span data-ttu-id="71887-124">В этом коде используется API JavaScript для Office, чтобы настроить поле `Name` и поле `Notes` выбранной задачи.</span><span class="sxs-lookup"><span data-stu-id="71887-124">This code uses the Office JavaScript API to set the `Name` field and `Notes` field of the selected task.</span></span>

```js
var taskGuid;

// Get the GUID of the selected task
Office.context.document.getSelectedTaskAsync(
    function (result) {
        if (result.status === Office.AsyncResultStatus.Succeeded) {
            taskGuid = result.value;

            // Set the specified fields for the selected task.
            var targetFields = [Office.ProjectTaskFields.Name, Office.ProjectTaskFields.Notes];
            var fieldValues = ['New task name', 'Notes for the task.'];

            // Set the field value. If the call is successful, set the next field.
            for (var i = 0; i < targetFields.length; i++) {
                Office.context.document.setTaskFieldAsync(
                    taskGuid,
                    targetFields[i],
                    fieldValues[i],
                    function (result) {
                        if (result.status === Office.AsyncResultStatus.Succeeded) {
                            i++;
                        }
                        else {
                            var err = result.error;
                            console.log(err.name + ' ' + err.code + ' ' + err.message);
                        }
                    }
                );
            }
        } else {
            var err = result.error;
            console.log(err.name + ' ' + err.code + ' ' + err.message);
        }
    }
);
```

## <a name="try-it-out"></a><span data-ttu-id="71887-125">Проверка</span><span class="sxs-lookup"><span data-stu-id="71887-125">Try it out</span></span>

1. <span data-ttu-id="71887-126">Запустите локальный веб-сервер, выполнив следующую команду:</span><span class="sxs-lookup"><span data-stu-id="71887-126">Start the local web server by running the following command:</span></span>

    ```command&nbsp;line
    npm start
    ```

    > [!NOTE]
    > <span data-ttu-id="71887-127">Надстройки Office должны использовать HTTPS, а не HTTP, даже в случае разработки.</span><span class="sxs-lookup"><span data-stu-id="71887-127">Office Add-ins should use HTTPS, not HTTP, even when you are developing.</span></span> <span data-ttu-id="71887-128">Если вам будет предложено установить сертификат после того, как вы запустите `npm start`, примите предложение установить сертификат от генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="71887-128">If you are prompted to install a certificate after you run `npm start`, accept the prompt to install the certificate that the Yeoman generator provides.</span></span> 

2. <span data-ttu-id="71887-129">В Project создайте простой план проекта.</span><span class="sxs-lookup"><span data-stu-id="71887-129">In Project, create a simple project plan.</span></span>

3. <span data-ttu-id="71887-130">Загрузите свою надстройку в Project, следуя инструкциям в статье [Загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="71887-130">Load your add-in in Project by following the instructions in [Sideload Office Add-ins on Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).</span></span>

4. <span data-ttu-id="71887-131">Выберите отдельную задачу в проекте.</span><span class="sxs-lookup"><span data-stu-id="71887-131">Select a single task within the project.</span></span>

5. <span data-ttu-id="71887-132">В нижней части области задач щелкните ссылку **Выполнить**, чтобы переименовать выбранную задачу и добавить к ней примечания.</span><span class="sxs-lookup"><span data-stu-id="71887-132">At the bottom of the task pane, choose the **Run** link to rename the selected task and add notes to the selected task.</span></span>

    ![Снимок экрана: приложение Project с загруженной надстройкой области задач](../images/project-quickstart-addin-1.png)

## <a name="next-steps"></a><span data-ttu-id="71887-134">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="71887-134">Next steps</span></span>

<span data-ttu-id="71887-135">Поздравляем! Вы успешно создали надстройку области задач Project!</span><span class="sxs-lookup"><span data-stu-id="71887-135">Congratulations, you've successfully created a PowerPoint task pane add-in!</span></span> <span data-ttu-id="71887-136">Следующим шагом узнайте больше о возможностях надстроек Project и изучите распространенные сценарии.</span><span class="sxs-lookup"><span data-stu-id="71887-136">Next, learn more about the capabilities of a Project add-in and explore common scenarios.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="71887-137">Надстройки Project</span><span class="sxs-lookup"><span data-stu-id="71887-137">Project add-ins</span></span>](../project/project-add-ins.md)

