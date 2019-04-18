---
ms.date: 03/06/2019
description: Разработка настраиваемых функций в кратком руководстве по Excel.
title: Краткое руководство по настраиваемым функциям (Предварительная версия)
ms.prod: excel
localization_priority: Normal
ms.openlocfilehash: 80c500e1e30e8751a7d969d33cd7e13b7943b1b5
ms.sourcegitcommit: 6d375518c119d09c8d3fb5f0cc4583ba5b20ac03
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/18/2019
ms.locfileid: "31914300"
---
# <a name="get-started-developing-excel-custom-functions"></a><span data-ttu-id="58200-103">Приступая к разработке пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="58200-103">Get started developing Excel custom functions</span></span>

<span data-ttu-id="58200-104">С помощью пользовательских функций разработчики могут добавлять новые функции в Excel, определяя их в JavaScript или typescript как часть надстройки.</span><span class="sxs-lookup"><span data-stu-id="58200-104">With custom functions, developers can now add new functions to Excel by defining them in JavaScript or Typescript as part of an add-in.</span></span> <span data-ttu-id="58200-105">Пользователи Excel могут получить доступ к пользовательским функциям так же, как и к любой собственной `SUM()`функции в Excel, например.</span><span class="sxs-lookup"><span data-stu-id="58200-105">Excel users can access custom functions just as they would any native function in Excel, such as `SUM()`.</span></span>

## <a name="prerequisites"></a><span data-ttu-id="58200-106">Обязательные условия</span><span class="sxs-lookup"><span data-stu-id="58200-106">Prerequisites</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="58200-107">Для создания пользовательских функций потребуются следующие средства и связанные ресурсы.</span><span class="sxs-lookup"><span data-stu-id="58200-107">You'll need the following tools and related resources to begin creating custom functions.</span></span>

- <span data-ttu-id="58200-108">[Node.js](https://nodejs.org/en/) (версия 8.0.0 или более поздняя)</span><span class="sxs-lookup"><span data-stu-id="58200-108">[Node.js](https://nodejs.org/en/) (version 8.0.0 or later)</span></span>

- <span data-ttu-id="58200-109">[Git Bash](https://git-scm.com/downloads) (или другой клиент Git)</span><span class="sxs-lookup"><span data-stu-id="58200-109">[Git Bash](https://git-scm.com/downloads) (or another Git client)</span></span>

- <span data-ttu-id="58200-110">Последняя версия [Yeoman](https://yeoman.io/) и [генератора Yeoman для надстроек Office](https://www.npmjs.com/package/generator-office). Выполните в командной строке указанную ниже команду, чтобы установить эти инструменты глобально.</span><span class="sxs-lookup"><span data-stu-id="58200-110">The latest version of [Yeoman](https://yeoman.io/) and the [Yeoman generator for Office Add-ins](https://www.npmjs.com/package/generator-office). To install these tools globally, run the following command via the command prompt:</span></span>

    ```
    npm install -g yo generator-office
    ```

    > [!NOTE]
    > <span data-ttu-id="58200-111">Даже если вы ранее установили генератор Yeoman, рекомендуем обновить его до последней версии из NPM.</span><span class="sxs-lookup"><span data-stu-id="58200-111">Even if you've previously installed the Yeoman generator, we recommend you update your package to the latest version from npm.</span></span>

## <a name="build-your-first-custom-functions-project"></a><span data-ttu-id="58200-112">Создание проекта пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="58200-112">Build your first custom functions project</span></span>

<span data-ttu-id="58200-113">Чтобы начать работу, создайте проект пользовательских функций с помощью генератора Yeoman.</span><span class="sxs-lookup"><span data-stu-id="58200-113">To start, you'll use the Yeoman generator to create the custom functions project.</span></span> <span data-ttu-id="58200-114">Это позволит настроить для проекта правильную структуру папок, исходные файлы и зависимости, чтобы начать написание кода пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="58200-114">This will set up your project with the correct folder structure, source files, and dependencies to begin coding your custom functions.</span></span>

1. <span data-ttu-id="58200-115">Выполните указанную ниже команду и ответьте на вопросы, как показано ниже.</span><span class="sxs-lookup"><span data-stu-id="58200-115">Run the following command and then answer the prompts as follows.</span></span>

    ```
    yo office
    ```

    - <span data-ttu-id="58200-116">Выберите тип проекта: `Excel Custom Functions Add-in project (...)`</span><span class="sxs-lookup"><span data-stu-id="58200-116">Choose a project type: `Excel Custom Functions Add-in project (...)`</span></span>

    - <span data-ttu-id="58200-117">Выберите тип сценария: `JavaScript`</span><span class="sxs-lookup"><span data-stu-id="58200-117">Choose a script type: `JavaScript`</span></span>

    - <span data-ttu-id="58200-118">Как вы хотите назвать свою надстройку?</span><span class="sxs-lookup"><span data-stu-id="58200-118">What do you want to name your add-in?</span></span> `stock-ticker`

    ![Генератор Yeoman для надстройки Office, приглашающий к созданию пользовательских функций](../images/12-10-fork-cf-pic.jpg)

    <span data-ttu-id="58200-120">Генератор Yeoman создаст файлы проекта и установит вспомогательные компоненты Node.</span><span class="sxs-lookup"><span data-stu-id="58200-120">The Yeoman generator will create the project files and install supporting Node components.</span></span>

2. <span data-ttu-id="58200-121">Перейдите к только что созданной папке проекта.</span><span class="sxs-lookup"><span data-stu-id="58200-121">Navigate to the project folder you just created.</span></span>

    ```
    cd stock-ticker
    ```

3. <span data-ttu-id="58200-122">Доверяете самозаверяющий сертификату, который требуется для запуска этого проекта.</span><span class="sxs-lookup"><span data-stu-id="58200-122">Trust the self-signed certificate you need to run this project.</span></span> <span data-ttu-id="58200-123">Подробные инструкции для Windows или Mac см. в статье [Добавление самозаверяющих сертификатов в качестве доверенных корневых сертификатов](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span><span class="sxs-lookup"><span data-stu-id="58200-123">For detailed instructions for either Windows or Mac, see [Adding Self Signed Certificates as Trusted Root Certificate](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).</span></span>  

4. <span data-ttu-id="58200-124">Выполните сборку проекта.</span><span class="sxs-lookup"><span data-stu-id="58200-124">Build the project.</span></span>

    ```
    npm run build
    ```

5. <span data-ttu-id="58200-125">Запустите локальный веб-сервер, работающий на Node.js.</span><span class="sxs-lookup"><span data-stu-id="58200-125">Start the local web server, which runs in Node.js.</span></span>

    - <span data-ttu-id="58200-126">Если вы используете Excel для Windows для тестирования пользовательских функций, выполните следующую команду для запуска локального веб-сервера, запуска Excel и Загрузка неопубликованных надстройки:</span><span class="sxs-lookup"><span data-stu-id="58200-126">If you use Excel for Windows to test your custom functions, run the following command to start the local web server, launch Excel, and sideload the add-in:</span></span>

        ```
         npm run start
        ```
        <span data-ttu-id="58200-127">После выполнения этой команды в командной строки будут отображаться сведения о запуске веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="58200-127">After running this command, your command prompt will show details about starting the web server.</span></span> <span data-ttu-id="58200-128">Excel начнет работу с загруженной надстройкой.</span><span class="sxs-lookup"><span data-stu-id="58200-128">Excel will start with your add-in loaded.</span></span> <span data-ttu-id="58200-129">Если надстройка не загружается, проверьте правильность выполнения шага 3.</span><span class="sxs-lookup"><span data-stu-id="58200-129">If you add-in does not load, check that you have completed step 3 properly.</span></span>

    - <span data-ttu-id="58200-130">Если вы используете Excel Online для тестирования пользовательских функций, выполните следующую команду для запуска локального веб-сервера:</span><span class="sxs-lookup"><span data-stu-id="58200-130">If you use Excel Online to test your custom functions, run the following command to start the local web server:</span></span>

        ```
        npm run start-web
        ```

         <span data-ttu-id="58200-131">После выполнения этой команды в командной строки будут отображаться сведения о запуске веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="58200-131">After running this command, your command prompt will show details about starting the web server.</span></span> <span data-ttu-id="58200-132">Чтобы использовать функции, откройте новую книгу в Excel Online.</span><span class="sxs-lookup"><span data-stu-id="58200-132">To use your functions, open a new workbook in Excel Online.</span></span> <span data-ttu-id="58200-133">В этой книге вам потребуется загрузить надстройку.</span><span class="sxs-lookup"><span data-stu-id="58200-133">In this workbook, you'll need to load your add-in.</span></span> 

        <span data-ttu-id="58200-134">Для этого перейдите на вкладку **Вставка** на ленте и выберите **получить надстройки**. В открывшемся новом окне убедитесь, что вы используете вкладку **Мои надстройки** . Затем выберите **Управление моими надстройкАми _Гт_ отправить мою надстройку**.</span><span class="sxs-lookup"><span data-stu-id="58200-134">To do this, select the **Insert** tab on the ribbon and select **Get Add-ins**. In the resulting new window, ensure you are on the **My Add-ins** tab. Next, select **Manage My Add-ins > Upload My Add-in**.</span></span> <span data-ttu-id="58200-135">Найдите файл манифеста и отправьте его.</span><span class="sxs-lookup"><span data-stu-id="58200-135">Browse for your manifest file and upload it.</span></span> <span data-ttu-id="58200-136">Если надстройка не загружается, проверьте правильность выполнения шага 3.</span><span class="sxs-lookup"><span data-stu-id="58200-136">If your add-in does not load, check you've completed step 3 correctly.</span></span>

## <a name="try-out-the-prebuilt-custom-functions"></a><span data-ttu-id="58200-137">Опробуйте готовые пользовательские функции</span><span class="sxs-lookup"><span data-stu-id="58200-137">Try out the prebuilt custom functions</span></span>

<span data-ttu-id="58200-138">Проект пользовательских функций, созданный с помощью генератора Yeoman, содержит некоторые готовые пользовательские функции, определенные в файле **src/customfunctions.js**.</span><span class="sxs-lookup"><span data-stu-id="58200-138">The custom functions project that you created by using the Yeoman generator contains some prebuilt custom functions, defined within the **src/customfunctions.js** file.</span></span> <span data-ttu-id="58200-139">Файл **manifest.xml** в корневом каталоге проекта указывает, что все пользовательские функции принадлежат пространству имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="58200-139">The **manifest.xml** file in the root directory of the project specifies that all custom functions belong to the `CONTOSO` namespace.</span></span>

<span data-ttu-id="58200-140">В книге Excel проверьте `ADD` пользовательскую функцию, выполнив следующие действия:</span><span class="sxs-lookup"><span data-stu-id="58200-140">In your Excel workbook, try out the `ADD` custom function by completing the following steps:</span></span>

1. <span data-ttu-id="58200-141">Выберите ячейку и введите текст `=CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="58200-141">Select a cell and type `=CONTOSO`.</span></span> <span data-ttu-id="58200-142">Обратите внимание на то, что в меню автозаполнения содержится список всех функций в пространстве имен `CONTOSO`.</span><span class="sxs-lookup"><span data-stu-id="58200-142">Notice that the autocomplete menu shows the list of all functions in the `CONTOSO` namespace.</span></span>

2. <span data-ttu-id="58200-143">Выполните `CONTOSO.ADD` функцию, используя числа `10` и `200` входные параметры, введя значение `=CONTOSO.ADD(10,200)` в ячейке и нажав клавишу ВВОД.</span><span class="sxs-lookup"><span data-stu-id="58200-143">Run the `CONTOSO.ADD` function, using numbers `10` and `200` as input parameters, by typing the value `=CONTOSO.ADD(10,200)` in the cell and pressing enter.</span></span>

<span data-ttu-id="58200-144">Пользовательская функция `ADD` вычисляет сумму двух чисел, которые вы указываете в качестве входных параметров.</span><span class="sxs-lookup"><span data-stu-id="58200-144">The `ADD` custom function computes the sum of the two numbers that you specify as input parameters.</span></span> <span data-ttu-id="58200-145">При вводе `=CONTOSO.ADD(10,200)` в ячейке должен отобразиться результат **210** после нажатия клавиши ВВОД.</span><span class="sxs-lookup"><span data-stu-id="58200-145">Typing `=CONTOSO.ADD(10,200)` should produce the result **210** in the cell after you press enter.</span></span>

## <a name="next-steps"></a><span data-ttu-id="58200-146">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="58200-146">Next steps</span></span>

<span data-ttu-id="58200-147">Поздравляем, вы успешно создали пользовательскую функцию в надстройке Excel!</span><span class="sxs-lookup"><span data-stu-id="58200-147">Congratulations, you've successfully created a custom function in an Excel add-in!</span></span> <span data-ttu-id="58200-148">Затем создайте более сложную надстройку с возможностью потоковой передачи данных.</span><span class="sxs-lookup"><span data-stu-id="58200-148">Next, build a more complex add-in with streaming data capability.</span></span> <span data-ttu-id="58200-149">Следующая ссылка поможет вам выполнить следующие действия в руководстве надстройки Excel с пользовательскими функциями.</span><span class="sxs-lookup"><span data-stu-id="58200-149">The following link takes you through the next steps in the Excel add-in with custom functions tutorial.</span></span>

> [!div class="nextstepaction"]
> [<span data-ttu-id="58200-150">Руководство по надстройкам Excel для пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="58200-150">Excel custom functions add-in tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md#create-a-custom-function-that-requests-data-from-the-web
)

## <a name="see-also"></a><span data-ttu-id="58200-151">См. также</span><span class="sxs-lookup"><span data-stu-id="58200-151">See also</span></span>

* [<span data-ttu-id="58200-152">Обзор пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="58200-152">Custom functions overview</span></span>](../excel/custom-functions-overview.md)
* [<span data-ttu-id="58200-153">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="58200-153">Custom functions metadata</span></span>](../excel/custom-functions-json.md)
* [<span data-ttu-id="58200-154">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="58200-154">Runtime for Excel custom functions</span></span>](../excel/custom-functions-runtime.md)
* [<span data-ttu-id="58200-155">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="58200-155">Custom functions best practices</span></span>](../excel/custom-functions-best-practices.md)
