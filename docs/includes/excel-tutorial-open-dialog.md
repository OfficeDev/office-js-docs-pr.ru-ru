<span data-ttu-id="d09eb-101">На данном заключительном этапе, указанном в руководстве, вы откроете диалоговое окно в своей надстройке, передадите сообщение из процесса диалогового окна в процесс области задач и закроете диалоговое окно.</span><span class="sxs-lookup"><span data-stu-id="d09eb-101">In this final step of the tutorial, you'll open a dialog in your add-in, pass a message from the dialog process to the task pane process, and close the dialog.</span></span> <span data-ttu-id="d09eb-102">Диалоговые окна надстройки Office *не модальные*: пользователь может продолжать работать и с документом в ведущем приложении Office, и с главной страницей в области задач.</span><span class="sxs-lookup"><span data-stu-id="d09eb-102">Office Add-in dialogs are *nonmodal*: a user can continue to interact with both the document in the host Office application and with the host page in the task pane.</span></span>

> [!NOTE]
> <span data-ttu-id="d09eb-103">Это один из разделов руководства по надстройкам Excel.</span><span class="sxs-lookup"><span data-stu-id="d09eb-103">This page describes an individual step of the Excel add-in tutorial.</span></span> <span data-ttu-id="d09eb-104">Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Excel](../tutorials/excel-tutorial.yml), чтобы начать обучение с самого начала.</span><span class="sxs-lookup"><span data-stu-id="d09eb-104">If you’ve arrived at this page via search engine results or other direct link, please go to the [Excel add-in tutorial](../tutorials/excel-tutorial.yml) introduction page to start the tutorial from the beginning.</span></span>

## <a name="create-the-dialog-page"></a><span data-ttu-id="d09eb-105">Создание страницы диалогового окна</span><span class="sxs-lookup"><span data-stu-id="d09eb-105">Create the dialog page</span></span>

1. <span data-ttu-id="d09eb-106">Откройте проект в редакторе кода.</span><span class="sxs-lookup"><span data-stu-id="d09eb-106">Open the project in your code editor.</span></span>
2. <span data-ttu-id="d09eb-107">Создайте в корневой папке проекта (где находится index.html) файл popup.html.</span><span class="sxs-lookup"><span data-stu-id="d09eb-107">Create a file in the root of the project (where index.html is) called popup.html.</span></span>
3. <span data-ttu-id="d09eb-p103">Добавьте в файл popup.html приведенный ниже код. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="d09eb-p103">Add the following markup to popup.html. Note:</span></span>
   - <span data-ttu-id="d09eb-110">На странице находится `<input>`, где пользователь будет вводить свое имя, и кнопка, при нажатии которой имя будет отправлено на страницу области задач, где оно отобразится.</span><span class="sxs-lookup"><span data-stu-id="d09eb-110">The page has a `<input>` where the user will enter his or her name and a button that will send the name to the page in the task pane where it will be displayed.</span></span>
   - <span data-ttu-id="d09eb-111">Код загружает скрипт под названием popup.js, который будет создан на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="d09eb-111">The markup loads a script called popup.js that you will create in a later step.</span></span>
   - <span data-ttu-id="d09eb-112">Он загружает также библиотеку Office.JS и jQuery, так как они будут использоваться в popup.js.</span><span class="sxs-lookup"><span data-stu-id="d09eb-112">It also loads the Office.JS library and jQuery because they will be used in popup.js.</span></span>

    ```html
    <!DOCTYPE html>
    <html>
        <head lang="en">
            <title>Dialog for My Office Add-in</title>
            <meta charset="UTF-8">
            <meta name="viewport" content="width=device-width, initial-scale=1">
        
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.min.css" />
            <link rel="stylesheet" href="node_modules/office-ui-fabric-js/dist/css/fabric.components.css" />
            <link rel="stylesheet" href="app.css">
    
            <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>
            <script type="text/javascript" src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.2.1.min.js"></script>
            <script type="text/javascript" src="popup.js"></script>
    
        </head>
         <body style="display:flex;flex-direction:column;align-items:center;justify-content:center">
         <div class="padding">
            <p class="ms-font-xl">ENTER YOUR NAME</p>
         </div>        
        <div class="padding">
            <input id="name-box" type="text"/>
        <div>
        <div class="padding">
            <button id="ok-button" class="ms-Button">OK</button>
        </div>
    </body>
    </html>
    ```

4. <span data-ttu-id="d09eb-113">Создайте в корневой папке проекта файл popup.js.</span><span class="sxs-lookup"><span data-stu-id="d09eb-113">Create a file in the root of the project called popup.js.</span></span>
5. <span data-ttu-id="d09eb-p104">Добавьте в файл popup.js приведенный ниже код. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="d09eb-p104">Add the following code to popup.js. Note:</span></span>
   - <span data-ttu-id="d09eb-116">*Каждая страница, вызывающая API в библиотеке Office.JS, должна назначать функцию свойству `Office.initialize`.*</span><span class="sxs-lookup"><span data-stu-id="d09eb-116">*Every page that calls APIs in the Office.JS library must assign a function to the `Office.initialize` property.*</span></span> <span data-ttu-id="d09eb-117">Если в инициализации нет необходимости, тело функции может быть пустым, но свойство не должно оставаться неопределенным, иметь значение NULL или значение, не предназначенное для функции.</span><span class="sxs-lookup"><span data-stu-id="d09eb-117">If no initialization is needed, then the function can have an empty body, but the property must not be left undefined, assigned to null or to a non-function value.</span></span> <span data-ttu-id="d09eb-118">Файл app.js в корневом каталоге проекта можно рассматривать как пример.</span><span class="sxs-lookup"><span data-stu-id="d09eb-118">For an example, see the app.js file in the project root.</span></span> <span data-ttu-id="d09eb-119">Код, который выполняет назначение, должен быть запущен до каких-либо вызовов Office.JS, поэтому назначение указано в файле скрипта, загружаемом страницей, как в этом случае.</span><span class="sxs-lookup"><span data-stu-id="d09eb-119">The code that makes the assignment must run before any calls to Office.JS; hence the assignment is in a script file that is loaded by the page, as it is in this case.</span></span>
   - <span data-ttu-id="d09eb-p106">Функция jQuery `ready` вызывается в методе `initialize`. Существует почти универсальное правило: код загрузки (в том числе начальной) или инициализации из других библиотек JavaScript должен находиться в функции `Office.initialize`.</span><span class="sxs-lookup"><span data-stu-id="d09eb-p106">The jQuery `ready` function is called inside the `initialize` method. It is an almost universal rule that the loading, initializing, or bootstrapping code of other JavaScript libraries should be inside the `Office.initialize` function.</span></span>

    ```js
    (function () {
    "use strict";

        Office.initialize = function() {        
            $(document).ready(function () {  
    
                // TODO1: Assign handler to the OK button.
    
            });
        }

        // TODO2: Create the OK button handler
    
    }());    
    ```

6. <span data-ttu-id="d09eb-122">Замените `TODO1` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="d09eb-122">Replace `TODO1` with the following code.</span></span> <span data-ttu-id="d09eb-123">Вы создадите функцию `sendStringToParentPage` на следующем этапе.</span><span class="sxs-lookup"><span data-stu-id="d09eb-123">You'll create the `sendStringToParentPage` function in the next step.</span></span>

    ```js
    $('#ok-button').click(sendStringToParentPage);
    ```

7. <span data-ttu-id="d09eb-124">Замените `TODO2` приведенным ниже кодом.</span><span class="sxs-lookup"><span data-stu-id="d09eb-124">Replace `TODO2` with the following code.</span></span> <span data-ttu-id="d09eb-125">Метод `messageParent` передает свой параметр родительской странице (в данном случае это страница на панели задач).</span><span class="sxs-lookup"><span data-stu-id="d09eb-125">The `messageParent` method passes its parameter to the parent page, in this case, the page in the task pane.</span></span> <span data-ttu-id="d09eb-126">Параметр может быть логическим или строковым. Во втором случае подразумевается все, что можно сериализовать, представив в виде строки (например, XML или JSON).</span><span class="sxs-lookup"><span data-stu-id="d09eb-126">The parameter can be a boolean or a string, which includes anything that can be serialized as a string, such as XML or JSON.</span></span> 

    ```js
    function sendStringToParentPage() {
        var userName = $('#name-box').val();
        Office.context.ui.messageParent(userName);
    }
    ```

8. <span data-ttu-id="d09eb-127">Сохраните файл.</span><span class="sxs-lookup"><span data-stu-id="d09eb-127">Save the file.</span></span>

   > [!NOTE]
   > <span data-ttu-id="d09eb-128">Файл popup.html и загружаемый им файл popup.js выполняются в полностью отдельном процессе Internet Explorer из области задач надстройки.</span><span class="sxs-lookup"><span data-stu-id="d09eb-128">The popup.html file, and the popup.js file that it loads, run in an entirely separate Internet Explorer process from the add-in's task pane.</span></span> <span data-ttu-id="d09eb-129">Если файл popup.js был передан в тот же файл bundle.js, что и файл app.js, надстройка загрузит два экземпляра файла bundle.js, и это отменяет цель объединения.</span><span class="sxs-lookup"><span data-stu-id="d09eb-129">If the popup.js was transpiled into the same bundle.js file as the app.js file, then the add-in would have to load two copies of the bundle.js file, which defeats the purpose of bundling.</span></span> <span data-ttu-id="d09eb-130">Кроме того, файл popup.js не содержит код JavaScript, который не поддерживается в IE.</span><span class="sxs-lookup"><span data-stu-id="d09eb-130">In addition, the popup.js file does not contain any JavaScript that is unsupported by IE.</span></span> <span data-ttu-id="d09eb-131">По этим двум причинам эта надстройка не передает файл popup.js вообще.</span><span class="sxs-lookup"><span data-stu-id="d09eb-131">For these two reasons, this add-in does not transpile the popup.js file at all.</span></span> 


## <a name="open-the-dialog-from-the-task-pane"></a><span data-ttu-id="d09eb-132">Открытие диалогового окна из области задач</span><span class="sxs-lookup"><span data-stu-id="d09eb-132">Open the dialog from the task pane</span></span>

1. <span data-ttu-id="d09eb-133">Откройте файл index.html.</span><span class="sxs-lookup"><span data-stu-id="d09eb-133">Open the file index.html.</span></span>
2. <span data-ttu-id="d09eb-134">Под `div` с кнопкой `freeze-header` добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="d09eb-134">Below the `div` that contains the `freeze-header` button, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <button class="ms-Button" id="open-dialog">Open Dialog</button>          
    </div>
    ```

3. <span data-ttu-id="d09eb-135">В диалоговом окне пользователю будет предложено ввести имя и передать имя пользователя в область задач.</span><span class="sxs-lookup"><span data-stu-id="d09eb-135">The dialog will prompt the user to enter a name and pass the user's name to the task pane.</span></span> <span data-ttu-id="d09eb-136">Область задач отобразит его в подписи.</span><span class="sxs-lookup"><span data-stu-id="d09eb-136">The task pane will display it in a label.</span></span> <span data-ttu-id="d09eb-137">Непосредственно под только что добавленным тегом `div` добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="d09eb-137">Immediately below the `div` that you just added, add the following markup:</span></span>

    ```html
    <div class="padding">            
        <label id="user-name"></label>            
    </div>
    ```

4. <span data-ttu-id="d09eb-138">Откройте файл app.js.</span><span class="sxs-lookup"><span data-stu-id="d09eb-138">Open the app.js file.</span></span>

5. <span data-ttu-id="d09eb-139">Под строкой, назначающей обработчик щелчков для кнопки `freeze-header`, добавьте приведенный ниже код.</span><span class="sxs-lookup"><span data-stu-id="d09eb-139">Below the line that assigns a click handler to the `freeze-header` button, add the following code.</span></span> <span data-ttu-id="d09eb-140">Вы создадите метод `openDialog` на одном из следующих шагов.</span><span class="sxs-lookup"><span data-stu-id="d09eb-140">You'll create the `openDialog` method in a later step.</span></span>

    ```js
    $('#open-dialog').click(openDialog);
    ```

6. <span data-ttu-id="d09eb-p112">Под функцией `freezeHeader` добавьте указанное ниже объявление. Эта переменная удерживает объект в контексте выполнения родительской страницы, который служит посредником для контекста выполнения страницы диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="d09eb-p112">Below the `freezeHeader` function add the following declaration. This variable is used to hold an object in the parent page's execution context that acts as an intermediator to the dialog page's execution context.</span></span>

    ```js
    let dialog = null;
    ```

7. <span data-ttu-id="d09eb-143">Добавьте приведенную ниже функцию под объявлением `dialog`.</span><span class="sxs-lookup"><span data-stu-id="d09eb-143">Below the declaration of `dialog`, add the following function.</span></span> <span data-ttu-id="d09eb-144">Важно отметить, что в этом коде *отсутствует* вызов `Excel.run`.</span><span class="sxs-lookup"><span data-stu-id="d09eb-144">The important thing to notice about this code is what is *not* there: there is no call of `Excel.run`.</span></span> <span data-ttu-id="d09eb-145">Это связано с тем, что API, открывающий диалоговое окно, совместно используется всеми ведущими приложениями Office, поэтому относится к общему API JavaScript для Office, а не API для Excel.</span><span class="sxs-lookup"><span data-stu-id="d09eb-145">This is because the API to open a dialog is shared among all Office hosts, so it is part of the Office JavaScript Common API, not the Excel-specific API.</span></span>

    ```js
    function openDialog() {
        // TODO1: Call the Office Shared API that opens a dialog
    }
    ``` 

8. <span data-ttu-id="d09eb-p114">Замените `TODO1` приведенным ниже кодом. Примечание.</span><span class="sxs-lookup"><span data-stu-id="d09eb-p114">Replace `TODO1` with the following code. Note:</span></span>
   - <span data-ttu-id="d09eb-148">Метод `displayDialogAsync` открывает диалоговое окно в центре экрана.</span><span class="sxs-lookup"><span data-stu-id="d09eb-148">The `displayDialogAsync` method opens a dialog in the center of the screen.</span></span>
   - <span data-ttu-id="d09eb-149">Первый параметр — это URL-адрес открываемой страницы.</span><span class="sxs-lookup"><span data-stu-id="d09eb-149">The first parameter is the URL of the page to open.</span></span>
   - <span data-ttu-id="d09eb-p115">Второй параметр передает параметры. `height` и `width` — процентные значения размера окна для приложения Office.</span><span class="sxs-lookup"><span data-stu-id="d09eb-p115">The second parameter passes options. `height` and `width` are percentages of the size of the Office application's window.</span></span> 
   
    ```js
    Office.context.ui.displayDialogAsync(
        'https://localhost:3000/popup.html',
        {height: 45, width: 55},
        
        // TODO2: Add callback parameter.
    );
    ``` 

## <a name="process-the-message-from-the-dialog-and-close-the-dialog"></a><span data-ttu-id="d09eb-152">Обработка сообщения из диалогового окна и закрытие диалогового окна</span><span class="sxs-lookup"><span data-stu-id="d09eb-152">Process the message from the dialog and close the dialog</span></span>

1. <span data-ttu-id="d09eb-p116">Продолжайте работать в файле app.js. Замените `TODO2` приведенным ниже кодом. Обратите внимание:</span><span class="sxs-lookup"><span data-stu-id="d09eb-p116">Continue in the app.js file, and replace `TODO2` with the following code. Note:</span></span>
   - <span data-ttu-id="d09eb-155">Обратный вызов выполняется сразу же после успешного открытия диалогового окна и до того, как пользователь предпримет какие-либо действия в диалоговом окне.</span><span class="sxs-lookup"><span data-stu-id="d09eb-155">The callback is executed immediately after the dialog successfully opens and before the user has taken any action in the dialog.</span></span>
   - <span data-ttu-id="d09eb-156">— это объект, который выступает в качестве посредника между контекстами выполнения родительских страниц и страниц диалоговых окон.`result.value`</span><span class="sxs-lookup"><span data-stu-id="d09eb-156">The `result.value` is the object that acts as a kind of middleman between the execution contexts of the parent and dialog pages.</span></span>
   - <span data-ttu-id="d09eb-157">Функция `processMessage` будет создана на более позднем этапе.</span><span class="sxs-lookup"><span data-stu-id="d09eb-157">The `processMessage` function will be created in a later step.</span></span> <span data-ttu-id="d09eb-158">Этот обработчик будет обрабатывать любые значения, которые отправляются со страницы диалогового окна с вызовами функции `messageParent`.</span><span class="sxs-lookup"><span data-stu-id="d09eb-158">This handler will process any values that are sent from the dialog page with calls of the `messageParent` function.</span></span>

    ```js
    function (result) {
        dialog = result.value;
        dialog.addEventHandler(Microsoft.Office.WebExtension.EventType.DialogMessageReceived, processMessage);
    }
    ```

2. <span data-ttu-id="d09eb-159">Добавьте указанную ниже функцию под функцией `openDialog`.</span><span class="sxs-lookup"><span data-stu-id="d09eb-159">Below the `openDialog` function, add the following function.</span></span>

    ```js
    function processMessage(arg) {
        $('#user-name').text(arg.message);
        dialog.close();
    }
    ```

## <a name="test-the-add-in"></a><span data-ttu-id="d09eb-160">Тестирование надстройки</span><span class="sxs-lookup"><span data-stu-id="d09eb-160">Test the add-in</span></span>

1. <span data-ttu-id="d09eb-161">Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера.</span><span class="sxs-lookup"><span data-stu-id="d09eb-161">If the Git bash window, or Node.JS-enabled system prompt, from the previous stage tutorial is still open, enter Ctrl-C twice to stop the running web server.</span></span> <span data-ttu-id="d09eb-162">Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.</span><span class="sxs-lookup"><span data-stu-id="d09eb-162">Otherwise, open a Git bash window, or Node.JS-enabled system prompt, and navigate to the **Start** folder of the project.</span></span>

     > [!NOTE]
     > <span data-ttu-id="d09eb-163">Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу.</span><span class="sxs-lookup"><span data-stu-id="d09eb-163">Although the browser-sync server reloads your add-in in the task pane every time you make a change to any file, including the app.js file, it does not retranspile the JavaScript, so you must repeat the build command in order for your changes to app.js to take effect.</span></span> <span data-ttu-id="d09eb-164">Для этого следует завершить процесс сервера, чтобы можно было получить приглашение на ввод команды сборки.</span><span class="sxs-lookup"><span data-stu-id="d09eb-164">In order to do this, you need to kill the server process in so that you can get a prompt to enter the build command.</span></span> <span data-ttu-id="d09eb-165">После сборки необходимо перезапустить сервер.</span><span class="sxs-lookup"><span data-stu-id="d09eb-165">After the build, you restart the server.</span></span> <span data-ttu-id="d09eb-166">Для этого выполните указанные ниже действия.</span><span class="sxs-lookup"><span data-stu-id="d09eb-166">The next few steps carry out this process.</span></span>

1. <span data-ttu-id="d09eb-167">Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в JavaScript более ранней версии, которую поддерживает Internet Explorer (используется приложением Excel в фоновом режиме для запуска надстроек Excel).</span><span class="sxs-lookup"><span data-stu-id="d09eb-167">Run the command `npm run build` to transpile your ES6 source code to an earlier version of JavaScript that is supported by Internet Explorer (which is used under-the-hood by Excel to run Excel add-ins).</span></span>
2. <span data-ttu-id="d09eb-168">Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.</span><span class="sxs-lookup"><span data-stu-id="d09eb-168">Run the command `npm start` to start a web server running on localhost.</span></span>
4. <span data-ttu-id="d09eb-169">Повторно загрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Show Taskpane** (Показать область задач) для повторного открытия надстройки.</span><span class="sxs-lookup"><span data-stu-id="d09eb-169">Reload the task pane by closing it, and then on the **Home** menu, select **Show Taskpane** to reopen the add-in.</span></span>
6. <span data-ttu-id="d09eb-170">Нажмите кнопку **Open Dialog** (Открыть диалоговое окно) в области задач.</span><span class="sxs-lookup"><span data-stu-id="d09eb-170">Choose the **Open Dialog** button in the task pane.</span></span> 
7. <span data-ttu-id="d09eb-171">Когда диалоговое окно открыто, перетащите его и измените его размер.</span><span class="sxs-lookup"><span data-stu-id="d09eb-171">While the dialog is open, drag it and resize it.</span></span> <span data-ttu-id="d09eb-172">Обратите внимание на то, что вы можете работать на листе и нажимать другие кнопки в области задач.</span><span class="sxs-lookup"><span data-stu-id="d09eb-172">Note that you can interact with the worksheet and press other buttons on the taskpane.</span></span> <span data-ttu-id="d09eb-173">Но запустить второе диалоговое окно с той же страницы области задач невозможно.</span><span class="sxs-lookup"><span data-stu-id="d09eb-173">But you cannot launch a second dialog from the same task pane page.</span></span>
8. <span data-ttu-id="d09eb-174">В диалоговом окне введите имя и нажмите кнопку **OK**.</span><span class="sxs-lookup"><span data-stu-id="d09eb-174">In the dialog, enter a name and choose **OK**.</span></span> <span data-ttu-id="d09eb-175">В области задач отобразится имя, и диалоговое окно закроется.</span><span class="sxs-lookup"><span data-stu-id="d09eb-175">The name appears on the task pane and the dialog closes.</span></span>
9. <span data-ttu-id="d09eb-176">При желании можно закомментировать строку `dialog.close();` в функции `processMessage`.</span><span class="sxs-lookup"><span data-stu-id="d09eb-176">Optionally, comment out the line `dialog.close();` in the `processMessage` function.</span></span> <span data-ttu-id="d09eb-177">Повторите шаги этого раздела.</span><span class="sxs-lookup"><span data-stu-id="d09eb-177">Then repeat the steps of this section.</span></span> <span data-ttu-id="d09eb-178">Диалоговое окно остается открытым, и вы можете изменить имя.</span><span class="sxs-lookup"><span data-stu-id="d09eb-178">The dialog stays open and you can change the name.</span></span> <span data-ttu-id="d09eb-179">Можно закрыть его вручную, нажав кнопку **X** в правом верхнему углу.</span><span class="sxs-lookup"><span data-stu-id="d09eb-179">You can close it manually by pressing the **X** button in the upper right corner.</span></span>

    ![Руководство по Excel: диалоговое окно](../images/excel-tutorial-dialog-open.png)

