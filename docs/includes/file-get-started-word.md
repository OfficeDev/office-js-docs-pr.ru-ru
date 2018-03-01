# <a name="build-your-first-word-add-in"></a>Создание первой надстройки Word

_Относится к: Word 2016, Word для iPad, Word для Mac_

В этой статье мы разберем, как создать надстройку Word, используя jQuery и API JavaScript для Word. 

## <a name="create-the-add-in"></a>Создание надстройки 

[!include[Choose your editor](../includes/quickstart-choose-editor.md)]

# <a name="visual-studiotabvisual-studio"></a>[Visual Studio](#tab/visual-studio)

### <a name="prerequisites"></a>Необходимые компоненты

[!include[Quickstart prerequisites](../includes/quickstart-vs-prerequisites.md)]

### <a name="create-the-add-in-project"></a>Создание проекта надстройки

1. В строке меню Visual Studio выберите **Файл** > **Создать** > **Проект**.
    
2. В списке типов проекта разверните узел **Visual C#** или **Visual Basic**, разверните **Office/SharePoint**, а затем выберите **Надстройки** > **Веб-надстройка Word**. 

3. Укажите имя проекта и нажмите кнопку **ОК**.

4. Visual Studio создаст решение, и два соответствующих проекта появятся в **обозревателе решений**. Файл **Home.html** откроется в Visual Studio.
    
### <a name="explore-the-visual-studio-solution"></a>Обзор решения Visual Studio

[!include[Description of Visual Studio projects](../includes/quickstart-vs-solution.md)]

### <a name="update-the-code"></a>Обновление кода

1. Файл **Home.html** содержит HTML-контент, который будет отображаться в области задач надстройки. В файле **Home.html** замените элемент `<body>` на приведенную ниже часть кода и сохраните файл.
 
    ```html
    <body>
        <div id="content-header">
            <div class="padding">
                <h1>Welcome</h1>
            </div>
        </div>    
        <div id="content-main">
            <div class="padding">
                <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <br /><br />
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <br /><br />
                <button id="proverb">Add Chinese proverb</button>
            </div>
        </div>
        <br />
        <div id="supportedVersion"/>
    </body>
    ```

2. Откройте файл **Home.js** в корневой папке проекта веб-приложения. Этот файл содержит скрипт надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.

    ```js
    'use strict';
    
    (function () {

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. Откройте файл **Home.css** в корневой папке проекта веб-приложения. Этот файл определяет специальные стили надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a>Обновление манифеста

1. Откройте XML-файл манифеста в проекте надстройки. Этот файл определяет параметры и возможности надстройки.

2. Элемент `ProviderName` содержит заполнитель. Замените его на свое имя.

3. Атрибут `DefaultValue` элемента `DisplayName` содержит заполнитель. Замените его на строку **Моя надстройка Office**.

4. Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его строкой **Надстройка области задач для Word**.

5. Сохраните файл.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="try-it-out"></a>Проверка

1. Испытайте новую надстройку Word в Visual Studio, нажав клавишу F5 или кнопку **Запустить**, чтобы запустить Word с кнопкой надстройки **Показать область задач** на ленте. Надстройка будет размещена на локальном сервере IIS.

2. В Word выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Снимок экрана: приложение Word с выделенной кнопкой "Показать область задач"](../images/word-quickstart-addin-0.png)

3. В области задач нажмите любую кнопку, чтобы добавить стандартный текст в документ.

    ![Снимок экрана: приложение Word с загруженной надстройкой, добавляющей стандартный текст.](../images/word-quickstart-addin-1b.png)

# <a name="any-editortabvisual-studio-code"></a>[Любой редактор](#tab/visual-studio-code)

### <a name="prerequisites"></a>Необходимые компоненты

- [Node.js](https://nodejs.org)

- Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).

    ```bash
    npm install -g yo generator-office
    ```

### <a name="create-the-add-in-project"></a>Создание проекта надстройки

1. Создайте на локальном диске папку и назовите ее `my-word-addin`. В ней вы будете создавать файлы для надстройки.

2. Перейдите к новой папке.

    ```bash
    cd my-word-addin
    ```

3. С помощью генератора Yeoman создайте проект надстройки Word. Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.

    ```bash
    yo office
    ```

    - **Вы хотите создать новую вложенную папку для проекта?:** `No`
    - **Как вы хотите назвать надстройку?:** `My Office Add-in`
    - **Какое клиентское приложение Office должно поддерживаться?:** `Word`
    - **Вы хотите создать новую надстройку?:** `Yes`
    - **Вы хотите использовать TypeScript?:** `No`
    - **Choose a framework:** `Jquery` (Выберите платформу)

    Затем генератор предложит вам открыть файл **resource.html**. В нашем случае открывать его не обязательно, но можете заглянуть, если вам интересно! Выберите Yes (Да) или No (Нет), чтобы завершить работу мастера, и подождите, пока работает генератор.

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-word-jquery.png)

### <a name="update-the-code"></a>Обновление кода

1. В редакторе кода откройте файл **index.html** из корневой папки проекта. Этот файл содержит HTML-контент, который будет отображаться в области задач надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл. В этой надстройке будет три кнопки. При нажатии любой из них в документ будет добавляться стандартный текст.

    ```html
    <!DOCTYPE html>
    <html>
        <head>
            <meta charset="UTF-8" />
            <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
            <title>Boilerplate text app</title>
            <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
            <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
            <script src="app.js" type="text/javascript"></script>
            <link href="app.css" rel="stylesheet" type="text/css" />
        </head>
        <body>
            <div id="content-header">
                <div class="padding">
                    <h1>Welcome</h1>
                </div>
            </div>    
            <div id="content-main">
                <div class="padding">
                    <p>Choose the buttons below to add boilerplate text to the document by using the Word JavaScript API.</p>
                    <br />
                    <h3>Try it out</h3>
                    <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                    <br /><br />
                    <button id="checkhov">Add quote from Anton Chekhov</button>
                    <br /><br />
                    <button id="proverb">Add Chinese proverb</button>
                </div>
            </div>
            <br />
            <div id="supportedVersion"/>
        </body>
    </html>
    ```

2. Откройте файл **app.js**, чтобы указать скрипт для надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл. Этот скрипт содержит код инициализации, а также код, вносящий изменения в документ Word, вставляя текст при нажатии кнопки. 

    ```js
    'use strict';
    
    (function () {

        // The initialize function is run each time the page is loaded.
        Office.initialize = function (reason) {
            $(document).ready(function () {

                // Use this to check whether the API is supported in the Word client.
                if (Office.context.requirements.isSetSupported('WordApi', 1.1)) {
                    // Do something that is only available via the new APIs
                    $('#emerson').click(insertEmersonQuoteAtSelection);
                    $('#checkhov').click(insertChekhovQuoteAtTheBeginning);
                    $('#proverb').click(insertChineseProverbAtTheEnd);
                    $('#supportedVersion').html('This code is using Word 2016 or greater.');
                }
                else {
                    // Just letting you know that this code will not work with your version of Word.
                    $('#supportedVersion').html('This code requires Word 2016 or greater.');
                }
            });
        };

        function insertEmersonQuoteAtSelection() {
            Word.run(function (context) {

                // Create a proxy object for the document.
                var thisDocument = context.document;

                // Queue a command to get the current selection.
                // Create a proxy range object for the selection.
                var range = thisDocument.getSelection();

                // Queue a command to replace the selected text.
                range.insertText('"Hitch your wagon to a star."\n', Word.InsertLocation.replace);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Ralph Waldo Emerson.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChekhovQuoteAtTheBeginning() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the start of the document body.
                body.insertText('"Knowledge is of no value unless you put it into practice."\n', Word.InsertLocation.start);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from Anton Chekhov.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }

        function insertChineseProverbAtTheEnd() {
            Word.run(function (context) {

                // Create a proxy object for the document body.
                var body = context.document.body;

                // Queue a command to insert text at the end of the document body.
                body.insertText('"To know the road ahead, ask those coming back."\n', Word.InsertLocation.end);

                // Synchronize the document state by executing the queued commands,
                // and return a promise to indicate task completion.
                return context.sync().then(function () {
                    console.log('Added a quote from a Chinese proverb.');
                });
            })
            .catch(function (error) {
                console.log('Error: ' + JSON.stringify(error));
                if (error instanceof OfficeExtension.Error) {
                    console.log('Debug info: ' + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

3. Откройте файл **app.css** в корневой папке проекта, чтобы указать специальные стили для надстройки. Замените все его содержимое на приведенный ниже код и сохраните файл.

    ```css
    #content-header {
        background: #2a8dd4;
        color: #fff;
        position: absolute;
        top: 0;
        left: 0;
        width: 100%;
        height: 80px; 
        overflow: hidden;
    }

    #content-main {
        background: #fff;
        position: fixed;
        top: 80px;
        left: 0;
        right: 0;
        bottom: 0;
        overflow: auto; 
    }

    .padding {
        padding: 15px;
    }
    ```

### <a name="update-the-manifest"></a>Обновление манифеста

1. Откройте файл **my-office-add-in-manifest.xml**, чтобы определить параметры и возможности надстройки.

2. Элемент `ProviderName` содержит заполнитель. Замените его на свое имя.

3. Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его строкой **Надстройка области задач для Word**.

4. Сохраните файл.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="My Office Add-in" />
    <Description DefaultValue="A task pane add-in for Word"/>
    ...
    ```

### <a name="start-the-dev-server"></a>Запуск сервера разработки

[!include[Start server section](../includes/quickstart-yo-start-server.md)] 

### <a name="try-it-out"></a>Проверка

1. Следуйте инструкциям для нужной платформы, чтобы загрузить неопубликованную надстройку в Word.

    - [Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - [Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - [iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. В Word выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Снимок экрана: приложение Word с выделенной кнопкой "Показать область задач"](../images/word-quickstart-addin-2.png)

3. В области задач нажмите любую кнопку, чтобы добавить стандартный текст в документ.

    ![Снимок экрана: приложение Word с загруженной надстройкой, добавляющей стандартный текст.](../images/word-quickstart-addin-1.png)

---

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку Word с помощью jQuery! Следующим шагом узнайте больше об основных понятиях, связанных с созданием надстроек Word.

> [!div class="nextstepaction"]
> [Обзор надстроек Word](../word/word-add-ins-programming-overview.md)

## <a name="see-also"></a>См. также

* [Обзор надстроек Word](../word/word-add-ins-programming-overview.md)
* [Примеры кода надстроек Word](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [Справочник по API JavaScript для Word](https://dev.office.com/reference/add-ins/word/word-add-ins-reference-overview)
