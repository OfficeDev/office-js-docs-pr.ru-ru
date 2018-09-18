# <a name="build-your-first-onenote-add-in"></a>Создание первой надстройки OneNote

В этой статье мы разберем, как создать надстройку OneNote, используя jQuery и API JavaScript для Office.

## <a name="prerequisites"></a>Необходимые компоненты

- [Node.js](https://nodejs.org)

- Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

1. Создайте на локальном диске папку и назовите ее `my-onenote-addin`. В ней вы будете создавать файлы для надстройки.

2. Перейдите к новой папке.

    ```bash
    cd my-onenote-addin
    ```

3. С помощью генератора Yeoman создайте проект надстройки OneNote. Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.

    ```bash
    yo office
    ```

    - **Выберите тип проекта:** `Office Add-in project using Jquery framework`
    - **Выберите тип сценария:** `Javascript`
    - **Как вы хотите назвать надстройку?** `My Office Add-in`
    - **Какое клиентское приложение Office должно поддерживаться?** `Onenote`

    ![Снимок экрана с вопросами и ответами в генераторе Yeoman](../images/yo-office-onenote-jquery.png)
    
    После завершения работы мастера, генератор создаст проект и установит поддерживающие компоненты узла.


## <a name="update-the-code"></a>Обновление кода

1. В редакторе кода откройте файл **index.html** из корневой папки проекта. Этот файл содержит HTML-контент, который будет отображаться в области задач надстройки.

2. Замените элемент `<main>` внутри элемента `<body>` приведенной ниже разметкой и сохраните файл. Эта разметка добавляет текстовую область и кнопку, используя [компоненты Office UI Fabric](https://developer.microsoft.com/en-us/fabric#/components).

    ```html
    <main class="ms-welcome__main">
        <br />
        <p class="ms-font-l">Enter content below</p>
        <div class="ms-TextField ms-TextField--placeholder">
            <textarea id="textBox" rows="5"></textarea>
        </div>
        <button id="addOutline" class="ms-welcome__action ms-Button ms-Button--hero ms-u-slideUpIn20">
            <span class="ms-Button-label">Add Outline</span>
            <span class="ms-Button-icon"><i class="ms-Icon"></i></span>
            <span class="ms-Button-description">Adds the content above to the current page.</span>
        </button>
    </main>
    ```

3. Откройте файл **src\index.js**, чтобы указать сценарий для надстройки. Замените все его содержимое указанным ниже кодом и сохраните файл.

    ```js
    'use strict';

    (function () {

        Office.initialize = function (reason) {
            $(document).ready(function () {
                // Set up event handler for the UI.
                $('#addOutline').click(addOutlineToPage);
            });
        };

        // Add the contents of the text area to the page.
        function addOutlineToPage() {        
            OneNote.run(function (context) {
                var html = '<p>' + $('#textBox').val() + '</p>';

                // Get the current page.
                var page = context.application.getActivePage();

                // Queue a command to load the page with the title property.             
                page.load('title'); 

                // Add an outline with the specified HTML to the page.
                var outline = page.addOutline(40, 90, html);

                // Run the queued commands, and return a promise to indicate task completion.
                return context.sync()
                    .then(function() {
                        console.log('Added outline to page ' + page.title);
                    })
                    .catch(function(error) {
                        app.showNotification("Error: " + error); 
                        console.log("Error: " + error); 
                        if (error instanceof OfficeExtension.Error) { 
                            console.log("Debug info: " + JSON.stringify(error.debugInfo)); 
                        } 
                    }); 
            });
        }
    })();
    ```

## <a name="update-the-manifest"></a>Обновление манифеста

1. Откройте файл **one-note-add-in-manifest.xml**, чтобы определить параметры и возможности надстройки.

2. Элемент `ProviderName` содержит заполнитель. Замените его на свое имя.

3. Атрибут `DefaultValue` элемента `Description` содержит заполнитель. Замените его строкой **Надстройка области задач для OneNote**.

4. Сохраните файл.

    ```xml
    ...
    <ProviderName>John Doe</ProviderName>
    <DefaultLocale>en-US</DefaultLocale>
    <!-- The display name of your add-in. Used on the store and various places of the Office UI such as the add-ins dialog. -->
    <DisplayName DefaultValue="OneNote Add-in" />
    <Description DefaultValue="A task pane add-in for OneNote"/>
    ...
    ```

## <a name="start-the-dev-server"></a>Запуск сервера разработки

[!include[Start server section](../includes/quickstart-yo-start-server.md)]

## <a name="try-it-out"></a>Проверка

1. Откройте записную книжку в [OneNote Online](https://www.onenote.com/notebooks).

2. Выберите **Вставка > Надстройки Office**. Откроется диалоговое окно "Надстройки Office".

    - Если вы вошли с помощью обычной учетной записи, выберите **Отправить надстройку** на вкладке **МОИ НАДСТРОЙКИ**.

    - Если вы вошли с помощью рабочей или учебной учетной записи, выберите **Отправить надстройку** на вкладке **МОЯ ОРГАНИЗАЦИЯ**. 

    На следующем изображении показана вкладка **МОИ НАДСТРОЙКИ** для обычных записных книжек.

    <img alt="The Office Add-ins dialog showing the MY ADD-INS tab" src="../images/onenote-office-add-ins-dialog.png" width="500">

3. В диалоговом окне "Отправить надстройку" выберите файл **one-note-add-in-manifest.xml** в папке проекта и нажмите **Отправить**. 

4. На вкладке **Главная** нажмите кнопку **Показать область задач** на ленте. Надстройка откроется в iFrame рядом со страницей OneNote.

5. Введите текст в текстовой области и нажмите кнопку **Добавить структуру**. Введенный текст будет добавлен на страницу. 

    ![Надстройка OneNote, созданная на основе этого руководства](../images/onenote-first-add-in.png)

## <a name="troubleshooting-and-tips"></a>Устранение неполадок и советы

- Для отладки надстройки можно использовать имеющиеся в браузере средства разработчика. При использовании веб-сервера Gulp и отладке в Internet Explorer или Chrome вы можете сохранить внесенные изменения в локальном расположении, а затем просто обновить iFrame надстройки.

- Просматривая объект OneNote, вы увидите, что доступные для использования свойства имеют действительные значения. Свойства, которые необходимо загрузить, имеют значение *undefined*. Разверните узел `_proto_`, чтобы увидеть свойства, которые определены для объекта, но еще не загружены.

   ![Выгруженный объект OneNote в отладчике](../images/onenote-debug.png)

- Если надстройка использует какие-либо HTTP-ресурсы, то вам потребуется включить смешанное содержимое в браузере. Надстройки, которые применяются в рабочей среде, должны использовать только безопасные HTTPS-ресурсы.

- Надстройки области задач можно открыть откуда угодно, но контентные надстройки вставляются только в содержимое стандартной страницы (не в заголовки, изображения, iFrames и т. д.). 

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку OneNote! Следующим шагом узнайте больше об основных понятиях, связанных с созданием надстроек OneNote.

> [!div class="nextstepaction"]
> [Обзор API JavaScript для OneNote](../onenote/onenote-add-ins-programming-overview.md)

## <a name="see-also"></a>См. также

- [Обзор создания кода с помощью API JavaScript для OneNote](../onenote/onenote-add-ins-programming-overview.md)
- [Справочник по API JavaScript для OneNote](https://docs.microsoft.com/javascript/office/overview/onenote-add-ins-javascript-reference?view=office-js)
- [Пример надстройки Rubric Grader](https://github.com/OfficeDev/OneNote-Add-in-Rubric-Grader)
- [Обзор платформы надстроек Office](../overview/office-add-ins.md)
