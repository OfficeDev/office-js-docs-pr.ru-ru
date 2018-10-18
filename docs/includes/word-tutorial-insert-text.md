На этом этапе руководства мы программным способом проверим, поддерживает ли надстройка текущую версию Word, установленную у пользователя, а затем вставим абзац в документ.

> [!NOTE]
> На этой странице описывается отдельный этап из руководства по надстройкам Word. Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Word](../tutorials/word-tutorial.yml), чтобы начать обучение с самого начала.

## <a name="code-the-add-in"></a>Написание кода надстройки

1. Откройте проект в редакторе кода. 
2. Откройте файл index.html.
3. Замените `TODO1` на следующую разметку:

    ```html
    <button class="ms-Button" id="insert-paragraph">Insert Paragraph</button>
    ```

4. Откройте файл app.js.
5. Замените `TODO1` на приведенный ниже код. Этот код определяет, поддерживает ли установленная у пользователя версия Word ту версию файла Word.js, которая включает все API, используемые на всех этапах данного руководства. В рабочей надстройке можно использовать текст условного блока, чтобы скрыть или отключить пользовательский интерфейс, где вызываются неподдерживаемые API. При этом пользователь по-прежнему сможет использовать те части надстройки, которые поддерживаются в его версии Word.

    ```js
    if (!Office.context.requirements.isSetSupported('WordApi', 1.3)) {
        console.log('Sorry. The tutorial add-in uses Word.js APIs that are not available in your version of Office.');
    } 
    ```

6. Замените `TODO2` на следующий код:

    ```js
    $('#insert-paragraph').click(insertParagraph);
    ```

7. Замените `TODO3` на приведенный ниже код. Обратите внимание на следующее:
   - Бизнес-логика Word.js будет добавлена в функцию, передаваемую методу `Word.run`. Эта логика выполняется не сразу. Вместо этого она добавляется в очередь ожидания команд.
   - Метод `context.sync` отправляет все команды из очереди в Word для выполнения.
   - За методом `Word.run` следует блок `catch`. Рекомендуется всегда следовать этой методике. 

    ```js
    function insertParagraph() {
        Word.run(function (context) {
            
            // TODO4: Queue commands to insert a paragraph into the document.

            return context.sync();
        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

8. Замените `TODO4` на приведенный ниже код. Обратите внимание:
   - Первый параметр метода `insertParagraph` — это текст нового абзаца.
   - Второй параметр — расположение в основном тексте, где будет вставлен абзац. Другие варианты вставки абзаца, родительским объектом которого является основной текст, — End и Replace. 

    ```js
    const docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");   
    ``` 

## <a name="test-the-add-in"></a>Тестирование надстройки

1. Откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.
2. Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в более раннюю версию JavaScript, поддерживаемую всеми ведущими приложениями, в которых могут работать надстройки Office.
3. Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.   
4. Загрузите неопубликованную надстройку одним из следующих способов:
    - Windows[](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Office Online[](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad и Mac[](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)
5. В меню **Главная** в Word выберите пункт **Показать область задач**.
6. В области задач нажмите кнопку **Insert Paragraph** (Вставить абзац).
7. Внесите изменение в абзац. 
8. Снова нажмите кнопку **Insert Paragraph**. Обратите внимание, что новый абзац находится над предыдущим, так как метод `insertParagraph` вставляет текст в начале основного текста документа.

    ![Руководство по Word: вставка абзаца](../images/word-tutorial-insert-paragraph.png)
