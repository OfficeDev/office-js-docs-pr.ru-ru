На этом этапе руководства мы добавим текст в выбранные диапазоны текста и за их пределами, а также заменим текст выбранного диапазона. 

> [!NOTE]
> На этой странице описывается отдельный этап из руководства по надстройкам Word. Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Word](../tutorials/word-tutorial.yml), чтобы начать обучение с самого начала.

## <a name="add-text-inside-a-range"></a>Добавление текста в диапазон

1. Откройте проект в редакторе кода. 
2. Откройте файл index.html.
3. Под элементом `div`, содержащим кнопку `change-font`, добавьте следующую разметку:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-text-into-range">Insert Abbreviation</button>            
    </div>
    ```

4. Откройте файл app.js.

5. Под строкой, назначающей обработчик нажатия кнопки `change-font`, добавьте следующий код:

    ```js
    $('#insert-text-into-range').click(insertTextIntoRange);
    ```

6. Добавьте приведенную ниже функцию под функцией `changeFont`.

    ```js
    function insertTextIntoRange() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert text into a selected range.

            // TODO2: Load the text of the range and sync so that the 
            //        current range text can be read.

            // TODO3: Queue commands to repeat the text of the original
            //        range at the end of the document.

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

7. Замените `TODO1` на приведенный ниже код. Обратите внимание:
   - Этот метод призван вставить аббревиатуру ["(C2R)"] в конце диапазона с текстом "Click-to-Run". Для простоты предполагается, что такая строка существует и пользователь выделил ее.
   - Первый параметр метода `Range.insertText` — это строка, вставляемая в объект `Range`.
   - Второй параметр указывает, в каком месте диапазона требуется вставить дополнительный текст. Помимо значения End, можно использовать значения Start, Before, After и Replace. 
   - Разница между значениями End и After состоит в том, что End вставляет новый текст в конце имеющегося диапазона, а After создает новый диапазон со строкой и вставляет его после имеющегося. Аналогично, Start вставляет текст в начале имеющегося диапазона, а Before вставляет новый диапазон. Replace заменяет текст существующего диапазона на строку из первого параметра.
   - На одном из предыдущих этапов руководства вы могли заметить, что в методах insert* объекта body нет параметров Before и After. Это связано с тем, что содержимое невозможно добавлять за пределами основного текста документа.

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ``` 

8. Пропустим заполнитель `TODO2` до следующего этапа. Замените `TODO3` на приведенный ниже код. Он похож на код, созданный на первом этапе руководства, но теперь мы вставляем новый абзац в конце, а не в начале документа. Новый абзац покажет, что новый текст теперь входит в исходный диапазон.
 
    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text,
                             "End");
    ``` 

## <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>Добавление кода для получения свойств документа в объекты скриптов области задач

В случае всех предыдущих функций из этой серии руководств вы ставили в очередь команды для *записи* данных в документ Office. Каждая функция заканчивалась вызовом метода `context.sync()`, который отправляет поставленные в очередь команды документу для выполнения. Но код, который вы добавили на последнем этапе, вызывает свойство `originalRange.text`, и в этом заключается существенное отличие от ранее написанных функций, так как `originalRange` является лишь объектом прокси, существующим в скрипте вашей области задач. В нем нет сведений о фактическом тексте диапазона в документе, поэтому его свойство `text` может не содержать настоящего значения. Необходимо сначала получить из документа текстовое значение диапазона, а затем задать с его помощью значение для свойства `originalRange.text`. Только после этого можно будет вызвать метод `originalRange.text` без исключения. Процесс получения делится на три этапа:

   1. Добавление в очередь команды для загрузки (т. е. получения) свойств, которые должен прочесть ваш код.
   2. Вызов метода `sync` объекта контекста, чтобы можно было отправить документу находящуюся в очереди команду для выполнения, а также для возврата запрошенных данных.
   3. Метод `sync` асинхронный, поэтому его выполнение должно быть завершено до того, как код вызовет полученные свойства.

Эти три действия должны выполняться каждый раз, когда коду нужно *считывать* данные из документа Office.

1. Замените `TODO2` на приведенный ниже код.
  
    ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO4: Move the doc.body.insertParagraph line here.
    
            }
        )
            // TODO5: Move the final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has 
            //        been queued.
    ``` 

2. Для двух операторов `return` не может использоваться один путь кода, который не разветвляется, поэтому удалите последнюю строку `return context.sync();` в конце метода `Word.run`. Последний метод `context.sync` будет добавлен позже в этом руководстве. 
3. Вырежьте строку `doc.body.insertParagraph` и вставьте ее вместо заполнителя `TODO4`. 
4. Замените `TODO5` на приведенный ниже код. Обратите внимание:
   - Передача метода `sync` в функцию `then` гарантирует, что он не будет выполняться, пока логика `insertParagraph` не будет поставлена в очередь.
   - Метод `then` вызывает любую функцию, которая ему передана. Не нужно вызывать `sync` дважды, поэтому уберите "()" в конце вызова context.sync.

    ```js
    .then(context.sync);
    ```

Когда все будет готово, функция должна выглядеть так:

  
```js
function insertTextIntoRange() {
    Word.run(function (context) {
        
        const doc = context.document;
        const originalRange = doc.getSelection();
        originalRange.insertText(" (C2R)", "End");

        originalRange.load("text");
        return context.sync()
            .then(function() {        
                        doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                                                "End");            
                }
            )
            .then(context.sync);
    })
    .catch(function (error) {
        console.log("Error: " + error);
        if (error instanceof OfficeExtension.Error) {
            console.log("Debug info: " + JSON.stringify(error.debugInfo));
        }
    });
}
``` 

## <a name="add-text-between-ranges"></a>Добавление текста между диапазонами

1. Откройте файл index.html.
2. Под элементом `div`, содержащим кнопку `insert-text-into-range`, добавьте следующую разметку:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="insert-text-outside-range">Add Version Info</button>            
    </div>
    ```

3. Откройте файл app.js.

4. Под строкой, назначающей обработчик нажатия кнопки `insert-text-into-range`, добавьте следующий код:

    ```js
    $('#insert-text-outside-range').click(insertTextBeforeRange);
    ```

5. Добавьте приведенную ниже функцию под функцией `insertTextIntoRange`.

    ```js
    function insertTextBeforeRange() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to insert a new range before the 
            //        selected range.

            // TODO2: Load the text of the original range and sync so that the 
            //        range text can be read and inserted.

        })
        .catch(function (error) {
            console.log("Error: " + error);
            if (error instanceof OfficeExtension.Error) {
                console.log("Debug info: " + JSON.stringify(error.debugInfo));
            }
        });
    }
    ``` 

6. Замените `TODO1` на приведенный ниже код. Обратите внимание:
   - Этот метод предназначен для добавления диапазона с текстом "Office 2019, " перед диапазоном с текстом "Office 365". Для простоты предполагается, что такая строка существует и пользователь выделил ее.
   - Первый параметр метода `Range.insertText` — это добавляемая строка.
   - Второй параметр указывает, в каком месте диапазона требуется вставить дополнительный текст. Дополнительные сведения о вариантах расположения см. выше в описании функции `insertTextIntoRange`.

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("Office 2019, ", "Before");
    ``` 

7. Замените `TODO2` на приведенный ниже код. 
 
     ```js
    originalRange.load("text");
    return context.sync()
        .then(function() {

                // TODO3: Queue commands to insert the original range as a
                //        paragraph at the end of the document.
    
                }
            )

            // TODO4: Make a final call of context.sync here and ensure
            //        that it does not run until the insertParagraph has 
            //        been queued.
    ``` 

8. Замените `TODO3` на приведенный ниже код. Этот абзац покажет, что новый текст ***не*** входит в исходный выделенный диапазон. Исходный диапазон по-прежнему содержит такой же текст, как и когда он был выделен.
 
    ```js
    doc.body.insertParagraph("Current text of original range: " + originalRange.text,
                             "End");
    ``` 

9. Замените `TODO4` на приведенный ниже код.

    ```js
    .then(context.sync);
    ```


## <a name="replace-the-text-of-a-range"></a>Замена текста диапазона

1. Откройте файл index.html.
2. Под элементом `div`, содержащим кнопку `insert-text-outside-range`, добавьте следующую разметку:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="replace-text">Change Quantity Term</button>            
    </div>
    ```

3. Откройте файл app.js.

4. Под строкой, назначающей обработчик нажатия кнопки `insert-text-outside-range`, добавьте следующий код:

    ```js
    $('#replace-text').click(replaceText);
    ```

5. Добавьте приведенную ниже функцию под функцией `insertTextBeforeRange`.

    ```js
    function replaceText() {
        Word.run(function (context) {
             
            // TODO1: Queue commands to replace the text.

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

6. Замените `TODO1` на приведенный ниже код. Обратите внимание, что этот метод предназначен для замены строки "several" на строку "many". Для простоты предполагается, что такая строка существует и пользователь выделил ее.

    ```js
    const doc = context.document;
    const originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace"); 
    ``` 

## <a name="test-the-add-in"></a>Тестирование надстройки

1. Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера. Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.

     > [!NOTE]
     > Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу. Для этого необходимо завершить процесс сервера, чтобы появился запрос и вы могли ввести команду сборки. После сборки перезапустите сервер. Для этого выполните указанные ниже действия.

2. Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в более раннюю версию JavaScript, поддерживаемую всеми ведущими приложениями, в которых могут работать надстройки Office.
3. Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.
4. Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.
5. В области задач нажмите кнопку **Insert Paragraph** (Вставить абзац), чтобы убедиться, что в начале документа есть абзац.
6. Выделите какой-нибудь текст. Лучше всего выбрать фразу "Click-to-Run". *Будьте осторожны, чтобы не выделить пробел в начале или конце фразы.*
7. Нажмите кнопку **Insert Abbreviation** (Вставить аббревиатуру). Обратите внимание на добавленную строку " (C2R)". Кроме того, обратите внимание, что в конце документа добавлен новый абзац со всем развернутым текстом, так как новая строка была добавлена к имеющемуся диапазону.
8. Выделите какой-нибудь текст. Лучше всего выбрать фразу "Office 365". *Будьте осторожны, чтобы не выделить пробел в начале или конце фразы.*
9. Нажмите кнопку **Add Version Info** (Добавить сведения о версии). Обратите внимание, что между строками "Office 2016" и "Office 365" вставлена строка "Office 2019, ". Кроме того, обратите внимание, что в конце документа появился новый абзац, содержащий только изначально выделенный текст, так как новая строка стала новым диапазоном, а не была добавлена к существующему.
10. Выделите какой-нибудь текст. Лучше всего выделить слово "several". *Будьте осторожны, чтобы не выделить пробел в начале или конце фразы.*
11. Нажмите кнопку **Change Quantity Term** (Изменить числительное). Обратите внимание, что слово "many" заменило выделенный текст.

    ![Руководство по Word: добавленный и замененный текст](../images/word-tutorial-text-replace.png)
