На этом этапе руководства мы изменим шрифт текста и применим к нему как встроенные, так и пользовательские стили.

> [!NOTE]
> На этой странице описывается отдельный этап из руководства по надстройкам Word. Если вы перешли на эту страницу со страницы результатов поисковой системы или по другой прямой ссылке, перейдите на вводную страницу [руководства по надстройкам Word](../tutorials/word-tutorial.yml), чтобы начать обучение с самого начала.

## <a name="apply-a-built-in-style-to-text"></a>Применение встроенного стиля к тексту

1. Откройте проект в редакторе кода. 
2. Откройте файл index.html.
3. Под элементом `div`, содержащим кнопку `insert-paragraph`, добавьте следующую разметку:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-style">Apply Style</button>            
    </div>
    ```

4. Откройте файл app.js.

5. Под строкой, назначающей обработчик нажатия кнопки `insert-paragraph`, добавьте следующий код:

    ```js
    $('#apply-style').click(applyStyle);
    ```

6. Под функцией `insertParagraph` добавьте следующую функцию:

    ```js
    function applyStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to style text.

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

7. Замените `TODO1` на приведенный ниже код. Обратите внимание, что этот код применяет стиль к абзацу, но стили также можно применять к диапазонам текста.

    ```js
    const firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

## <a name="apply-a-custom-style-to-text"></a>Применение пользовательского стиля к тексту

1. Откройте файл index.html.
2. Под элементом `div`, содержащим кнопку `apply-style`, добавьте следующую разметку:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="apply-custom-style">Apply Custom Style</button>            
    </div>
    ```

3. Откройте файл app.js.

4. Под строкой, назначающей обработчик нажатия кнопки `apply-style`, добавьте следующий код:

    ```js
    $('#apply-custom-style').click(applyCustomStyle);
    ```

5. Добавьте приведенную ниже функцию под функцией `applyStyle`.

    ```js
    function applyCustomStyle() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply the custom style.

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

7. Замените `TODO1` на приведенный ниже код. Обратите внимание, что этот код применяет пользовательский стиль, который еще не существует. Мы создадим стиль с именем **MyCustomStyle** во время [тестирования настройки](#test-the-add-in).

    ```js
    const lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

## <a name="change-the-font-of-text"></a>Изменение шрифта для текста

1. Откройте файл index.html.
2. Под элементом `div`, содержащим кнопку `apply-custom-style`, добавьте следующую разметку:

    ```html
    <div class="padding">            
        <button class="ms-Button" id="change-font">Change Font</button>            
    </div>
    ```

3. Откройте файл app.js.

4. Под строкой, назначающей обработчик нажатия кнопки `apply-custom-style`, добавьте следующий код:

    ```js
    $('#change-font').click(changeFont);
    ```

5. Добавьте приведенную ниже функцию под функцией `applyCustomStyle`.

    ```js
    function changeFont() {
        Word.run(function (context) {
            
            // TODO1: Queue commands to apply a different font.

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

7. Замените `TODO1` на приведенный ниже код. Обратите внимание, что этот код получает ссылку на второй абзац с помощью метода `ParagraphCollection.getFirst`, привязанного к методу `Paragraph.getNext`.

    ```js
    const secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

## <a name="test-the-add-in"></a>Тестирование надстройки

1. Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера. Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.

     > [!NOTE]
     > Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу. Для этого необходимо завершить процесс сервера, чтобы появился запрос и вы могли ввести команду сборки. После сборки необходимо перезапустить сервер. Для этого выполните указанные ниже действия.

2. Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в более раннюю версию JavaScript, поддерживаемую всеми ведущими приложениями, в которых могут работать надстройки Office.
3. Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.   
4. Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.
5. Убедитесь, что в тексте есть по крайней мере три абзаца. Вы можете три раза нажать кнопку **Insert Paragraph** (Вставить абзац). *Внимательно проверьте, нет ли в конце документа пустого абзаца. Если он есть, удалите его.*
6. В Word создайте пользовательский стиль с именем "MyCustomStyle". Его форматирование может быть любым.
7. Нажмите кнопку **Apply Style** (Применить стиль). К первому абзацу будет применен встроенный стиль **Сильная ссылка**.
8. Нажмите кнопку **Apply Custom Style** (Применить пользовательский стиль). К последнему абзацу будет применен созданный вами стиль. Если ничего не происходит, возможно, последний абзац пуст. Если это так, добавьте в него какой-нибудь текст.
9. Нажмите кнопку **Change Font** (Изменить шрифт). Шрифт второго абзаца изменится на полужирный Courier New с размером 18.

    ![Руководство по Word: применение стилей и шрифта](../images/word-tutorial-apply-styles-and-font.png)
