---
title: Руководство по надстройкам Word
description: В этом руководстве показано создание надстройки Word, которая вставляет (и заменяет) диапазоны текста, абзацы, изображения, HTML-код, таблицы и элементы управления контентом. Вы также узнаете, как форматировать текст, вставлять и заменять содержимое в элементах управления контентом.
ms.date: 12/31/2018
ms.prod: word
ms.topic: tutorial
localization_priority: Normal
ms.openlocfilehash: 019329db156e63148a047466b9b3770128cb7fbf
ms.sourcegitcommit: 33dcf099c6b3d249811580d67ee9b790c0fdccfb
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/05/2019
ms.locfileid: "29742403"
---
# <a name="tutorial-create-a-word-task-pane-add-in"></a>Учебник: Создание надстройки области задач Word

С помощью данного учебника вы сможете создать надстройку области задач Word, которая выполняет следующие действия:

> [!div class="checklist"]
> * Вставляет диапазон текста
> * Форматирует текст
> * Заменяет и вставляет текст в различных расположениях
> * Вставляет изображения, HTML-код и таблицы
> * Создает и обновляет элементы управления содержимым 

## <a name="prerequisites"></a>Необходимые компоненты

Для работы с этим руководством необходимо установить указанные ниже компоненты. 

- Word 2016, версия 1711 (сборка 8730.1000 "нажми и работай") или более поздняя. Чтобы установить эту версию, необходимо быть участником программы предварительной оценки Office. [Дополнительные сведения](https://products.office.com/office-insider?tab=tab-1)

- [Node](https://nodejs.org/en/) 

- [Git Bash](https://git-scm.com/downloads) (или другой клиент Git)

## <a name="create-your-add-in-project"></a>Создание проекта надстройки

Выполните указанные ниже действия для создания проекта надстройки Word, который будет использоваться в качестве основы для этого учебника.

1. Клонируйте репозиторий GitHub [Word-Add-in-Tutorial](https://github.com/OfficeDev/Word-Add-in-Tutorial).

2. Откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.

3. Выполните команду `npm install`, чтобы установить инструменты и библиотеки, указанные в файле package.json. 

4. Сделайте так, чтобы операционная система компьютера разработки доверяла сертификату. Для этого выполните действия, описанные в [этой статье](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

## <a name="insert-a-range-of-text"></a>Вставка диапазона текста

На этом этапе руководства мы программным способом проверим, поддерживает ли надстройка текущую версию Word, установленную у пользователя, а затем вставим абзац в документ.

### <a name="code-the-add-in"></a>Написание кода надстройки

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

7. Замените `TODO3` приведенным ниже кодом. Примечание.

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

   - Первый параметр метода `insertParagraph` — это текст нового абзаца.

   - Второй параметр — расположение в основном тексте, где будет вставлен абзац. Другие варианты вставки абзаца, родительским объектом которого является основной текст, — End и Replace.

    ```js
    var docBody = context.document.body;
    docBody.insertParagraph("Office has several versions, including Office 2016, Office 365 Click-to-Run, and Office Online.",
                            "Start");
    ```

### <a name="test-the-add-in"></a>Тестирование надстройки

1. Откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.

2. Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в более раннюю версию JavaScript, поддерживаемую всеми ведущими приложениями, в которых могут работать надстройки Office.

3. Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.

4. Загрузите неопубликованную надстройку одним из следующих способов:

    - [Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)

    - [Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-in-office-online)

    - [iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

5. В меню **Главная** в Word выберите пункт **Показать область задач**.

6. В области задач нажмите кнопку **Insert Paragraph** (Вставить абзац).

7. Внесите изменение в абзац.

8. Снова нажмите кнопку **Insert Paragraph**. Обратите внимание, что новый абзац находится над предыдущим, так как метод `insertParagraph` вставляет текст в начале основного текста документа.

    ![Руководство по Word: вставка абзаца](../images/word-tutorial-insert-paragraph.png)

## <a name="format-text"></a>Форматирование текста

На этом этапе учебника вы сможете применить встроенный стиль к тексту, использовать пользовательский стиль для текста и изменить шрифт текста.

### <a name="apply-a-built-in-style-to-text"></a>Применение встроенного стиля к тексту

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
    var firstParagraph = context.document.body.paragraphs.getFirst();
    firstParagraph.styleBuiltIn = Word.Style.intenseReference;
    ``` 

### <a name="apply-a-custom-style-to-text"></a>Применение пользовательского стиля к тексту

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

6. Замените `TODO1` на приведенный ниже код. Обратите внимание, что этот код применяет пользовательский стиль, который еще не существует. Мы создадим стиль с именем **MyCustomStyle** во время [тестирования настройки](#test-the-add-in).

    ```js
    var lastParagraph = context.document.body.paragraphs.getLast();
    lastParagraph.style = "MyCustomStyle";
    ``` 

### <a name="change-the-font-of-text"></a>Изменение шрифта для текста

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

6. Замените `TODO1` на приведенный ниже код. Обратите внимание, что этот код получает ссылку на второй абзац с помощью метода `ParagraphCollection.getFirst`, привязанного к методу `Paragraph.getNext`.

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    secondParagraph.font.set({
            name: "Courier New",
            bold: true,
            size: 18
        });
    ``` 

### <a name="test-the-add-in"></a>Тестирование надстройки

1. Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши Ctrl+C, чтобы остановить работу веб-сервера. Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.

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

## <a name="replace-text-and-insert-text"></a>Замена текста и добавление текста

На этом этапе руководства мы добавим текст в выбранные диапазоны текста и за их пределами, а также заменим текст выбранного диапазона.

### <a name="add-text-inside-a-range"></a>Добавление текста в диапазон

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

   - Первый параметр метода `Range.insertText` — это строка, вставляемая в объект `Range`.

   - Второй параметр указывает, в каком месте диапазона требуется вставить дополнительный текст. Помимо значения End, можно использовать значения Start, Before, After и Replace. 

   - Разница между значениями End и After состоит в том, что End вставляет новый текст в конце имеющегося диапазона, а After создает новый диапазон со строкой и вставляет его после имеющегося. Аналогично, Start вставляет текст в начале имеющегося диапазона, а Before вставляет новый диапазон. Replace заменяет текст существующего диапазона на строку из первого параметра.

   - На одном из предыдущих этапов руководства вы могли заметить, что в методах insert* объекта body нет параметров Before и After. Это связано с тем, что содержимое невозможно добавлять за пределами основного текста документа.

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText(" (C2R)", "End");
    ```

8. Пропустим заполнитель `TODO2` до следующего этапа. Замените `TODO3` на приведенный ниже код. Он похож на код, созданный на первом этапе руководства, но теперь мы вставляем новый абзац в конце, а не в начале документа. Новый абзац покажет, что новый текст теперь входит в исходный диапазон.

    ```js
    doc.body.insertParagraph("Original range: " + originalRange.text, "End");
    ```

### <a name="add-code-to-fetch-document-properties-into-the-task-panes-script-objects"></a>Добавление кода для получения свойств документа в объекты скриптов области задач

В случае всех предыдущих функций из этой серии руководств вы ставили в очередь команды для *записи* данных в документ Office. Каждая функция заканчивалась вызовом метода `context.sync()`, который отправляет поставленные в очередь команды документу для выполнения. Но код, который вы добавили на последнем этапе, вызывает свойство `originalRange.text`, и в этом заключается существенное отличие от ранее написанных функций, так как `originalRange` является лишь объектом прокси, существующим в скрипте вашей области задач. В нем нет сведений о фактическом тексте диапазона в документе, поэтому его свойство `text` может не содержать настоящего значения. Необходимо сначала получить из документа текстовое значение диапазона, а затем задать с его помощью значение для свойства `originalRange.text`. Только после этого можно будет вызвать метод `originalRange.text` без исключения. Процесс получения делится на три этапа:

   1. Добавление в очередь команды для загрузки (т. е. получения) свойств, которые должен прочесть ваш код.

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

Когда все будет готово, функция должна будет выглядеть так:

```js
function insertTextIntoRange() {
    Word.run(function (context) {

        var doc = context.document;
        var originalRange = doc.getSelection();
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

### <a name="add-text-between-ranges"></a>Добавление текста между диапазонами

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

   - Этот метод предназначен для добавления диапазона с текстом "Office 2019, " перед диапазоном с текстом "Office 365". Для простоты предполагается, что такая строка существует и пользователь выделил ее.

   - Первый параметр метода `Range.insertText` — это добавляемая строка.

   - Второй параметр указывает, в каком месте диапазона требуется вставить дополнительный текст. Дополнительные сведения о вариантах расположения см. выше в описании функции `insertTextIntoRange`.

    ```js
    var doc = context.document;
    var originalRange = doc.getSelection();
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

### <a name="replace-the-text-of-a-range"></a>Замена текста диапазона

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
    var doc = context.document;
    var originalRange = doc.getSelection();
    originalRange.insertText("many", "Replace");
    ```

### <a name="test-the-add-in"></a>Тестирование надстройки

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

## <a name="insert-images-html-and-tables"></a>Вставка изображений, HTML-кода и таблиц

На этом этапе руководства мы рассмотрим вставку изображений, HTML-кода и таблиц в документ.

### <a name="insert-an-image"></a>Вставка изображения

1. Откройте проект в редакторе кода.

2. Откройте файл index.html.

3. Под элементом `div`, содержащим кнопку `replace-text`, добавьте следующую разметку:

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-image">Insert Image</button>
    </div>
    ```

4. Откройте файл app.js.

5. Добавьте приведенную ниже строку сразу под строкой use-strict в верхней части файла. Эта строка импортирует переменную из другого файла. Переменная представляет собой строку с кодировкой Base 64, кодирующую изображение. Чтобы просмотреть закодированную строку, откройте файл base64Image.js в корневой папке проекта.

    ```js
    import { base64Image } from "./base64Image";
    ```

6. Под строкой, назначающей обработчик нажатия кнопки `replace-text`, добавьте следующий код:

    ```js
    $('#insert-image').click(insertImage);
    ```

7. Добавьте приведенную ниже функцию под функцией `replaceText`.

    ```js
    function insertImage() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert an image.

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

8. Замените `TODO1` на приведенный ниже код. Обратите внимание, что эта строка вставляет изображение с кодировкой Base 64 в конце документа. У объекта `Paragraph` также есть метод `insertInlinePictureFromBase64` и другие методы `insert*`. Пример представлен в следующем разделе, посвященном вставке HTML.

    ```js
    context.document.body.insertInlinePictureFromBase64(base64Image, "End");
    ```

### <a name="insert-html"></a>Вставка HTML

1. Откройте файл index.html.

2. Под элементом `div`, содержащим кнопку `insert-image`, добавьте следующую разметку:

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-html">Insert HTML</button>
    </div>
    ```

3. Откройте файл app.js.

4. Под строкой, назначающей обработчик нажатия кнопки `insert-image`, добавьте следующий код:

    ```js
    $('#insert-html').click(insertHTML);
    ```

5. Добавьте приведенную ниже функцию под функцией `insertImage`.

    ```js
    function insertHTML() {
        Word.run(function (context) {

            // TODO1: Queue commands to insert a string of HTML.

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

6. Замените `TODO1` на приведенный ниже код. Обратите внимание:

   - Первая строка добавляет пустой абзац в конце документа. 

   - Вторая команда вставляет строку HTML-кода в конце абзаца. В частности, вставляются два абзаца, в одном из которых используется шрифт Verdana, а в другом — стандартный стиль документа Word. (Как видно по вышеописанному методу `insertImage`, у объекта `context.document.body` также есть методы `insert*`).

    ```js
    var blankParagraph = context.document.body.paragraphs.getLast().insertParagraph("", "After");
    blankParagraph.insertHtml('<p style="font-family: verdana;">Inserted HTML.</p><p>Another paragraph</p>', "End");
    ```

### <a name="insert-a-table"></a>Вставка таблицы

1. Откройте файл index.html.

2. Под элементом `div`, содержащим кнопку `insert-html`, добавьте следующую разметку:

    ```html
    <div class="padding">
        <button class="ms-Button" id="insert-table">Insert Table</button>
    </div>
    ```

3. Откройте файл app.js.

4. Под строкой, назначающей обработчик нажатия кнопки `insert-html`, добавьте следующий код:

    ```js
    $('#insert-table').click(insertTable);
    ```

5. Добавьте приведенную ниже функцию под функцией `insertHTML`.

    ```js
    function insertTable() {
        Word.run(function (context) {

            // TODO1: Queue commands to get a reference to the paragraph
            //        that will proceed the table.

            // TODO2: Queue commands to create a table and populate it with data.

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

6. Замените `TODO1` на приведенный ниже код. Обратите внимание, что в этой строке используется метод `ParagraphCollection.getFirst`, чтобы получить ссылку на первый абзац, а затем — метод `Paragraph.getNext`, чтобы получить ссылку на второй абзац.

    ```js
    var secondParagraph = context.document.body.paragraphs.getFirst().getNext();
    ```

7. Замените `TODO2` на приведенный ниже код. Обратите внимание:

   - Первые два параметра метода `insertTable` задают количество строк и столбцов.

   - Третий параметр указывает, где вставить таблицу (в данном случае — после абзаца).

   - Четвертый параметр представляет собой двумерный массив, задающий значения ячеек таблицы.

   - К таблице применяется простой стиль по умолчанию, но метод `insertTable` возвращает объект `Table` со множеством элементов, некоторые из которых используются для настройки стиля таблицы.

    ```js
    var tableData = [
            ["Name", "ID", "Birth City"],
            ["Bob", "434", "Chicago"],
            ["Sue", "719", "Havana"],
        ];
    secondParagraph.insertTable(3, 3, "After", tableData);
    ```

### <a name="test-the-add-in"></a>Тестирование надстройки

1. Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера. Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.

     > [!NOTE]
     > Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу. Для этого необходимо завершить процесс сервера, чтобы появился запрос и вы могли ввести команду сборки. После сборки перезапустите сервер. Для этого выполните указанные ниже действия.

2. Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в более раннюю версию JavaScript, поддерживаемую всеми ведущими приложениями, в которых могут работать надстройки Office.

3. Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.

4. Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.

5. В области задач нажмите кнопку **Insert Paragraph** (Вставить абзац) не менее трех раз, чтобы убедиться, что в документе есть несколько абзацев.

6. Нажмите кнопку **Insert Image** (Вставить изображение) и обратите внимание, что изображение вставляется в конце документа.

7. Нажмите кнопку **Insert HTML** (Вставить HTML) и обратите внимание, что в конце документа вставляются два абзаца, в первом из которых используется шрифт Verdana.

8. Нажмите кнопку **Insert Table** (Вставить таблицу) и обратите внимание, что после второго абзаца вставляется таблица.

    ![Руководство по Word: вставка изображения, HTML-кода и таблицы](../images/word-tutorial-insert-image-html-table.png)

## <a name="create-and-update-content-controls"></a>Создание и обновление элементов управления содержимым

На этом этапе руководства мы рассмотрим создание элементов управления форматированным текстом в документе, а также вставку и замену содержимого этих элементов.

> [!NOTE]
> Существует несколько типов элементов управления содержимым, которые можно добавить в документ Word через пользовательский интерфейс. Однако в настоящее время Word.js поддерживает только элементы управления форматированным текстом.
>
> Прежде чем приступать к этому этапу руководства, рекомендуем создать элементы управления форматированным текстом и управлять ими через пользовательский интерфейс Word, чтобы получить представление об этих элементах и их свойствах. Дополнительные сведения см. в статье [Создание форм, предназначенных для заполнения или печати в приложении Word](https://support.office.com/article/create-forms-that-users-complete-or-print-in-word-040c5cc1-e309-445b-94ac-542f732c8c8b).

### <a name="create-a-content-control"></a>Создание элемента управления содержимым

1. Откройте проект в редакторе кода.

2. Откройте файл index.html.

3. Под элементом `div`, содержащим кнопку `replace-text`, добавьте следующую разметку:

    ```html
    <div class="padding">
        <button class="ms-Button" id="create-content-control">Create Content Control</button>
    </div>
    ```

4. Откройте файл app.js.

5. Под строкой, назначающей обработчик нажатия кнопки `insert-table`, добавьте следующий код:

    ```js
    $('#create-content-control').click(createContentControl);
    ```

6. Добавьте приведенную ниже функцию под функцией `insertTable`.

    ```js
    function createContentControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to create a content control.

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

   - Этот код заключает фразу "Office 365" в элемент управления содержимым. Для простоты предполагается, что такая строка существует и пользователь выделил ее.

   - Свойство `ContentControl.title` задает видимый заголовок элемента управления содержимым.

   - Свойство `ContentControl.tag` задает тег, с помощью которого можно получить ссылку на элемент управления содержимым путем вызова метода `ContentControlCollection.getByTag`, который будет использоваться в последующей функции.

   - Свойство `ContentControl.appearance` задает внешний вид элемента управления. Значение Tags указывает, что элемент управления будет заключен в открывающие и закрывающие теги, а открывающий тег будет содержать заголовок элемента управления содержимым. Другие возможные значения: BoundingBox и None.

   - Свойство `ContentControl.color` задает цвет тегов или рамки ограничивающего прямоугольника.

    ```js
    var serviceNameRange = context.document.getSelection();
    var serviceNameContentControl = serviceNameRange.insertContentControl();
    serviceNameContentControl.title = "Service Name";
    serviceNameContentControl.tag = "serviceName";
    serviceNameContentControl.appearance = "Tags";
    serviceNameContentControl.color = "blue";
    ```

### <a name="replace-the-content-of-the-content-control"></a>Замена содержимого элемента управления

1. Откройте файл index.html.

2. Под элементом `div`, содержащим кнопку `create-content-control`, добавьте следующую разметку:

    ```html
    <div class="padding">
        <button class="ms-Button" id="replace-content-in-control">Rename Service</button>
    </div>
    ```

3. Откройте файл app.js.

4. Под строкой, назначающей обработчик нажатия кнопки `create-content-control`, добавьте следующий код:

    ```js
    $('#replace-content-in-control').click(replaceContentInControl);
    ```

5. Добавьте приведенную ниже функцию под функцией `createContentControl`.

    ```js
    function replaceContentInControl() {
        Word.run(function (context) {

            // TODO1: Queue commands to replace the text in the Service Name
            //        content control.

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

6. Замените `TODO1` приведенным ниже кодом. Обратите внимание:

    - Метод `ContentControlCollection.getByTag` возвращает значение `ContentControlCollection` для всех элементов управления контентом указанного тега. Чтобы получить ссылку на нужный элемент управления, используйте `getFirst`.

    ```js
    var serviceNameContentControl = context.document.contentControls.getByTag("serviceName").getFirst();
    serviceNameContentControl.insertText("Fabrikam Online Productivity Suite", "Replace");
    ```

### <a name="test-the-add-in"></a>Тестирование надстройки

1. Если окно Git Bash или системная командная строка с поддержкой Node.JS, открытые на предыдущем этапе руководства, все еще открыты, дважды нажмите клавиши CTRL+C, чтобы остановить работу веб-сервера. Если они закрыты, откройте окно Git Bash или системную командную строку с поддержкой Node.JS и перейдите к папке **Start** проекта.

     > [!NOTE]
     > Хотя сервер синхронизации браузера будет повторно загружать надстройку в области задач при каждом изменении любого файла (в том числе app.js), он не передает повторно код JavaScript, поэтому нужно будет снова выполнить команду сборки, чтобы изменения, внесенные в файл app.js, вступили в силу. Для этого необходимо завершить процесс сервера, чтобы появился запрос и вы могли ввести команду сборки. После сборки перезапустите сервер. Для этого выполните указанные ниже действия.

2. Выполните команду `npm run build`, чтобы преобразовать исходный код ES6 в более раннюю версию JavaScript, поддерживаемую всеми ведущими приложениями, в которых могут работать надстройки Office.

3. Выполните команду `npm start`, чтобы запустить веб-сервер, работающий на localhost.

4. Перезагрузите область задач. Для этого закройте ее, а затем выберите в меню **Главная** пункт **Показать область задач**, чтобы заново открыть надстройку.

5. В области задач нажмите кнопку **Insert Paragraph** (Вставить абзац), чтобы убедиться, что в начале документа есть абзац с фразой "Office 365".

6. Выделите фразу "Office 365" в добавленном абзаце, а затем нажмите кнопку **Create Content Control** (Создать элемент управления содержимым). Обратите внимание, что фраза заключена в теги с меткой Service Name.

7. Нажмите кнопку **Rename Service** (Переименовать службу) и обратите внимание, что текст элемента управления содержимым меняется на "Fabrikam Online Productivity Suite".

    ![Руководство по Word: создание элемента управления содержимым и изменение его текста](../images/word-tutorial-content-control.png)

## <a name="next-steps"></a>Дальнейшие действия

В этом руководстве вы создали надстройку области задач Word, которая вставляет и заменяет текст, изображения и другое содержимое в документе Word. Чтобы узнать больше о создании надстроек Word, перейдите к следующей статье:

> [!div class="nextstepaction"]
> [Обзор надстроек Word](../word/word-add-ins-programming-overview.md)
