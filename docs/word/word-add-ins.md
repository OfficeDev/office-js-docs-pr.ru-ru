# <a name="build-your-first-word-add-in"></a>Создание первой надстройки Word

_Область применения: Word 2016, Word для iPad, Word для Mac_

Надстройка Word работает в Word и может взаимодействовать с содержимым документа при помощи API JavaScript для Word, входящего в состав модели программирования надстроек Office, для расширения возможностей приложений Office. В этой модели программирования надстроек можно использовать любую платформу и любой язык для создания веб-приложения, в котором размещается расширение для Word, а затем определить его параметры и возможности с помощью [манифеста](../../docs/overview/add-in-manifests.md) надстройки.

В этой статье описывается процесс создания надстройки Word с помощью jQuery и API JavaScript для Word. 

> **Примечание.** Чтобы создать надстройку для Word 2013, необходимо использовать общий [API JavaScript для Office]( https://dev.office.com/docs/add-ins/word/word-add-ins-programming-overview#javascript-apis-for-word). Дополнительные сведения о платформах и различных доступных API см. в статье [Доступность ведущих приложений и платформ для надстроек Office](https://dev.office.com/add-in-availability). 

## <a name="create-the-web-app"></a>Создание веб-приложения 

1. Создайте на локальном диске папку и назовите ее **BoilerplateAddin**. В ней вы будете создавать файлы для приложения.

2. В папке приложения создайте файл с именем **home.html**, чтобы указать содержимое HTML, которое будет отображаться в области задач надстройки. В надстройке будут отображаться три кнопки, а при нажатии какой-либо из кнопок в документ добавляется стандартный текст. Добавьте приведенный ниже код в файл **home.html** и сохраните его.

    ```html
    <!DOCTYPE html>
    <html>
      <head>
        <meta charset="UTF-8" />
        <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
        <title>Boilerplate text app</title>
        <script src="https://ajax.aspnetcdn.com/ajax/jQuery/jquery-2.1.4.min.js"></script>
        <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js" type="text/javascript"></script>
        <script src="home.js" type="text/javascript"></script>
        </head>
        <body>
            <div>
                <h1>Welcome</h1>
            </div>
            <div>
                <p>This sample shows how to add boilerplate text to a document by using the Word JavaScript API.</p>
                <br />
                <h3>Try it out</h3>
                <button id="emerson">Add quote from Ralph Waldo Emerson</button>
                <button id="checkhov">Add quote from Anton Chekhov</button>
                <button id="proverb">Add Chinese proverb</button>
            </div>
            <h3><div id="supportedVersion"/></h3>
        </body>
    </html>
    ```

3. В папке приложения создайте файл с именем **home.js**, чтобы указать скрипт jQuery для надстройки. Этот скрипт содержит код инициализации, а также код, вносящий изменения в документ Word, вставляя текст при нажатии кнопки. Добавьте приведенный ниже код в файл **home.js** и сохраните его.

    ```javascript
    (function () {
        "use strict";

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

## <a name="create-the-manifest-file"></a>Создание файла манифеста

1. В папке приложения создайте файл с именем **BoilerplateManifest.xml**, чтобы определить параметры и возможности надстройки. Добавьте указанный ниже код в файл. 

    ```xml
    <?xml version="1.0" encoding="UTF-8"?>
        <OfficeApp xmlns="http://schemas.microsoft.com/office/appforoffice/1.1"
                xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance"
                xsi:type="TaskPaneApp">
            <Id>2b88100c-656e-4bab-9f1e-f6731d86e464</Id>
            <Version>1.0.0.0</Version>
            <ProviderName>Microsoft</ProviderName>
            <DefaultLocale>en-US</DefaultLocale>
            <DisplayName DefaultValue="Boilerplate content" />
            <Description DefaultValue="Insert boilerplate content into a Word document." />
            <Hosts>
                <Host Name="Document"/>
            </Hosts>
            <DefaultSettings>
                <SourceLocation DefaultValue="\\MyShare\boilerplate\home.html" />
            </DefaultSettings>
            <Permissions>ReadWriteDocument</Permissions>
        </OfficeApp>
    ```

2. Создайте GUID с помощью любого веб-генератора. Затем замените значение элемента **Id**, указанного на предыдущем этапе, этим GUID.

3. Сохраните файл манифеста.

## <a name="deploy-the-web-app-and-update-the-manifest"></a>Развертывание веб-приложения и обновление манифеста

1. Разверните веб-приложение (т. е. содержимое папки приложения) на нужном веб-сервере.

2. В локальной папке приложения откройте файл манифеста (**BoilerplateManifest.xml**). Измените значение атрибута в элементе **SourceLocation**, чтобы указать расположение файла **home.html** на веб-сервере, и сохраните файл.

## <a name="try-it-out"></a>Проверка

1. Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Word.

    - Windows: [Загрузка неопубликованных надстроек Office в Windows для тестирования](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Word Online: [Загрузка неопубликованных надстроек Office в Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad и Mac: [Загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).

2. В области задач в правой части экрана нажмите любую кнопку, чтобы добавить стандартный текст в документ.

![Изображение приложения Word с загруженной надстройкой, добавляющей стандартный текст.](../../images/boilerplateAddin.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку Word с помощью jQuery! Теперь вы можете узнать больше об [основных понятиях](word-add-ins-programming-overview.md), связанных с созданием надстроек Word.

## <a name="additional-resources"></a>Дополнительные ресурсы

* [Обзор надстроек Word](word-add-ins-programming-overview.md)
* [Изучение фрагментов кода с помощью Script Lab](https://store.office.com/en-001/app.aspx?assetid=WA104380862&ui=en-US&rs=en-001&ad=US&appredirect=false)
* [Примеры кода надстроек Word](http://dev.office.com/code-samples#?filters=word,office%20add-ins)
* [Справочник по API JavaScript для Word](../../reference/word/word-add-ins-reference-overview.md)