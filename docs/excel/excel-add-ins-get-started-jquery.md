# <a name="build-an-excel-add-in-using-jquery"></a>Создание надстройки Excel с помощью jQuery

В этой статье описывается процесс создания надстройки Excel с помощью jQuery и API JavaScript для Excel.

## <a name="prerequisites"></a>Необходимые условия

Если это еще не сделано, необходимо глобально установить [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).

    ```bash
    npm install -g yo generator-office
    ```

## <a name="create-the-web-app"></a>Создание веб-приложения

1. Создайте на локальном диске папку и назовите ее **my-addin**. В ней вы будете создавать файлы для приложения.

2. Перейдите к папке приложения.

    ```bash
    cd my-addin
    ```

3. С помощью генератора Yeoman создайте файл манифеста для надстройки. Выполните приведенную ниже команду, а затем укажите ответы на вопросы, как показано на следующем снимке экрана:

    ```bash
    yo office
    ```
    ![Генератор Yeoman](../../images/yo-office-jquery.png)


4. В редакторе кода откройте файл **index.html** из корневой папки проекта. Этот файл содержит HTML-контент, который будет отображаться в области задач надстройки. 
 
5. Замените созданный тег `header` приведенной ниже частью кода.
 
    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    ```

6. Замените созданный тег `main` приведенной ниже частью кода и сохраните файл.

    ```html
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button class="ms-Button" id="set-color">Set color</button>
        </div>
    </div>
    ```

7. Откройте файл **app.js**, чтобы указать скрипт для надстройки. Замените созданное выражение с функцией, которая вызывается сразу, приведенным ниже кодом и сохраните файл.

    ```js
    (function () {
        "use strict";

        Office.initialize = function (reason) {
            $(document).ready(function () {
                $('#set-color').click(setColor);
            });
        };

        function setColor() {
            Excel.run(function (context) {
                var range = context.workbook.getSelectedRange();
                range.format.fill.color = 'green';

                return context.sync();
            }).catch(function (error) {
                console.log("Error: " + error);
                if (error instanceof OfficeExtension.Error) {
                    console.log("Debug info: " + JSON.stringify(error.debugInfo));
                }
            });
        }
    })();
    ```

8. Откройте файл **app.js**, чтобы указать настраиваемые стили для надстройки. Замените содержимое (кроме примечания об авторских правах) приведенным ниже кодом и сохраните файл.

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

## <a name="configure-the-manifest-file-and-sideload-the-add-in"></a>Настройка файла манифеста и загрузка неопубликованной надстройки

1. Откройте файл **my-office-add-in-manifest.xml**, чтобы определить параметры и возможности надстройки. 

2. Тег **ProviderName** содержит замещающее значение. Замените его значением `Microsoft`.

3. Параметр **DefaultValue** тега **DisplayName** содержит замещающее значение. Замените его значением `A task pane add-in for Excel`. 

4. Сохраните файл, но пока не закрывайте его.

## <a name="configure-to-use-http"></a>Настройка на использование HTTP

Веб-надстройки Office должны использовать протокол HTTPS, а не HTTP, даже во время разработки. Однако для быстрого запуска надстройки в этом кратком руководстве используется протокол HTTP. Чтобы сделать это возможным, выполните указанные ниже действия.

1. В файле манифеста **my-office-add-in-manifest.xml** замените все вхождения "https" на "http". Затем сохраните и закройте файл.

2. Откройте файл **bsconfig.json** в корневой папке проекта. Замените значение свойства **https** на `false`. Сохраните файл.


## <a name="try-it-out"></a>Проверка

1. Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.

    - Windows: [Загрузка неопубликованных надстроек Office в Windows для тестирования](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md).
    - Excel Online: [Загрузка неопубликованных надстроек Office в Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online).
    - iPad и Mac: [Загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md).

2. Откройте терминал bash в корневой папке проекта и выполните приведенную ниже команду, чтобы запустить сервер разработки.

    ```bash
    npm start
    ```

   > **Примечание.** Откроется окно браузера с надстройкой. Закройте это окно.

3. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Кнопка надстройки Excel](../../images/excel_quickstart_addin_2a.png)

4. Выберите любой диапазон ячеек на листе.

5. В области задач нажмите кнопку **Color Me** (Раскрасить), чтобы сделать выбранный диапазон зеленым.

    ![Надстройка Excel](../../images/excel_quickstart_addin_2b.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку Excel с помощью jQuery! Теперь вы можете узнать больше об [основных понятиях](excel-add-ins-core-concepts.md), связанных с созданием надстроек Excel.

## <a name="additional-resources"></a>Дополнительные ресурсы

* [Основные понятия API JavaScript для Excel](excel-add-ins-core-concepts.md)
* [Изучайте фрагменты кода с помощью Script Lab](https://store.office.com/en-001/app.aspx?assetid=WA104380862&ui=en-US&rs=en-001&ad=US&appredirect=false)
* [Примеры кода надстроек Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Справочник по API JavaScript для Excel](../../reference/excel/excel-add-ins-reference-overview.md)
