# <a name="build-an-excel-add-in-using-vue"></a>Создание надстройки Excel с помощью Vue

Из этой статье вы узнаете, как создать надстройку Excel, используя Vue и API JavaScript для Excel.

## <a name="prerequisites"></a>Необходимые компоненты

- Установите [Vue CLI](https://github.com/vuejs/vue-cli) глобально.

    ```bash
    npm install -g vue-cli
    ```

- Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-vue-app"></a>Создание нового приложения Vue

Используйте Vue CLI, чтобы создать новое приложение Vue. Используя терминал, выполните приведенную ниже команду и ответьте на вопросы, как описано ниже.

```bash
vue init webpack my-add-in
```

Отвечая на вопросы, появляющиеся при выполнении предыдущей команды, переопределите стандартные ответы на 3 указанных ниже вопроса. Вы можете оставить стандартные ответы на все остальные вопросы.

- ****Install vue-router?** (Установить vue-router?)** `No`
- ****Set up unit tests?** `No` (Настроить модульные тесты?)** `No`
- ****Setup e2e tests with Nightwatch?** (Настроить тесты e2e с помощью Nightwatch?)** `No`

![Вопросы Vue CLI](../images/vue-cli-prompts.png)

## <a name="generate-the-manifest-file"></a>Создание файла манифеста

У каждой надстройки должен быть файл манифеста, в нем определяются ее параметры и возможности.

1. Перейдите к папке приложения.

    ```bash
    cd my-add-in
    ```

2. Используя генератор Yeoman, создайте файл манифеста для надстройки. Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.

    ```bash
    yo office 
    ```

    - **Выберите тип проекта:** `Office Add-in containing the manifest only`
    - **Как вы хотите назвать надстройку?:** `My Office Add-in`
    - **Какое клиентское приложение Office должно поддерживаться?:** `Excel`

    После завершения работы мастера вы сможете создать файл манифеста и файл ресурсов для создания вашего проекта.

    ![Генератор Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > Если вам будет предложено переписать файл **package.json**, выберите **No** (Нет).

## <a name="secure-the-app"></a>Защита приложения

[!include[HTTPS guidance](../includes/https-guidance.md)]

Чтобы включить HTTPS для своего приложения, откройте файл **package.json** в корневой папке проекта, добавьте флаг `--https` в скрипт `dev` и сохраните файл.

```json
"dev": "webpack-dev-server --https --inline --progress --config build/webpack.dev.conf.js"
```

## <a name="update-the-app"></a>Обновление приложения

1. В редакторе кода откройте файл манифеста (т. е. файл в корневом каталоге приложения, имя которого заканчивается на "manifest.xml"). Замените все вхождения `https://localhost:3000` на `https://localhost:8080` и сохраните файл.

2. Откройте файл **index.html**, добавьте тег `<script>` сразу перед тегом `</head>` и сохраните.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

3. Откройте файл **src/main.js** и *удалите* следующий блок кода:

    ```js
    new Vue({
        el: '#app',
        components: {App},
        template: '<App/>'
    })
    ```
    
    Затем добавьте приведенный ниже код в этом же месте и сохраните файл. 
                                                         
    ```js
    const Office = window.Office
    Office.initialize = () => {
      new Vue({
        el: '#app',
        components: {App},
        template: '<App/>'
      })
    }
    ```

4. Откройте файл **src/App.vue**, замените его содержимое приведенным ниже кодом, добавьте разрыв строки в конце (т. е. после тега `</style>`) и сохраните файл. 

    ```html
    <template>
    <div id="app">
        <div id="content">
        <div id="content-header">
            <div class="padding">
            <h1>Welcome</h1>
            </div>
        </div>
        <div id="content-main">
            <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br/>
            <h3>Try it out</h3>
            <button @click="onSetColor">Set color</button>
            </div>
        </div>
        </div>
    </div>
    </template>

    <script>
    export default {
      name: 'App',
      methods: {
        onSetColor () {
          window.Excel.run(async (context) => {
            const range = context.workbook.getSelectedRange()
            range.format.fill.color = 'green'
            await context.sync()
          })
        }
      }
    }
    </script>

    <style>
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
    </style>
    ```

## <a name="start-the-dev-server"></a>Запуск сервера разработки

1. Используя терминал, выполните приведенную ниже команду, чтобы запустить сервер разработки.

    ```bash
    npm start
    ```

2. В веб-браузере перейдите по адресу `https://localhost:8080`. Если появится сообщение, что сертификат сайта не является доверенным, сделайте так, чтобы компьютер ему доверял. 

3. После того как браузер загрузит страницу надстройки без ошибок сертификата, вы можете протестировать надстройку. 

## <a name="try-it-out"></a>Проверка

1. Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.

    - Windows[](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Office Online[](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad и Mac[](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

2. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. Выберите любой диапазон ячеек на листе.

4. В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку Excel с помощью Vue! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.

> [!div class="nextstepaction"]
> [Руководство по надстройкам Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>См. также

* [Руководство по надстройкам Excel](../tutorials/excel-tutorial-create-table.md)
* [Основные понятия API JavaScript для Excel](../excel/excel-add-ins-core-concepts.md)
* [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Справочник по API JavaScript для Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
