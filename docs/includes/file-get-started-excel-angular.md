# <a name="build-an-excel-add-in-using-angular"></a>Создание надстройки Excel с помощью Angular

В этой статье описывается процесс создания надстройки Excel с помощью Angular и API JavaScript для Excel.

## <a name="prerequisites"></a>Необходимые компоненты

- Посмотрите, что [необходимо для использования CLI для Angular](https://github.com/angular/angular-cli#prerequisites), и установите все недостающие компоненты.

- Глобально установите [CLI для Angular](https://github.com/angular/angular-cli). 

    ```bash
    npm install -g @angular/cli
    ```

- Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-angular-app"></a>Создание приложения Angular

Используйте угловые CLI, чтобы создать угловые приложения. Используя терминал, выполните следующую команду:

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file"></a>Создание файла манифеста

В файле манифеста надстройки определяются ее параметры и возможности.

1. Перейдите к папке приложения.

    ```bash
    cd my-addin
    ```

2. Используя генератор Yeoman, создайте файл манифеста для надстройки. Выполните приведенную ниже команду и ответьте на запросы, как показано ниже.

    ```bash
    yo office 
    ```

    - **Выберите тип проекта:** `Office Add-in containing the manifest only`
    - **Как вы хотите назвать надстройку?:** `My Office Add-in`
    - **Какое клиентское приложение Office вы хотели бы поддерживать?:** `Excel`

    После завершения работы мастера вы сможете создать файл манифеста и файл ресурсов для создания вашего проекта.

    ![Генератор Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > Если вам будет предложено переписать файл **package.json**, выберите **No** (Не перезаписывать).

## <a name="secure-the-app"></a>Защита приложения

[!include[HTTPS guidance](../includes/https-guidance.md)]

Для краткого руководства можно использовать сертификаты, которые предоставляют **Генератор Yeoman для надстроек Office**. Вы уже установили генератор глобально  (как часть **Необходимых компонентов** этого краткого руководства), поэтому вам просто нужно скопировать сертификаты из приложения глобальной установки в папку приложения. Следуюшие шаги описывают как выполнить этот процесс.

1. Используя терминал, выполните следующую команду, чтобы определить папку, в которую установлены глобальные библиотеки **npm**:

    ```bash 
    npm list -g 
    ``` 
    
    > [!TIP]    
    > Первая строка выходных данных, создаваемых этой командой, указывает папку, в которую установлены глобальные библиотеки **npm**.          
    
2. Используя проводник, перейдите к папке  `{global libraries folder}/node_modules/generator-office/generators/app/templates/js/base` . Из этого расположения скопируйте папку `certs` в буфер обмена.

3. Перейдите в корневую папку приложения Angular, созданную на шаге 1 предыдущего раздела, и вставьте папку `certs` из буфера обмена в эту папку.

## <a name="update-the-app"></a>Обновление приложения

1. В редакторе кода откройте файл **package.json** в корневой папке проекта. Измените скрипт `start`, чтобы указать, что сервер должен использовать SSL и порт 3000, и сохраните файл.

    ```json
    "start": "ng serve --ssl true --port 3000"
    ```

2. Откройте файл **.angular cli.json** в корневой папке проекта. Измените объект **defaults**, чтобы указать расположение сертификатов, и сохраните файл.

    ```json
    "defaults": {
      "styleExt": "css",
      "component": {},
      "serve": {
        "sslKey": "certs/server.key",
        "sslCert": "certs/server.crt"
      }
    }
    ```

3. Откройте файл **src/index.html**, добавьте тег `<script>` сразу перед тегом `</head>` и сохраните.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

4. Откройте файл **src/main.ts**, замените `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` приведенным ниже кодом и сохраните файл. 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

5. Откройте файл **src/polyfills.ts**, добавьте приведенную ниже строку кода над всеми имеющимися операторами `import` и сохраните файл.

    ```typescript
    import 'core-js/client/shim';
    ```

6. В файле **src/polyfills.ts** раскомментируйте приведенные ниже строки и сохраните файл.

    ```typescript
    import 'core-js/es6/symbol';
    import 'core-js/es6/object';
    import 'core-js/es6/function';
    import 'core-js/es6/parse-int';
    import 'core-js/es6/parse-float';
    import 'core-js/es6/number';
    import 'core-js/es6/math';
    import 'core-js/es6/string';
    import 'core-js/es6/date';
    import 'core-js/es6/array';
    import 'core-js/es6/regexp';
    import 'core-js/es6/map';
    import 'core-js/es6/weak-map';
    import 'core-js/es6/set';
    ```

7. Откройте файл **src/app/app.component.html**, замените его содержимое приведенным ниже кодом HTML и сохраните файл. 

    ```html
    <div id="content-header">
        <div class="padding">
            <h1>Welcome</h1>
        </div>
    </div>
    <div id="content-main">
        <div class="padding">
            <p>Choose the button below to set the color of the selected range to green.</p>
            <br />
            <h3>Try it out</h3>
            <button (click)="onSetColor()">Set color</button>
        </div>
    </div>
    ```

8. Откройте файл **src/app/app.component.css**, замените его содержимое приведенным ниже кодом CSS и сохраните файл.

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

9. Откройте файл **src/app/app.component.ts**, замените его содержимое приведенным ниже кодом и сохраните. 

    ```typescript
    import { Component } from '@angular/core';

    declare const Excel: any;

    @Component({
    selector: 'app-root',
    templateUrl: './app.component.html',
    styleUrls: ['./app.component.css']
    })
    export class AppComponent {
    onSetColor() {
        Excel.run(async (context) => {
        const range = context.workbook.getSelectedRange();
        range.format.fill.color = 'green';
        await context.sync();
        });
    }
    }
    ```

## <a name="start-the-dev-server"></a>Запуск сервера разработки

1. Используя терминал, выполните приведенную ниже команду, чтобы запустить сервер разработки.

    ```bash
    npm run start
    ```

2. В веб-браузере перейдите по адресу `https://localhost:3000`. Если появится сообщение о том, что сертификат сайта не является доверенным, укажите, что ему можно доверять. Дополнительные сведения см. в статье [Добавление самозаверяющего сертификата как доверенного корневого сертификата](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md).

    > [!NOTE]
    > Веб-браузер Chrome может продолжать показывать предупреждение о том, что рабочая станция не доверяет сертификату сайта даже после его [добавления в список доверенных корневых сертификатов](https://github.com/OfficeDev/generator-office/blob/master/src/docs/ssl.md). Вы можете игнорировать это предупреждение. Чтобы убедиться, что сертификат является доверенным, перейдите по адресу  `https://localhost:3000` в Internet Explorer или Microsoft Edge. 

3. После того как браузер загрузит страницу надстройки без ошибок сертификата, вы можете протестировать надстройку. 

## <a name="try-it-out"></a>Проверка

1. Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.

    - Windows: [Загрузка неопубликованных надстроек Office в Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Excel Online: [Загрузка неопубликованных надстроек Office в Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad и Mac: [Загрузка неопубликованных надстроек Office на iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

   
2. В Excel перейдите на вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. Выберите любой диапазон ячеек на листе.

4. В области задач нажмите кнопку **Set color**, чтобы задать цвет выбранного диапазона зеленым.

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Дальнейшие шаги

Поздравляем, вы успешно создали надстройку Excel с помощью Angular! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.

> [!div class="nextstepaction"]
> [Руководство по надстройкам Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>См. также

* [Руководство по надстройкам Excel](../tutorials/excel-tutorial-create-table.md)
* [Основные принципы программирования с использованием интерфейса API JavaScript для Excel](../excel/excel-add-ins-core-concepts.md)
* [Примеры кода надстроек Excel](https://developer.microsoft.com/office/gallery/?filterBy=Samples,Excel)
* [Справочник по API JavaScript для Excel](https://docs.microsoft.com/javascript/office/overview/excel-add-ins-reference-overview?view=office-js)
