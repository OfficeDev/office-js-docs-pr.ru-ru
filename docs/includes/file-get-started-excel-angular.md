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

Используйте Angular CLI, чтобы создать приложение Angular. В терминале выполните следующую команду:

```bash
ng new my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a>Создание файла манифеста и загрузка неопубликованной надстройки

В файле манифеста надстройки определяются ее параметры и возможности.

1. Перейдите к папке приложения.

    ```bash
    cd my-addin
    ```

2. С помощью генератора Yeoman создайте файл манифеста для надстройки. Выполните приведенную ниже команду и ответьте на вопросы, как показано на снимке экрана ниже.

    ```bash
    yo office
    ```
    - **Would you like to create a new subfolder for your project?:** `No` (Создать новую вложенную папку для проекта?)
    - **Как вы хотите назвать надстройку?:** `My Office Add-in`
    - **Какое клиентское приложение Office должно поддерживаться?:** `Excel`
    - **Would you like to create a new add-in?:** `No` (Создать новую надстройку?)

    Затем генератор предложит вам открыть файл **resource.html**. В нашем случае открывать его не обязательно, но можете заглянуть, если вам интересно! Выберите Yes (Да) или No (Нет), чтобы завершить работу мастера, и подождите, пока генератор закончит работу.

    ![Генератор Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > Если вам будет предложено переписать файл **package.json**, выберите **No** (не переписывать).

3. Откройте файл манифеста (т. е. файл в корневом каталоге приложения, имя которого заканчивается на "manifest.xml"). Замените все вхождения `https://localhost:3000` на `http://localhost:4200` и сохраните файл.

    > [!TIP]
    > Обязательно измените протокол на **http**, а номер порта — на **4200**.

4. Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.

    - [Windows](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - [Office Online](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - [iPad и Mac](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

## <a name="update-the-app"></a>Обновление приложения

1. Откройте **src/index.html**, добавьте тег `<script>` сразу перед тегом `</head>` и сохраните файл.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. Откройте **src/main.ts**, замените `platformBrowserDynamic().bootstrapModule(AppModule).catch(err => console.log(err));` приведенным ниже кодом и сохраните файл. 

    ```typescript 
    declare const Office: any;

    Office.initialize = () => {
    platformBrowserDynamic().bootstrapModule(AppModule)
        .catch(err => console.log(err));
    };
    ```

3. Откройте **src/polyfills.ts**, добавьте приведенную ниже строку кода над всеми имеющимися операторами `import` и сохраните файл.

    ```typescript
    import 'core-js/client/shim';
    ```

4. В файле **src/polyfills.ts** раскомментируйте приведенные ниже строки и сохраните файл.

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

5. Откройте **src/app/app.component.html**, замените его содержимое приведенным ниже кодом HTML и сохраните файл. 

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

6. Откройте **src/app/app.component.css**, замените его содержимое приведенным ниже кодом CSS и сохраните файл.

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

7. Откройте **src/app/app.component.ts**, замените его содержимое приведенным ниже кодом и сохраните файл. 

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

## <a name="try-it-out"></a>Проверка

1. Выполните в терминале приведенную ниже команду, чтобы запустить сервер разработки.

    ```bash
    npm start
    ```
   
2. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2a.png)

3. Выберите любой диапазон ячеек на листе.

4. В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку Excel с помощью Angular! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.

> [!div class="nextstepaction"]
> [Руководство по надстройкам Excel](../tutorials/excel-tutorial-create-table.md)

## <a name="see-also"></a>См. также

* [Руководство по надстройкам Excel](../tutorials/excel-tutorial-create-table.md)
* [Основные понятия API JavaScript для Excel](../excel/excel-add-ins-core-concepts.md)
* [Примеры кода надстроек Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Справочник по API JavaScript для Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)

