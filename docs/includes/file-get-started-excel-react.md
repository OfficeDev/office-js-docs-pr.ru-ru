# <a name="build-an-excel-add-in-using-react"></a>Создание надстройки Excel с помощью React

В этой статье описывается процесс создания надстройки Excel с помощью React и API JavaScript для Excel.

## <a name="environment"></a>Среда

- **Классическое приложение Office.** Убедитесь, что у вас установлена ​​последняя версия Office. Команды надстроек требуют сборку 16.0.6769.0000 или более позднюю (рекомендуется сборка **16.0.6868.0000**). Узнайте, как [установить последнюю версию приложений Office](http://aka.ms/latestoffice). 
 
- **Office Online.** Не требуется выполнять дополнительную настройку. Обратите внимание, что поддержка команд в Office Online для рабочих и учебных учетных записей предоставляется в тестовом режиме.

## <a name="prerequisites"></a>Необходимые компоненты

- Глобально установите [Create React App](https://github.com/facebookincubator/create-react-app).

    ```bash
    npm install -g create-react-app
    ```

- Глобально установите последнюю версию [Yeoman](https://github.com/yeoman/yo) и [генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).

    ```bash
    npm install -g yo generator-office
    ```

## <a name="generate-a-new-react-app"></a>Создание приложения React

Создайте приложение React с помощью Create React App. В терминале выполните следующую команду:

```bash
create-react-app my-addin
```

## <a name="generate-the-manifest-file-and-sideload-the-add-in"></a>Создание файла манифеста и загрузка неопубликованной надстройки

Каждой надстройке необходим файл манифеста, чтобы определить ее параметры и возможности.

1. Перейдите к папке приложения.

    ```bash
    cd my-addin
    ```

2. С помощью генератора Yeoman создайте файл манифеста для надстройки. Выполните приведенную ниже команду и ответьте на вопросы, как показано на следующем снимке экрана:

    ```bash
    yo office
    ```

    - ****Would you like to create a new subfolder for your project?:** `No` (Создать новую вложенную папку для проекта?)** `No`
    - **Как вы хотите назвать надстройку?:** `My Office Add-in`
    - **Какое клиентское приложение Office должно поддерживаться?:** `Excel`
    - ****Would you like to create a new add-in?:** `No` (Создать новую надстройку?)** `No`

    Затем генератор предложит вам открыть файл **resource.html**. В нашем случае открывать его не обязательно, но можете заглянуть, если вам интересно! Выберите Yes (Да) или No (Нет), чтобы завершить работу мастера, и подождите, пока генератор закончит работу.

    ![Генератор Yeoman](../images/yo-office.png)
    
    > [!NOTE]
    > Если вам будет предложено переписать файл **package.json**, выберите **No** (не переписывать).

3. Следуя указаниям для нужной платформы, загрузите неопубликованную надстройку в Excel.

    - Windows[](../testing/create-a-network-shared-folder-catalog-for-task-pane-and-content-add-ins.md)
    - Office Online[](../testing/sideload-office-add-ins-for-testing.md#sideload-an-office-add-in-on-office-online)
    - iPad и Mac[](../testing/sideload-an-office-add-in-on-ipad-and-mac.md)

## <a name="update-the-app"></a>Обновление приложения

1. Откройте **public/index.html**, добавьте тег `<script>` сразу перед тегом `</head>` и сохраните файл.

    ```html
    <script src="https://appsforoffice.microsoft.com/lib/1/hosted/office.js"></script>
    ```

2. Откройте **src/index.js**, замените `ReactDOM.render(<App />, document.getElementById('root'));` приведенным ниже кодом и сохраните файл. 

    ```typescript
    const Office = window.Office;
    
    Office.initialize = () => {
      ReactDOM.render(<App />, document.getElementById('root'));
    };
    ```

3. Откройте **src/App.js**, замените его содержимое приведенным ниже кодом и сохраните файл. 

    ```js
    import React, { Component } from 'react';
    import './App.css';

    class App extends Component {
      constructor(props) {
        super(props);

        this.onSetColor = this.onSetColor.bind(this);
      }

      onSetColor() {
        window.Excel.run(async (context) => {
          const range = context.workbook.getSelectedRange();
          range.format.fill.color = 'green';
          await context.sync();
        });
      }

      render() {
        return (
          <div id="content">
            <div id="content-header">
              <div className="padding">
                  <h1>Welcome</h1>
              </div>
            </div>
            <div id="content-main">
              <div className="padding">
                  <p>Choose the button below to set the color of the selected range to green.</p>
                  <br />
                  <h3>Try it out</h3>
                  <button onClick={this.onSetColor}>Set color</button>
              </div>
            </div>
          </div>
        );
      }
    }

    export default App;
    ```

4. Откройте **src/App.css**, замените его содержимое приведенным ниже кодом и сохраните файл. 

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

## <a name="try-it-out"></a>Проверка

1. Выполните в терминале приведенную ниже команду, чтобы запустить сервер разработки.

    Windows:
    ```bash
    set HTTPS=true&&npm start
    ```

    macOS:
    ```bash
    HTTPS=true npm start
    ```

   > [!NOTE]
   > Откроется окно браузера с надстройкой. Закройте это окно.

2. В Excel выберите вкладку **Главная** и нажмите кнопку **Показать область задач** на ленте, чтобы открыть область задач надстройки.

    ![Кнопка надстройки Excel](../images/excel-quickstart-addin-2b.png)

3. Выберите любой диапазон ячеек на листе.

4. В области задач нажмите кнопку **Set color** (Задать цвет), чтобы сделать выбранный диапазон зеленым.

    ![Надстройка Excel](../images/excel-quickstart-addin-2c.png)

## <a name="next-steps"></a>Дальнейшие действия

Поздравляем, вы успешно создали надстройку Excel с помощью React! Чтобы узнать больше о возможностях надстроек Excel и создать более сложную надстройку, воспользуйтесь руководством по надстройкам Excel.

> [!div class="nextstepaction"]
> [Руководство по надстройкам Excel](../tutorials/excel-tutorial.yml)

## <a name="see-also"></a>См. также

* [Руководство по надстройкам Excel](../tutorials/excel-tutorial-create-table.md)
* [Основные понятия API JavaScript для Excel](../excel/excel-add-ins-core-concepts.md)
* [Примеры кода надстроек Excel](http://dev.office.com/code-samples#?filters=excel,office%20add-ins)
* [Справочник по API JavaScript для Excel](https://dev.office.com/reference/add-ins/excel/excel-add-ins-reference-overview)
