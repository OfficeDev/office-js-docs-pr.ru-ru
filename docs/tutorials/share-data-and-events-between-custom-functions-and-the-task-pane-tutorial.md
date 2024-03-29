---
title: 'Учебное руководство: обмен данными и событиями между пользовательскими функциями Excel и областью задач'
description: Узнайте, как обмениваться данными и событиями между пользовательскими функциями и областью задач в Excel.
ms.date: 06/15/2022
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: b19569ce191f0c7dafc0877984a0f05595380e05
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422749"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a>Учебное руководство: обмен данными и событиями между пользовательскими функциями Excel и областью задач

Общий доступ к глобальным данным и отправка событий между областью задач и пользовательскими функциями надстройки Excel с общей средой выполнения. Рекомендуется использовать общую среду выполнения для большинства пользовательских функций, если у вас нет особой причины применять пользовательскую надстройку только для функций. В этом учебном руководстве предполагается, что вы знакомы с использованием [генератора Yeoman для надстроек Office](../develop/yeoman-generator-overview.md) с целью создания проектов надстроек. Если вы еще этого не сделали, рекомендуется ознакомиться с [руководством по пользовательским функциям в Excel](excel-tutorial-create-custom-functions.md).

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

Используйте [генератор Yeoman для надстроек Office](../develop/yeoman-generator-overview.md), чтобы создать проект надстройки Excel.

- Чтобы создать надстройку Excel с пользовательскими функциями, выполните указанную ниже команду.

    ```command&nbsp;line
    yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true
    ```

Генератор создаст проект и установит вспомогательные компоненты Node.

## <a name="configure-the-manifest"></a>Настройка манифеста

Выполните следующие действия, чтобы настроить проект надстройки для использования общей среды выполнения.

1. Запустите Visual Studio Code и откройте созданный вами проект надстройки.
1. Откройте файл **manifest.xml**.
1. Замените (или добавьте) следующий **\<Requirements\>** XML-код раздела, чтобы требовать набор обязательных [элементов общей среды выполнения](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets).

    ```xml
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    ```

    После обновления ваш XML-манифест должен отображаться в следующем порядке.

    ```xml
    <Hosts>
      <Host Name="..."/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

1. Найдите раздел **\<VersionOverrides\>** и добавьте следующий раздел **\<Runtimes\>**. Время существования должно иметь значение **long**, чтобы код надстройки мог выполняться даже после закрытия области задач. Значение `resid` — **Taskpane.Url**, указывающее расположение файла **taskpane.html** в разделе `<bt:Urls>` в нижней части **manifest.xml**.

    ```xml
    <Runtimes>
      <Runtime resid="Taskpane.Url" lifetime="long" />
    </Runtimes>
    ```

    > [!IMPORTANT]
    > Раздел **\<Runtimes\>** должен быть введен после элемента `<Host xsi:type="...">` точно в таком же порядке, как показано в следующем XML-коде.

    ```xml
    <VersionOverrides ...>
      <Hosts>
        <Host xsi:type="...">
          <Runtimes>
            <Runtime resid="Taskpane.Url" lifetime="long" />
          </Runtimes>
        ...
        </Host>
    ```

    > [!NOTE]
    > **\<Runtimes\>** Если надстройка содержит элемент в манифесте (требуется для общей среды выполнения) и выполняются условия использования Microsoft Edge с WebView2 (на основе Chromium), он использует этот элемент управления WebView2. Если эти условия не выполнены, используется Internet Explorer 11 (в версии для Windows или Microsoft 365). Дополнительные сведения см. в статьях "[Элемент Runtimes](/javascript/api/manifest/runtimes)" и "[Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md)".

1. Найдите элемент **\<Page\>**. Затем измените расположение источника с **Functions.Page.Url** на **Taskpane.Url**.

    ```xml
    <AllFormFactors>
    ...
    <Page>
      <SourceLocation resid="Taskpane.Url"/>
    </Page>
    ...
    ```

1. Найдите тег `<FunctionFile ...>` и измените `resid` с **Commands.Url** на **Taskpane.Url**.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Сохраните файл **manifest.xml**.

## <a name="configure-the-webpackconfigjs-file"></a>Настройка файла webpack.config.js.

Файл **webpack.config.js** создает несколько загрузчиков среды выполнения. Его необходимо изменить, чтобы загрузить только общую среду выполнения с **помощьюtaskpane.htmlфайла** .

1. Откройте файл **webpack.config.js**.
1. Перейдите в раздел `plugins:`.
1. Удалите следующий подключаемый модуль `functions.html`, если он существует.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. Удалите следующий подключаемый модуль `commands.html`, если он существует.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. Если вы удалили подключаемый модуль `functions` или `commands`, добавьте их в качестве `chunks`. В коде JavaScript ниже показана обновленная запись, если вы удалили оба подключаемых модуля: `functions` и `commands`.

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. Сохраните изменения и выполните повторную сборку проекта.

    ```command&nbsp;line
    npm run build
    ```

    > [!NOTE]
    > Вы также можете удалить файлы **functions.html** и **commands.html**. Этот **taskpane.html** загружает код **functions.js** **иcommands.js** в общую среду выполнения с помощью только что выполненных обновлений webpack.

1. Сохраните изменения и запустите проект. Убедитесь, что он загружается и выполняется без ошибок.

   ```command&nbsp;line
   npm run start
   ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a>Общий доступ к состоянию для пользовательской функции и кода области задач

Теперь пользовательские функции выполняются в том же контексте, что и код области задач, и они могут получить общий доступ к состоянию, не используя объект **Storage**. В приведенных ниже инструкциях показано, как предоставить общий доступ к глобальной переменной для пользовательской функции и кода области задач.

### <a name="create-custom-functions-to-get-or-store-shared-state"></a>Создание пользовательских функций для получения или сохранения общего состояния

1. В Visual Studio Code откройте файл **src/functions/functions.js**.
1. В строке 1 в самом верху вставьте следующий код. При этом будет инициализирована глобальная переменная **sharedState**.

    ```js
    window.sharedState = "empty";
    ```

1. Добавьте следующий код, чтобы создать пользовательскую функцию, которая сохранит значения переменной **sharedState**.

    ```js
    /**
     * Saves a string value to shared state with the task pane
     * @customfunction STOREVALUE
     * @param {string} value String to write to shared state with task pane.
     * @return {string} A success value
     */
    function storeValue(sharedValue) {
      window.sharedState = sharedValue;
      return "value stored";
    }
    ```

1. Добавьте следующий код, чтобы создать пользовательскую функцию, которая получит текущее значение переменной **sharedState**.

    ```js
    /**
     * Gets a string value from shared state with the task pane
     * @customfunction GETVALUE
     * @returns {string} String value of the shared state with task pane.
     */
    function getValue() {
      return window.sharedState;
    }
    ```

1. Сохраните файл.

### <a name="create-task-pane-controls-to-work-with-global-data"></a>Создание элементов управления области задач для работы с глобальными данными

1. Откройте файл **src/taskpane/taskpane.html**.
1. Добавьте следующий элемент сценария непосредственно перед закрывающим элементом `</head>`.

    ```HTML
    <script src="../functions/functions.js"></script>
    ```

1. После закрытия элемента `</main>` добавьте следующий HTML-код. С помощью HTML будут созданы два текстовых поля и кнопки для получения и хранения глобальных данных.

    ```HTML
    <ol>
      <li>
        Enter a value to send to the custom function and select
        <strong>Store</strong>.
      </li>
      <li>
        Enter <strong>=CONTOSO.GETVALUE()</strong> into a cell to retrieve it.
      </li>
      <li>
        To send data to the task pane, in a cell, enter
        <strong>=CONTOSO.STOREVALUE("new value")</strong>
      </li>
      <li>Select <strong>Get</strong> to display the value in the task pane.</li>
    </ol>

    <p>Store new value to shared state</p>
    <div>
      <input type="text" id="storeBox" />
      <button onclick="storeSharedValue()">Store</button>
    </div>

    <p>Get shared state value</p>
    <div>
      <input type="text" id="getBox" />
      <button onclick="getSharedValue()">Get</button>
    </div>
    ```

1. Перед закрывающим элементом `</body>` добавьте приведенный ниже сценарий. Этот код обрабатывает события нажатия кнопки, когда пользователь хочет сохранить или получить глобальные данные.

    ```HTML
    <script>
      function storeSharedValue() {
        let sharedValue = document.getElementById('storeBox').value;
        window.sharedState = sharedValue;
      }

      function getSharedValue() {
        document.getElementById('getBox').value = window.sharedState;
      }
   </script>
   ```

1. Сохраните файл.
1. Выполните построение проекта.

   ```command line
   npm run build
   ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a>Обмен данными между пользовательскими функциями и областью задач

- Запустите проект, выполнив приведенную ниже команду.

    ```command line
    npm run start
    ```

После запуска Excel можно использовать кнопки области задач для хранения или получения общих данных. Введите `=CONTOSO.GETVALUE()` в ячейку, чтобы пользовательская функция получила те же общие данные. Можно также использовать `=CONTOSO.STOREVALUE("new value")` для изменения значения общих данных.

> [!NOTE]
> Как показано в этой статье, при настройке проекта пользовательские функции и область задач совместно используют контекст. Вызов некоторых API Office из пользовательских функций невозможен. Дополнительные сведения см. в статье [Вызов API Microsoft Excel из пользовательской функции](../excel/call-excel-apis-from-custom-function.md).

## <a name="see-also"></a>См. также

- [Настройка надстройки Office для использования общей среды выполнения](../develop/configure-your-add-in-to-use-a-shared-runtime.md)
