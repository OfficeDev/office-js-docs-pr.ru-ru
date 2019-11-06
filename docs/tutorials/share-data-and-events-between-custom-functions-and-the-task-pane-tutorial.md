---
ms.date: 11/04/2019
title: Руководство по обмену данными и событиями между пользовательскими функциями и областью задач в Excel (предварительная версия)
ms.prod: excel
description: Осуществляйте обмен данными и событиями между пользовательскими функциями и областью задач в Excel.
localization_priority: Priority
ms.openlocfilehash: dcd4bced7e1419a57256f4ec54e3ff72c0edf9ef
ms.sourcegitcommit: 42bcf9059327a8d71a7ab223805aea68be9ed6b5
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/04/2019
ms.locfileid: "37962111"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane-preview"></a>Руководство по обмену данными и событиями между пользовательскими функциями и областью задач в Excel (предварительная версия)

Пользовательские функции и область задач в Excel совместно используют глобальные данные и могут вызывать функции друг друга. Следуя инструкциям, приведенным в этой статье, настройте проект таким образом, чтобы пользовательские функции могли работать с областью задач.

> [!NOTE]
> Возможности, описанные в этой статье, в настоящее время доступны в предварительной версии и могут изменяться. Сейчас они не поддерживаются для использования в рабочих средах. Возможности предварительной версии, приведенные в этой статье, доступны только в Excel для Windows. Чтобы ознакомиться с ними, вам нужно [присоединиться к программе предварительной оценки Office](https://insider.office.com/ru-RU/join).  Хороший способ ознакомиться с такими возможностями — использование подписки на Office 365. Если у вас еще нет подписки на Office 365, вы можете оформить ее, присоединившись к [программе для разработчиков Office 365](https://developer.microsoft.com/office/dev-program).

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

Создайте проект надстройки Excel помощью генератора Yeoman. Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.

```command&nbsp;line
yo office
```

- Выберите тип проекта: **проект надстройки пользовательских функций Excel**
- Выберите тип сценария: **JavaScript**
- Как вы хотите назвать надстройку? **Моя надстройка Office**

![Снимок экрана: ответы на вопросы Office о создании проекта надстройки.](../images/yo-office-excel-project.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

## <a name="configure-the-manifest"></a>Настройка манифеста

1. Запустите Visual Studio Code и откройте проект **Моя надстройка Office**.
2. Откройте файл **manifest.xml**.
3. Измените раздел `<Requirements>`, чтобы использовать **CustomFunctionsRuntime** версии **1.2**, как показано в приведенном ниже примере кода.
    
    ```xml
    <Requirements> 
    <Sets DefaultMinVersion="1.1">
    <Set Name="CustomFunctionsRuntime" MinVersion="1.2"/>
    </Sets>
    </Requirements>
    ```
    
4. Под элементом `<Host>` листа добавьте следующий раздел `<Runtimes>`. Время существования должно быть **длительным**, чтобы пользовательские функции могли работать даже после закрытия области задач.
    
    ```xml
    <Hosts>
    <Host xsi:type="Workbook">
    <Runtimes>
    <Runtime resid="TaskPaneAndCustomFunction.Url" lifetime="long" />
    </Runtimes>
    ```
    
5. В элементе `<Page>` измените расположение источника с **Functions.Page.Url** на **TaskPaneAndCustomFunction.Url**.

    ```xml
    <AllFormFactors>
    ...
    <Page>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Page>
    ...
    ```

6. В разделе `<DesktopFormFactor>` измените расположение **FunctionFile** с **Commands.Url** на **TaskPaneAndCustomFunction.Url**.
    
    ```xml
    <DesktopFormFactor>
    <GetStarted>
    ...
    </GetStarted>
    <FunctionFile resid="TaskPaneAndCustomFunction.Url"/>
    ```
    
7. В разделе `<Action>` измените расположение источника с **Taskpane.Url** на **TaskPaneAndCustomFunction.Url**.
    
    ```xml
    <Action xsi:type="ShowTaskpane">
    <TaskpaneId>ButtonId1</TaskpaneId>
    <SourceLocation resid="TaskPaneAndCustomFunction.Url"/>
    </Action>
    ```
    
8. Добавьте новый **Url-идентификатор** для **TaskPaneAndCustomFunction.Url**, указывающий на **taskpane.html**.
     
    ```xml
    <bt:Urls>
    <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
    ...
    <bt:Url id="TaskPaneAndCustomFunction.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
    ...
    ```
    
9. Сохраните изменения и перестройте проект.
    
    ```command&nbsp;line
    npm run build
    ```

## <a name="share-state-between-custom-function-and-task-pane-code"></a>Общий доступ к состоянию для пользовательской функции и кода области задач 

Теперь пользовательские функции выполняются в том же контексте, что и код области задач, и они могут получить общий доступ к состоянию, не используя объект **Storage**. В приведенных ниже инструкциях показано, как предоставить общий доступ к глобальной переменной для пользовательской функции и кода области задач.

### <a name="create-custom-functions-to-get-or-store-shared-state"></a>Создание пользовательских функций для получения или сохранения общего состояния

1. В Visual Studio Code откройте файл **src/functions/functions.js**.
2. В строке 1 в самом верху вставьте следующий код. При этом будет инициализирована глобальная переменная **sharedState**.
    
    ```js
    window.sharedState = "empty";
    ```
    
3. Добавьте следующий код, чтобы создать пользовательскую функцию, которая сохранит значения переменной **sharedState**.
    
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
    
4. Добавьте следующий код, чтобы создать пользовательскую функцию, которая получит текущее значение переменной **sharedState**.

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
    
5. Сохраните файл.

### <a name="create-task-pane-controls-to-work-with-global-data"></a>Создание элементов управления области задач для работы с глобальными данными 

1. Откройте файл**src/taskpane/taskpane.html**.
2. После закрытия элемента `</main>` добавьте следующий HTML-код. С помощью HTML будут созданы два текстовых поля и кнопки для получения и хранения глобальных данных.

    ```html
    <ol>
    <li>Enter a value to send to the custom function and select <strong>Store</strong>.</li>
    <li>Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve it.</li>
    <li>To send data to the task pane, in a cell, enter <strong>=CONTOSO.STOREVALUE("new value")</strong></li>
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
    
3. Перед элементом `<body>` добавьте приведенный ниже сценарий. Этот код обрабатывает события нажатия кнопки, когда пользователь хочет сохранить или получить глобальные данные.
    
    ```js
    <script>
    function storeSharedValue() {
    let sharedValue = document.getElementById('storeBox').value;
    window.sharedState = sharedValue;
    }
    
    function getSharedValue() {
    document.getElementById('getBox').value = window.sharedState;
    }</script>
    ```
    
4. Сохраните файл.
5. Построение проекта
    
    ```command&nbsp;line
    npm run build 
    ```

### <a name="try-sharing-data-between-the-custom-functions-and-task-pane"></a>Обмен данными между пользовательскими функциями и областью задач

- Запустите проект, выполнив приведенную ниже команду.

    ```command&nbsp;line
    npm run start
    ```

После запуска Excel можно использовать кнопки области задач для хранения или получения общих данных. Введите `=CONTOSO.GETVALUE()` в ячейку, чтобы пользовательская функция получила те же общие данные. Можно также использовать `=CONTOSO.STOREVALUE(“new value”)` для изменения значения общих данных.

> [!NOTE]
> Как показано в этой статье, при настройке проекта пользовательские функции и область задач совместно используют контекст. Вызов API Office из пользовательских функций не поддерживается. При работе с документом выполните вызов API Office для [события onCalculated](https://docs.microsoft.com/javascript/api/excel/excel.worksheet?view=excel-js-preview#event-details).

