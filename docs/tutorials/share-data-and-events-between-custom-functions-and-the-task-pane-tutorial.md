---
title: 'Учебное руководство: обмен данными и событиями между пользовательскими функциями Excel и областью задач'
description: Узнайте, как обмениваться данными и событиями между пользовательскими функциями и областью задач в Excel.
ms.date: 09/23/2021
ms.prod: excel
ms.localizationpriority: high
ms.openlocfilehash: 714f7dc62c7357a67ac26179dee6abc1d229ea49
ms.sourcegitcommit: 517786511749c9910ca53e16eb13d0cee6dbfee6
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 09/29/2021
ms.locfileid: "59990532"
---
# <a name="tutorial-share-data-and-events-between-excel-custom-functions-and-the-task-pane"></a>Учебное руководство: обмен данными и событиями между пользовательскими функциями Excel и областью задач

Вы можете настроить свою надстройку Excel для использования общей среды выполнения. Это позволяет предоставлять общий доступ к глобальным данным или отправлять события между областью задач и пользовательскими функциями.

Для большинства пользовательских функций рекомендуется пользоваться общей средой выполнения, если у вас нет особой причины применять пользовательскую функцию без области задач (без пользовательского интерфейса).

В этом учебном руководстве предполагается, что вы знакомы с использованием генератора Yo Office для создания проектов надстроек. Если вы еще этого не сделали, рекомендуется ознакомиться с [руководством по пользовательским функциям в Excel](excel-tutorial-create-custom-functions.md).

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

Создайте проект надстройки Excel помощью генератора Yeoman. Выполните приведенную ниже команду и ответьте на вопросы, как показано ниже.

```command line
yo office
```

- Выберите тип проекта: **проект надстройки пользовательских функций Excel**
- Выберите тип сценария: **JavaScript**
- Как вы хотите назвать надстройку? **Моя надстройка Office**

![Снимок экрана: запросы и ответы для генератора Yeoman в интерфейсе командной строки.](../images/yo-office-excel-project.png)

После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

## <a name="configure-the-manifest"></a>Настройка манифеста

1. Запустите Visual Studio Code и откройте проект **Моя надстройка Office**.
2. Откройте файл **manifest.xml**.
3. Найдите раздел `<VersionOverrides>` и добавьте следующий раздел `<Runtimes>`. Время существования должно быть **длительным**, чтобы пользовательские функции могли работать даже после закрытия области задач.

   ```xml
   <VersionOverrides xmlns="http://schemas.microsoft.com/office/taskpaneappversionoverrides" xsi:type="VersionOverridesV1_0">
     <Hosts>
       <Host xsi:type="Workbook">
         <Runtimes>
           <Runtime resid="ContosoAddin.Url" lifetime="long" />
         </Runtimes>
       <AllFormFactors>
   ```

> [!NOTE]
> Если в манифесте вашей надстройки есть элемент `Runtimes`, необходимый для общей среды выполнения, и при этом выполнены условия для использования Microsoft Edge с WebView2 (на основе Chromium), то будет использоваться этот элемент управления WebView2. Если же эти условия не выполнены, используется Internet Explorer 11 (в версии для Windows или Microsoft 365). Дополнительные сведения см. в разделах документации "[Элемент Runtimes](../reference/manifest/runtimes.md)" и "[Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md)".

4. В элементе `<Page>` замените расположение источника с **Functions.Page.Url** на **ContosoAddin.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Page>
   ...
   ```

5. В разделе `<DesktopFormFactor>` измените **FunctionFile** с **Commands.Url** на **ContosoAddin.Url**.

   ```xml
   <DesktopFormFactor>
   <GetStarted>
   ...
   </GetStarted>
   <FunctionFile resid="ContosoAddin.Url"/>
   ```

6. В разделе `<Action>` измените расположение источника с **Taskpane.Url** на **ContosoAddin.Url**.

   ```xml
   <Action xsi:type="ShowTaskpane">
   <TaskpaneId>ButtonId1</TaskpaneId>
   <SourceLocation resid="ContosoAddin.Url"/>
   </Action>
   ```

7. Добавьте новый **Url-идентификатор** для **ContosoAddin.Url**, указывающий на **taskpane.html**.

   ```xml
   <bt:Urls>
   <bt:Url id="Functions.Script.Url" DefaultValue="https://localhost:3000/dist/functions.js"/>
   ...
   <bt:Url id="ContosoAddin.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
   ...
   ```

8. Сохраните изменения и перестройте проект.

   ```command line
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

1. Откройте файл **src/taskpane/taskpane.html**.
2. Добавьте следующий элемент сценария непосредственно перед закрывающим элементом `</head>`.

   ```html
   <script src="functions.js"></script>
   ```

3. После закрытия элемента `</main>` добавьте следующий HTML-код. С помощью HTML будут созданы два текстовых поля и кнопки для получения и хранения глобальных данных.

   ```html
   <ol>
     <li>
       Enter a value to send to the custom function and select
       <strong>Store</strong>.
     </li>
     <li>
       Enter <strong>=CONTOSO.GETVALUE()</strong>strong> into a cell to retrieve
       it.
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

4. Перед закрывающим элементом `</body>` добавьте приведенный ниже сценарий. Этот код обрабатывает события нажатия кнопки, когда пользователь хочет сохранить или получить глобальные данные.

   ```js
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

5. Сохраните файл.
6. Построение проекта

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
