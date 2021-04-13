---
ms.date: 04/08/2021
title: Настройка надстройки Office для использования общей среды выполнения JavaScript
ms.prod: non-product-specific
description: Настройте надстройку Office для использования общей среды выполнения JavaScript, чтобы применять дополнительные возможности ленты, области задач и пользовательских функций.
localization_priority: Priority
ms.openlocfilehash: d5f0a5b6d9053f23792012f1658d213a7972b970
ms.sourcegitcommit: 54fef33bfc7d18a35b3159310bbd8b1c8312f845
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/09/2021
ms.locfileid: "51652196"
---
# <a name="configure-your-office-add-in-to-use-a-shared-javascript-runtime"></a>Настройка надстройки Office для использования общей среды выполнения JavaScript

[!include[Shared JavaScript runtime requirements](../includes/shared-runtime-requirements-note.md)]

Вы можете настроить надстройку Office, чтобы выполнять весь ее код в единой общей среде выполнения JavaScript (также называемой общей средой выполнения). Это позволяет повысить слаженность работы всей вашей надстройки и обеспечить доступ к DOM и CORS из всех ее частей. Кроме того, это позволяет использовать дополнительные функции, например запуск кода при открытии документа, а также включение и отключение кнопок ленты. Чтобы настроить надстройку для использования общей среды выполнения JavaScript, следуйте инструкциям, приведенным в этой статье.

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

Если вы начинаете новый проект, выполните указанные ниже действия, чтобы с помощью [генератора Yeoman для настроек Office](https://github.com/OfficeDev/generator-office) создать проект надстройки Excel или PowerPoint.

Выполните одно из указанных ниже действий.

- Чтобы создать надстройку Excel с пользовательскими функциями, выполните команду `yo office --projectType excel-functions --name 'Excel shared runtime add-in' --host excel --js true`.

    или

- Чтобы создать надстройку PowerPoint, выполните команду `yo office --projectType taskpane --name 'PowerPoint shared runtime add-in' --host powerpoint --js true`.

Генератор создаст проект и установит вспомогательные компоненты Node.

## <a name="configure-the-manifest"></a>Настройка манифеста

Выполните указанные ниже действия для нового или существующего проекта, чтобы настроить его для использования общей среды выполнения. Эти действия подразумевают, что вы создали проект с помощью [генератора Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).

1. Запустите код Visual Studio и откройте созданный вами проект надстройки Excel или PowerPoint.
1. Откройте файл **manifest.xml**.
1. Если вы создали надстройку для Excel, обновите раздел требований, чтобы использовать [общую среду выполнения](../reference/requirement-sets/shared-runtime-requirement-sets.md), а не среду выполнения пользовательских функций. XML-код должен выглядеть следующим образом.

    ```xml
    <Requirements>
    <Sets DefaultMinVersion="1.1">
      <Set Name="SharedRuntime" MinVersion="1.1"/>
    </Sets>
    </Requirements>
    ```

1. Найдите раздел `<VersionOverrides>` и добавьте следующий раздел `<Runtimes>` внутри тега `<Host ...>`. Время существования должно иметь значение **long**, чтобы код надстройки мог выполняться даже после закрытия области задач. Значение `resid` — **Taskpane.Url**, указывающее расположение файла **taskpane.html** в разделе ` <bt:Urls>` в нижней части **manifest.xml**.

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
       ...
       <Runtimes>
         <Runtime resid="Taskpane.Url" lifetime="long" />
       </Runtimes>
       ...
   ```

1. Если вы создали надстройку Excel с пользовательскими функциями, найдите элемент `<Page>`. Затем измените расположение источника с **Functions.Page.Url** на **Taskpane.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. Найдите тег `<FunctionFile ...>` и измените `resid` с **Commands.Url** на **Taskpane.Url**. Обратите внимание: если у вас нет команд действий, у вас не будет записи **FunctionFile**, и этот шаг можно пропустить.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Сохраните файл **manifest.xml**.

## <a name="configure-the-webpackconfigjs-file"></a>Настройка файла webpack.config.js.

Файл **webpack.config.js** создает несколько загрузчиков среды выполнения. Вам требуется изменить его, чтобы загружать только общую среду выполнения JavaScript с помощью файла **taskpane.html**.

1. Запустите код Visual Studio и откройте созданный вами проект надстройки Excel или PowerPoint.
1. Откройте файл **webpack.config.js**.
1. Если файл **webpack.config.js** содержит следующий код подключаемого модуля **functions.html**, удалите его.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "functions.html",
        template: "./src/functions/functions.html",
        chunks: ["polyfill", "functions"]
      })
    ```

1. Если файл **webpack.config.js** содержит следующий код подключаемого модуля **commands.html**, удалите его.

    ```javascript
    new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"]
      })
    ```

1. Если в проекте используются блоки **functions** или **commands**, добавьте их в список блоков, как показано ниже (следующий код предназначен для проекта, применяющего оба блока).

    ```javascript
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane", "commands", "functions"]
      })
    ```

1. Сохраните изменения и выполните повторную сборку проекта.

   ```command line
   npm run build
   ```

> [!NOTE]
> Если в проекте есть файлы **functions.html** или **commands.html**, их можно удалить. **Taskpane.html** загружает код **functions.js** и **commands.js** в общую среду выполнения JavaScript с помощью созданных вами обновлений webpack.

## <a name="test-your-office-add-in-changes"></a>Тестирование изменений надстройки Office

Вы можете убедиться, что вы используете общую среду выполнения JavaScript надлежащим образом, воспользовавшись следующими инструкциями.

1. Откройте файл **manifest.xml**.
1. Найдите раздел `<Control xsi:type="Button" id="TaskpaneButton">` и измените следующий XML-код `<Action ...>`.

    с:

    ```xml
    <Action xsi:type="ShowTaskpane">
      <TaskpaneId>ButtonId1</TaskpaneId>
      <SourceLocation resid="Taskpane.Url"/>
    </Action>
    ```

    на:

    ```xml
    <Action xsi:type="ExecuteFunction">
      <FunctionName>action</FunctionName>
    </Action>
    ```

1. Откройте файл **./src/commands/commands.js**.
1. Замените имеющуюся функцию **action** указанным ниже кодом. При этом функция будет обновлена для открытия и изменения кнопки области задач, чтобы увеличить счетчик. Открытие модели DOM области задач и доступ к ней из команды поддерживается только в общей среде выполнения JavaScript.

    ```javascript
    var _count=0;
    
    function action(event) {
      // Your code goes here.
      _count++;
      Office.addin.showAsTaskpane();
      document.getElementById("run").textContent="Go"+_count;
    
      // Be sure to indicate when the add-in command function is complete.
      event.completed();
    }
    ```

1. Сохраните изменения и запустите проект.

   ```command line
   npm start
   ```

Каждый раз при нажатии кнопки надстройки текст кнопки **run** (выполнить) будет изменяться на **go** (перейти) с увеличением счетчика после этого.

## <a name="runtime-lifetime"></a>Срок существования среды выполнения

Добавляя элемент `Runtime`, вы также задаете срок существования со значением `long` или `short`. Установите значение `long`, чтобы воспользоваться такими функциями, как запуск надстройки при открытии документа, продолжение выполнения кода после закрытия области задач или использование CORS и DOM из пользовательских функций.

> [!NOTE]
> По умолчанию используется значение срока жизни `short`, но мы рекомендуем использовать `long` в надстройках Excel. Если вы настроите в этом примере для среды выполнения значение `short`, ваша надстройка Excel запустится при нажатии одной из кнопок на ленте, но может завершить работу после окончания функционирования обработчика ленты. Аналогичным образом надстройка запустится при открытии области задач, но может завершить работу после закрытия области задач.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> Если в манифесте надстройки есть элемент `Runtimes` (требуемый для общей среды выполнения), она использует Internet Explorer 11 независимо от того, какая у вас версия Windows или Microsoft 365. Дополнительные сведения см. в статье [Runtimes](../reference/manifest/runtimes.md).

## <a name="about-the-shared-javascript-runtime"></a>Сведения об общей среде выполнения JavaScript

На компьютере с Windows или Mac надстройка запускает код для кнопок ленты, пользовательских функций и области задач в отдельных средах выполнения JavaScript. Из-за этого возникают ограничения, например невозможность удобно предоставлять общий доступ к глобальным данным и отсутствие доступа ко всей функциональности CORS для пользовательской функции.

Однако вы можете настроить надстройку Office так, чтобы обеспечить общий доступ к коду в одной среде выполнения JavaScript (то есть в общей среде выполнения). За счет этого повышается скоординированность работы надстройки и упрощается доступ к модели DOM и CORS области задач из всех компонентов надстройки.

При настройке общей среды выполнения становятся возможными следующие сценарии.

- Надстройка Office может использовать дополнительные функции пользовательского интерфейса.
  - [Добавление пользовательских сочетаний клавиш в надстройки Office (предварительная версия)](../design/keyboard-shortcuts.md)
  - [Создание пользовательских контекстных вкладок в надстройках Office (предварительная версия)](../design/contextual-tabs.md)
  - [Включение и отключение команд надстроек](../design/disable-add-in-commands.md)
  - [Запуск кода в надстройке Office при открытии документа](run-code-on-document-open.md)
  - [Отображение и скрытие области задач надстройки Office](show-hide-add-in.md)
- Для надстроек Excel:
  - Пользовательские функции полностью поддерживают CORS.
  - Пользовательские функции могут вызывать API Office.js для чтения данных из электронной таблицы.

Для Office в Windows общая среда выполнения требует наличия экземпляра браузера Microsoft Internet Explorer 11, как описано в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md). Кроме того, все кнопки, отображаемые вашей надстройкой на ленте, будут работать в этой же общей среде выполнения. На следующем рисунке показано, как пользовательские функции, пользовательский интерфейс ленты и код области задач будут запускаться в одной среде выполнения JavaScript.

![Схема пользовательской функции, области задач и кнопок ленты, работающих в общей среде выполнения браузера IE в Excel](../images/custom-functions-in-browser-runtime.png)

### <a name="debugging"></a>Отладка

В настоящее время при использовании общей среды выполнения невозможно использовать Visual Studio Code для отладки пользовательских функций в Excel под управлением Windows. Вместо этого потребуется использовать средства разработчика. Дополнительные сведения см. в статье [Отладка надстроек с помощью средств разработчика в Windows 10](../testing/debug-add-ins-using-f12-developer-tools-on-windows-10.md).

### <a name="multiple-task-panes"></a>Несколько областей задач

Не планируйте использовать в своей надстройке несколько областей задач, если предполагается использование общей среды выполнения. Общая среда выполнения поддерживает только одну область задач. Обратите внимание: любая область задач без `<TaskpaneID>` считается другой областью задач.

## <a name="give-us-feedback"></a>Напишите нам свой отзыв

Мы будем рады услышать ваши отзывы об этой функции. Если вы обнаружите какие-либо ошибки или проблемы, если у вас есть запросы относительно этой функции, сообщите нам, создав проблему GitHub в [репозитории office-js](https://github.com/OfficeDev/office-js).

## <a name="see-also"></a>См. также

- [Вызов API Excel из пользовательской функции](../excel/call-excel-apis-from-custom-function.md)
- [Добавление пользовательских сочетаний клавиш в надстройки Office (предварительная версия)](../design/keyboard-shortcuts.md)
- [Создание пользовательских контекстных вкладок в надстройках Office (предварительная версия)](../design/contextual-tabs.md)
- [Включение и отключение команд надстроек](../design/disable-add-in-commands.md)
- [Запуск кода в надстройке Office при открытии документа](run-code-on-document-open.md)
- [Отображение и скрытие области задач надстройки Office](show-hide-add-in.md)
- [Учебное руководство. Обмен данными и событиями между пользовательскими функциями Excel и областью задач](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
