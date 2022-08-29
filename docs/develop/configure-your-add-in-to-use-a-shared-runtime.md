---
title: Настройка надстройки Office для использования общей среды выполнения
description: Настройте надстройку Office для использования общей среды выполнения для поддержки дополнительных функций ленты, области задач и пользовательских функций.
ms.date: 07/18/2022
ms.prod: non-product-specific
ms.localizationpriority: high
ms.openlocfilehash: e6b10cc2d342d95a8542146ecbd95d750322421f
ms.sourcegitcommit: 0be4cd0680d638cf96c12263a71af59ff9f51f5a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 08/24/2022
ms.locfileid: "67422938"
---
# <a name="configure-your-office-add-in-to-use-a-shared-runtime"></a>Настройка надстройки Office для использования общей среды выполнения

[!include[Shared runtime requirements](../includes/shared-runtime-requirements-note.md)]

Вы можете настроить надстройку Office для выполнения всего кода в одной общей [среде выполнения](../testing/runtimes.md#shared-runtime). Это позволяет повысить слаженность работы всей вашей надстройки и обеспечить доступ к DOM и CORS из всех ее частей. Кроме того, это позволяет использовать дополнительные функции, например запуск кода при открытии документа, а также включение и отключение кнопок ленты. Чтобы настроить надстройку для использования общей среды выполнения, следуйте инструкциям, приведенным в этой статье.

## <a name="create-the-add-in-project"></a>Создание проекта надстройки

Если вы начинаете новый проект, используйте [генератор Yeoman для настроек Office](yeoman-generator-overview.md), чтобы создать проект надстройки Excel, PowerPoint или Word.

Запустите команду `yo office --projectType taskpane --name "my office add in" --host <host> --js true`, где `<host>` имеет одно из следующих значений.

- excel
- powerpoint
- word

> [!IMPORTANT]
> Значение аргумента `--name` должно быть указано в двойных кавычках, даже если оно не содержит пробелов.

Вы можете использовать различные параметры для параметров командной строки **--projecttype**, **--name** и **--js**. Полный список вариантов см. в статье [Генератор Yeoman для надстроек Office](https://github.com/OfficeDev/generator-office).

Генератор создаст проект и установит вспомогательные компоненты Node. Кроме того, с помощью действий из этой статьи вы можете обновить проект Visual Studio, чтобы использовать общую среду выполнения. Однако вам может потребоваться обновить схемы XML для манифеста. Дополнительные сведения см. в статье [Устранение ошибок разработки с надстройками Office](../testing/troubleshoot-development-errors.md#manifest-schema-validation-errors-in-visual-studio-projects).

## <a name="configure-the-manifest"></a>Настройка манифеста

Выполните указанные ниже действия для нового или существующего проекта, чтобы настроить его для использования общей среды выполнения. Эти действия подразумевают, что вы создали проект с помощью [генератора Yeoman для надстроек Office](yeoman-generator-overview.md).

1. Запустите Visual Studio Code и откройте свою надстройку.
1. Откройте файл **manifest.xml**.
1. Для надстройки Excel или PowerPoint обновите раздел с требованиями, включив [общую среду выполнения](/javascript/api/requirement-sets/common/shared-runtime-requirement-sets). Обязательно удалите требование `CustomFunctionsRuntime`, если оно присутствует. XML-код должен выглядеть следующим образом.

    ```xml
    <Hosts>
      <Host Name="Workbook"/>
    </Hosts>
    <Requirements>
      <Sets DefaultMinVersion="1.1">
        <Set Name="SharedRuntime" MinVersion="1.1"/>
      </Sets>
    </Requirements>
    <DefaultSettings>
    ```

    > [!NOTE]
    > Не добавляйте набор требований `SharedRuntime` к манифесту надстройки Word. Это приведет к ошибке при загрузке надстройки, и на данный момент это известная проблема.

1. Найдите раздел **\<VersionOverrides\>** и добавьте следующий раздел **\<Runtimes\>**. Время существования должно иметь значение **long**, чтобы код надстройки мог выполняться даже после закрытия области задач. Значение `resid` — **Taskpane.Url**, указывающее расположение файла **taskpane.html** в разделе `<bt:Urls>` в нижней части **manifest.xml**.

    > [!IMPORTANT]
    > Раздел **\<Runtimes\>** должен быть введен после элемента **\<Host\>** точно в таком же порядке, как показано в следующем XML-коде.

   ```xml
   <VersionOverrides ...>
     <Hosts>
       <Host ...>
         <Runtimes>
           <Runtime resid="Taskpane.Url" lifetime="long" />
         </Runtimes>
       ...
       </Host>
   ```

1. Если вы создали надстройку Excel с пользовательскими функциями, найдите элемент **\<Page\>**. Затем измените расположение источника с **Functions.Page.Url** на **Taskpane.Url**.

   ```xml
   <AllFormFactors>
   ...
   <Page>
     <SourceLocation resid="Taskpane.Url"/>
   </Page>
   ...
   ```

1. Найдите тег **\<FunctionFile\>** и измените `resid` с **Commands.Url** на **Taskpane.Url**. Обратите внимание: если у вас нет команд действий, у вас не будет записи **\<FunctionFile\>**, и этот шаг можно пропустить.

    ```xml
    </GetStarted>
    ...
    <FunctionFile resid="Taskpane.Url"/>
    ...
    ```

1. Сохраните файл **manifest.xml**.

## <a name="configure-the-webpackconfigjs-file"></a>Настройка файла webpack.config.js.

Файл **webpack.config.js** создает несколько загрузчиков среды выполнения. Его необходимо изменить, чтобы загрузить только общую среду выполнения с **помощьюtaskpane.htmlфайла** .

1. Запустите Visual Studio Code и откройте созданный вами проект надстройки.
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
> Если в проекте есть файлы **functions.html** или **commands.html**, их можно удалить. Этот **taskpane.html** загрузит кодfunctions.jsи **** **commands.js** в общую среду выполнения с помощью только что выполненных обновлений webpack.

## <a name="test-your-office-add-in-changes"></a>Тестирование изменений надстройки Office

Чтобы убедиться, что общая среда выполнения используется правильно, выполните следующие инструкции.

1. Откройте файл **taskpane.js**.
1. Замените все содержимое файла указанным ниже кодом. Отобразится количество открытий области задач. Добавление события onVisibilityModeChanged поддерживается только в общей среде выполнения.

    ```javascript
    /*global document, Office*/

    let _count = 0;

    Office.onReady(() => {
      document.getElementById("sideload-msg").style.display = "none";
      document.getElementById("app-body").style.display = "flex";

      updateCount(); // Update count on first open.
      Office.addin.onVisibilityModeChanged(function (args) {
        if (args.visibilityMode === "Taskpane") {
          updateCount(); // Update count on subsequent opens.
        }
      });
    });

    function updateCount() {
      _count++;
      document.getElementById("run").textContent = "Task pane opened " + _count + " times.";
    }
    ```

1. Сохраните изменения и запустите проект.

   ```command line
   npm start
   ```

Каждый раз, когда вы открываете область задач, количество открытий увеличивается на единицу. Значение **_count** не будет потеряно, так как общая среда выполнения продолжает выполнение кода даже при закрытии области задач.

## <a name="runtime-lifetime"></a>Срок существования среды выполнения

При добавлении **\<Runtime\>** элемента также указывается время существования со значением или `long` `short`. Установите значение `long`, чтобы воспользоваться такими функциями, как запуск надстройки при открытии документа, продолжение выполнения кода после закрытия области задач или использование CORS и DOM из пользовательских функций.

> [!NOTE]
> По умолчанию используется значение срока жизни `short`, но мы рекомендуем использовать `long` в надстройках Excel, PowerPoint и Word. Если вы настроите в этом примере для среды выполнения значение `short`, ваша надстройка запустится при нажатии одной из кнопок на ленте, но может завершить работу после окончания функционирования обработчика ленты. Аналогичным образом надстройка запустится при открытии области задач, но может завершить работу после закрытия области задач.

```xml
<Runtimes>
  <Runtime resid="ContosoAddin.Url" lifetime="long" />
</Runtimes>
```

> [!NOTE]
> **\<Runtimes\>** Если надстройка содержит элемент в манифесте (требуется для общей среды выполнения) и выполняются условия использования Microsoft Edge с WebView2 (на основе Chromium), он использует этот элемент управления WebView2. Если эти условия не выполнены, используется Internet Explorer 11 (в версии для Windows или Microsoft 365). Дополнительные сведения см. в статьях "[Элемент Runtimes](/javascript/api/manifest/runtimes)" и "[Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md)".

## <a name="about-the-shared-runtime"></a>Сведения об общей среде выполнения

В Windows или Mac надстройка будет выполнять код для кнопок ленты, пользовательских функций и области задач в отдельных средах выполнения. Из-за этого возникают ограничения, например невозможность удобно предоставлять общий доступ к глобальным данным и отсутствие доступа ко всей функциональности CORS для пользовательской функции.

Однако вы можете настроить надстройку Office для совместного использования кода в той же среде выполнения (также называемой общей средой выполнения). За счет этого повышается скоординированность работы надстройки и упрощается доступ к модели DOM и CORS области задач из всех компонентов надстройки.

При настройке общей среды выполнения становятся возможными следующие сценарии.

- Надстройка Office может использовать дополнительные функции пользовательского интерфейса
  - [Включение и отключение команд надстроек](../design/disable-add-in-commands.md)
  - [Запуск кода в надстройке Office при открытии документа](run-code-on-document-open.md)
  - [Отображение и скрытие области задач надстройки Office](show-hide-add-in.md)
- Следующие функции доступны только для надстроек Excel.
  - [Добавление пользовательских сочетаний клавиш в надстройки Office (предварительная версия)](../design/keyboard-shortcuts.md)
  - [Создание пользовательских контекстных вкладок в надстройках Office (предварительная версия)](../design/contextual-tabs.md)
  - Пользовательские функции полностью поддерживают CORS.
  - Пользовательские функции могут вызывать API Office.js для чтения данных из электронной таблицы.

Для Office в Windows общая среда выполнения использует Microsoft Edge с WebView2 (на основе Chromium), если условия его использования выполнены, как объясняется в статье [Браузеры, используемые надстройками Office](../concepts/browsers-used-by-office-web-add-ins.md). В противном случае используется Internet Explorer 11. Кроме того, все кнопки, которые надстройка отображает на ленте, будут работать в одной и той же общей среде выполнения. На следующем рисунке показано, как пользовательские функции, пользовательский интерфейс ленты и код области задач будут выполняться в одной среде выполнения.

![Схема пользовательской функции, области задач и кнопок ленты, работающих в общей среде выполнения браузера в Excel.](../images/custom-functions-in-browser-runtime.png)

### <a name="debug"></a>Отладка

В настоящее время при использовании общей среды выполнения невозможно использовать Visual Studio Code для отладки пользовательских функций в Excel под управлением Windows. Вместо этого потребуется использовать средства разработчика. Дополнительные сведения см. в статье [Отладка надстроек с помощью средств разработчика для Internet Explorer](../testing/debug-add-ins-using-f12-tools-ie.md) или [Отладка надстроек с помощью средств разработчика в Microsoft Edge (на основе Chromium)](../testing/debug-add-ins-using-devtools-edge-chromium.md).

### <a name="multiple-task-panes"></a>Несколько областей задач

Не планируйте использовать в своей надстройке несколько областей задач, если предполагается использование общей среды выполнения. Общая среда выполнения поддерживает только одну область задач. Обратите внимание: любая область задач без `<TaskpaneID>` считается другой областью задач.

## <a name="see-also"></a>См. также

- [Вызов API Excel из пользовательской функции](../excel/call-excel-apis-from-custom-function.md)
- [Добавление пользовательских сочетаний клавиш в надстройки Office (предварительная версия)](../design/keyboard-shortcuts.md)
- [Создание пользовательских контекстных вкладок в надстройках Office (предварительная версия)](../design/contextual-tabs.md)
- [Включение и отключение команд надстроек](../design/disable-add-in-commands.md)
- [Запуск кода в надстройке Office при открытии документа](run-code-on-document-open.md)
- [Отображение и скрытие области задач надстройки Office](show-hide-add-in.md)
- [Учебное руководство. Обмен данными и событиями между пользовательскими функциями Excel и областью задач](../tutorials/share-data-and-events-between-custom-functions-and-the-task-pane-tutorial.md)
- [Среды выполнения в надстройки Office](../testing/runtimes.md)
