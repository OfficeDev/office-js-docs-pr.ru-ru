---
title: Руководство. Сборка надстройки Outlook для создания сообщения
description: В этом руководстве вы создадите надстройку Outlook, которая вставляет списки GitHub в тело нового сообщения.
ms.date: 05/01/2022
ms.prod: outlook
ms.localizationpriority: high
ms.openlocfilehash: a143ff743c5dfb692709d0291534fd060352f264
ms.sourcegitcommit: 5773c76912cdb6f0c07a932ccf07fc97939f6aa1
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2022
ms.locfileid: "65244795"
---
# <a name="tutorial-build-a-message-compose-outlook-add-in"></a>Руководство. Сборка надстройки Outlook для создания сообщения

В этом руководстве разъясняется, как выполнить сборку надстройки Outlook, которую можно использовать в режиме создания сообщения для вставки содержимого в его текст.

В этом руководстве описан порядок выполнения перечисленных ниже задач.

> [!div class="checklist"]
>
> - Создание проекта надстройки Outlook
> - Определение кнопок, отображаемых в окне создания сообщения
> - Реализация интерфейса первого запуска, который собирает сведения от пользователя и получает данные из внешней службы
> - Реализация кнопки без пользовательского интерфейса, вызывающей функцию
> - Реализация области задач, вставляющей содержимое в текст сообщения

## <a name="prerequisites"></a>Необходимые условия

[!INCLUDE [Yeoman generator prerequisites](../includes/quickstart-yo-prerequisites.md)]

- [Visual Studio Code (VS Code)](https://code.visualstudio.com/) или используемый вами редактор кода

- Outlook 2016 или более поздней версии для Windows (подключенный к учетной записи Microsoft 365) или Outlook в Интернете

- Учетная запись [GitHub](https://www.github.com)

## <a name="setup"></a>Настройка

Надстройка, создаваемая с помощью этого руководства, считывает элементы [gist](https://gist.github.com) из учетной записи GitHub пользователя и добавляет выбранные элементы gist в текст сообщения. Выполните указанные ниже действия для создания двух новых элементов gist, с помощью которых можно проверить создаваемую надстройку.

1. [Выполните вход в GitHub](https://github.com/login).

1. [Создайте новый элемент gist](https://gist.github.com).

    - В поле **Gist description...** (Описание gist) введите **Hello World Markdown**.

    - В поле **Filename including extension...** (Имя файла с расширением) введите **test.md**.

    - Добавьте в многострочное текстовое поле указанную ниже разметку.

        ```markdown
        # Hello World

        This is content converted from Markdown!

        Here's a JSON sample:

          ```json
          {
            "foo": "bar"
          }
          ```
        ```

    - Нажмите кнопку **Create public gist** (Создать общедоступный элемент gist).

1. [Создайте другой элемент gist](https://gist.github.com).

    - В поле **Gist description...** (Описание gist) введите **Hello World Html**.

    - В поле **Filename including extension...** (Имя файла с расширением) введите **test.html**.

    - Добавьте в многострочное текстовое поле указанную ниже разметку.

        ```HTML
        <html>
          <head>
            <style>
            h1 {
              font-family: Calibri;
            }
            </style>
          </head>
          <body>
            <h1>Hello World!</h1>
            <p>This is a test</p>
          </body>
        </html>
        ```

    - Нажмите кнопку **Create public gist** (Создать общедоступный элемент gist).

## <a name="create-an-outlook-add-in-project"></a>Создание проекта надстройки Outlook

1. [!include[Yeoman generator create project guidance](../includes/yo-office-command-guidance.md)]

    - **Выберите тип проекта** - `Office Add-in Task Pane project`

    - **Выберите тип сценария** - `JavaScript`

    - **Как вы хотите назвать надстройку?** - `Git the gist`

    - **Какое клиентское приложение Office должно поддерживаться?** - `Outlook`

    ![Снимок экрана: запросы и ответы для генератора Yeoman в интерфейсе командной строки.](../images/yeoman-prompts-2.png)

    После завершения работы мастера генератор создаст проект и установит вспомогательные компоненты Node.

    [!include[Yeoman generator next steps](../includes/yo-office-next-steps.md)]

1. Перейдите к корневому каталогу проекта.

    ```command&nbsp;line
    cd "Git the gist"
    ```

1. Эта надстройка будет использовать указанные ниже библиотеки.

    - Библиотека [Showdown](https://github.com/showdownjs/showdown) для преобразования Markdown в HTML.
    - Библиотека [URI.js](https://github.com/medialize/URI.js) для создания относительных URL-адресов.
    - Библиотеки [jquery](https://jquery.com/) для упрощения взаимодействий DOM.

     Чтобы установить эти инструменты для своего проекта, выполните в корневом каталоге проекта указанную ниже команду.

    ```command&nbsp;line
    npm install showdown urijs jquery --save
    ```

1. Откройте проект в VS Code или используемом вами редакторе кода.

    [!INCLUDE [Instructions for opening add-in project in VS Code via command line](../includes/vs-code-open-project-via-command-line.md)]

### <a name="update-the-manifest"></a>Обновление манифеста

Манифест надстройки управляет ее отображением в Outlook. Он определяет, как надстройка отображается в списке, а также задает кнопки на ленте и URL-адреса файлов HTML и JavaScript, используемых надстройкой.

#### <a name="specify-basic-information"></a>Указание основных сведений

Внесите следующие изменения в файл **manifest.xml**, чтобы указать некоторые основные сведения о надстройке.

1. Найдите элемент **ProviderName** и замените значение по умолчанию на название вашей компании.

    ```xml
    <ProviderName>Contoso</ProviderName>
    ```

1. Найдите элемент **Description**, замените значение по умолчанию на описание надстройки и сохраните файл.

    ```xml
    <Description DefaultValue="Allows users to access their GitHub gists."/>
    ```

#### <a name="test-the-generated-add-in"></a>Тестирование созданной надстройки

Прежде чем продолжить, протестируйте базовую надстройку, созданную генератором, чтобы подтвердить правильную настройку проекта.

[!INCLUDE [alert use https](../includes/alert-use-https.md)]

1. Выполните указанную ниже команду в корневом каталоге своего проекта. После выполнения этой команды запустится локальный веб-сервер и будет загружена ваша неопубликованная надстройка.

    ```command&nbsp;line
    npm start
    ```

1. Откройте в Outlook существующее сообщение и нажмите кнопку **Показать область задач**.

1. При появлении запроса с диалоговым окном **Остановка при загрузке веб-представления** выберите **ОК**.

    [!INCLUDE [Cancelling the WebView Stop On Load dialog box](../includes/webview-stop-on-load-cancel-dialog.md)]

    Если настройка выполнена правильно, откроется область задач и отобразится страница приветствия надстройки.

    ![Снимок экрана с кнопкой "Показать область задач" и областью задач Git the gist, добавленной после выполнения примера.](../images/button-and-pane.png)

## <a name="define-buttons"></a>Определение кнопок

После проверки работы базовой надстройки можно ее настроить для добавления функций. По умолчанию манифест определяет только кнопки для окна чтения сообщения. Можно обновить манифест, чтобы удалить кнопки из окна чтения сообщения и определить две новые кнопки для окна создания сообщения:

- **Insert gist** (Вставить gist): кнопка, открывающая область задач

- **Insert default gist** (Вставить gist по умолчанию): кнопка, вызывающая функцию

### <a name="remove-the-messagereadcommandsurface-extension-point"></a>Удаление точки расширения MessageReadCommandSurface

Откройте файл **manifest.xml** и найдите элемент **ExtensionPoint** типа **MessageReadCommandSurface**. Удалите этот элемент **ExtensionPoint** (включая его закрывающий тег), чтобы удалить кнопки из окна чтения сообщения.

### <a name="add-the-messagecomposecommandsurface-extension-point"></a>Добавление точки расширения MessageComposeCommandSurface

Найдите в манифесте строку `</DesktopFormFactor>`. Непосредственно перед ней вставьте приведенную ниже разметку XML. Обратите внимание на следующие особенности этой разметки.

- **ExtensionPoint** с `xsi:type="MessageComposeCommandSurface"` означает, что вы определяете кнопки для добавления в окно создания сообщений.

- С помощью элемента **OfficeTab** с параметром `id="TabDefault"` вы указываете, что нужно добавить кнопки на вкладку ленты по умолчанию.

- Элемент **Group** определяет группу новых кнопок, а ресурс **groupLabel** задает подпись группы.

- Первый элемент **Control** содержит элемент **Action** с параметром `xsi:type="ShowTaskPane"`, поэтому эта кнопка открывает область задач.

- Второй элемент **Control** содержит элемент **Action** с параметром `xsi:type="ExecuteFunction"`, поэтому кнопка вызывает функцию JavaScript, содержащуюся в файле функций.

```xml
<!-- Message Compose -->
<ExtensionPoint xsi:type="MessageComposeCommandSurface">
  <OfficeTab id="TabDefault">
    <Group id="msgComposeCmdGroup">
      <Label resid="GroupLabel"/>
      <Control xsi:type="Button" id="msgComposeInsertGist">
        <Label resid="TaskpaneButton.Label"/>
        <Supertip>
          <Title resid="TaskpaneButton.Title"/>
          <Description resid="TaskpaneButton.Tooltip"/>
        </Supertip>
        <Icon>
          <bt:Image size="16" resid="Icon.16x16"/>
          <bt:Image size="32" resid="Icon.32x32"/>
          <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ShowTaskpane">
          <SourceLocation resid="Taskpane.Url"/>
        </Action>
      </Control>
      <Control xsi:type="Button" id="msgComposeInsertDefaultGist">
        <Label resid="FunctionButton.Label"/>
        <Supertip>
          <Title resid="FunctionButton.Title"/>
          <Description resid="FunctionButton.Tooltip"/>
        </Supertip>
        <Icon>
          <bt:Image size="16" resid="Icon.16x16"/>
          <bt:Image size="32" resid="Icon.32x32"/>
          <bt:Image size="80" resid="Icon.80x80"/>
        </Icon>
        <Action xsi:type="ExecuteFunction">
          <FunctionName>insertDefaultGist</FunctionName>
        </Action>
      </Control>
    </Group>
  </OfficeTab>
</ExtensionPoint>
```

### <a name="update-resources-in-the-manifest"></a>Обновление ресурсов в манифесте

Приведенный выше код ссылается на подписи, подсказки и URL-адреса, которые необходимо определить, чтобы манифест был действительным. Вам нужно указать эту информацию в разделе **Resources** манифеста.

1. Найдите элемент **Resources** в файле манифеста и удалите весь элемент (включая его закрывающий тег).

1. Добавьте в то же местоположение следующую разметку, чтобы заменить только что удаленный элемент **Resources**.

    ```xml
    <Resources>
      <bt:Images>
        <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
        <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
        <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
      </bt:Images>
      <bt:Urls>
        <bt:Url id="Commands.Url" DefaultValue="https://localhost:3000/commands.html"/>
        <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
      </bt:Urls>
      <bt:ShortStrings>
        <bt:String id="GroupLabel" DefaultValue="Git the gist"/>
        <bt:String id="TaskpaneButton.Label" DefaultValue="Insert gist"/>
        <bt:String id="TaskpaneButton.Title" DefaultValue="Insert gist"/>
        <bt:String id="FunctionButton.Label" DefaultValue="Insert default gist"/>
        <bt:String id="FunctionButton.Title" DefaultValue="Insert default gist"/>
      </bt:ShortStrings>
      <bt:LongStrings>
        <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Displays a list of your gists and allows you to insert their contents into the current message."/>
        <bt:String id="FunctionButton.Tooltip" DefaultValue="Inserts the content of the gist you mark as default into the current message."/>
      </bt:LongStrings>
    </Resources>
    ```

1. Сохраните изменения манифеста.

### <a name="reinstall-the-add-in"></a>Переустановка надстройки

Чтобы изменения манифеста вступили в силу, переустановите надстройку.

1. Если веб-сервер работает, закройте окно команды узла.

1. Выполните указанные ниже команды, чтобы запустить локальный веб-сервер и автоматически загрузить неопубликованную надстройку.

    ```command&nbsp;line
    npm start
    ```

После повторной установки надстройки можно убедиться, что она установлена успешно, проверив команды **Insert gist** и **Insert default gist** в окне составления сообщений. Обратите внимание, что при выборе этих двух элементов ничего не происходит, так как вы еще не закончили создание этой надстройки.

- При запуске этой надстройки в Outlook 2016 или более поздней версии для Windows отобразятся две новые кнопки на ленте окна составления сообщений: **Insert gist** и **Insert default gist**.

    ![Снимок экрана: меню переполнения элементов ленты в Outlook для Windows с выделенными кнопками надстройки.](../images/add-in-buttons-in-windows.png)

- При запуске этой надстройки в Outlook в Интернете отобразится новая кнопка внизу окна составления сообщений. Нажмите эту кнопку, чтобы просмотреть варианты **Insert gist** (Вставить gist) и **Insert default gist** (Вставить gist по умолчанию).

    ![Снимок экрана: форма создания сообщения в Outlook в Интернете с выделенной кнопкой надстройки и контекстным меню.](../images/add-in-buttons-in-owa.png)

## <a name="implement-a-first-run-experience"></a>Реализация интерфейса первого запуска

Эта надстройка должна иметь возможность считывать элементы gist из учетной записи GitHub пользователя и определять, какой из них пользователь выбрал в качестве используемого по умолчанию. Для выполнения этих целей надстройка должна предложить пользователю указать его имя пользователя GitHub и выбрать элемент gist в качестве используемого по умолчанию из его коллекции существующих элементов gist. Выполните действия, описанные в этом разделе, чтобы реализовать интерфейс первого запуска, отображающий диалоговое окно для получения этих сведений от пользователя.

### <a name="collect-data-from-the-user"></a>Получение данных от пользователя

Начнем с создания пользовательского интерфейса для самого диалогового окна. Создайте в папке **./src** новую подпапку с именем **settings**. Создайте в папке **./src/settings** файл с именем **dialog.html** и добавьте следующую разметку, чтобы определить базовую форму с вводом текста для имени пользователя GitHub, а также пустой список элементов gist, который будет заполнен с помощью JavaScript.

```html
<!DOCTYPE html>
<html>

<head>
  <meta charset="UTF-8" />
  <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
  <title>Settings</title>

  <!-- Office JavaScript API -->
  <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

<!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui. -->
  <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

  <!-- Template styles -->
  <link href="dialog.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-l">
  <main>
    <section class="ms-font-m ms-fontColor-neutralPrimary">
      <div class="not-configured-warning ms-MessageBar ms-MessageBar--warning">
        <div class="ms-MessageBar-content">
          <div class="ms-MessageBar-icon">
            <i class="ms-Icon ms-Icon--Info"></i>
          </div>
          <div class="ms-MessageBar-text">
            Oops! It looks like you haven't configured <strong>Git the gist</strong> yet.
            <br/>
            Please configure your GitHub username and select a default gist, then try that action again!
          </div>
        </div>
      </div>
      <div class="ms-font-xxl">Settings</div>
      <div class="ms-Grid">
        <div class="ms-Grid-row">
          <div class="ms-TextField">
            <label class="ms-Label">GitHub Username</label>
            <input class="ms-TextField-field" id="github-user" type="text" value="" placeholder="Please enter your GitHub username">
          </div>
        </div>
        <div class="error-display ms-Grid-row">
          <div class="ms-font-l ms-fontWeight-semibold">An error occurred:</div>
          <pre><code id="error-text"></code></pre>
        </div>
        <div class="gist-list-container ms-Grid-row">
          <div class="list-title ms-font-xl ms-fontWeight-regular">Choose Default Gist</div>
          <form>
            <div id="gist-list">
            </div>
          </form>
        </div>
      </div>
      <div class="ms-Dialog-actions">
        <div class="ms-Dialog-actionsRight">
          <button class="ms-Dialog-action ms-Button ms-Button--primary" id="settings-done" disabled>
            <span class="ms-Button-label">Done</span>
          </button>
        </div>
      </div>
    </section>
  </main>
  <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
  <script type="text/javascript" src="../helpers/gist-api.js"></script>
  <script type="text/javascript" src="dialog.js"></script>
</body>

</html>
```

Вы могли заметить, что HTML-файл ссылается на файл JavaScript **gist-api.js**, который еще не существует. Этот файл будет создан в разделе [Получение данных из GitHub](#fetch-data-from-github) ниже.

Затем создайте в папке **./src/settings** файл с именем **dialog.css** и добавьте приведенный ниже код, чтобы указать стили, используемые файлом **dialog.html**.

```CSS
section {
  margin: 10px 20px;
}

.not-configured-warning {
  display: none;
}

.error-display {
  display: none;
}

.gist-list-container {
  margin: 10px -8px;
  display: none;
}

.list-title {
  border-bottom: 1px solid #a6a6a6;
  padding-bottom: 5px;
}

ul {
  margin-top: 10px;
}

.ms-ListItem-secondaryText,
.ms-ListItem-tertiaryText {
  padding-left: 15px;
}
```

Теперь, после определения пользовательского интерфейса диалогового окна, можно написать код для выполнения в нем действий. Создайте в папке **./src/settings** файл с именем **dialog.js** и добавьте приведенный ниже код. Обратите внимание, что в этом коде используется jQuery для регистрации событий, а также функция **messageParent** для возвращения выбранных пользователем параметров вызывающей стороне.

```js
(function(){
  'use strict';

  // The Office initialize function must be run each time a new page is loaded.
  Office.initialize = function(reason){
    jQuery(document).ready(function(){
      if (window.location.search) {
        // Check if warning should be displayed.
        var warn = getParameterByName('warn');
        if (warn) {
          $('.not-configured-warning').show();
        } else {
          // See if the config values were passed.
          // If so, pre-populate the values.
          var user = getParameterByName('gitHubUserName');
          var gistId = getParameterByName('defaultGistId');

          $('#github-user').val(user);
          loadGists(user, function(success){
            if (success) {
              $('.ms-ListItem').removeClass('is-selected');
              $('input').filter(function() {
                return this.value === gistId;
              }).addClass('is-selected').attr('checked', 'checked');
              $('#settings-done').removeAttr('disabled');
            }
          });
        }
      }

      // When the GitHub username changes,
      // try to load gists.
      $('#github-user').on('change', function(){
        $('#gist-list').empty();
        var ghUser = $('#github-user').val();
        if (ghUser.length > 0) {
          loadGists(ghUser);
        }
      });

      // When the Done button is selected, send the
      // values back to the caller as a serialized
      // object.
      $('#settings-done').on('click', function() {
        var settings = {};

        settings.gitHubUserName = $('#github-user').val();

        var selectedGist = $('.ms-ListItem.is-selected');
        if (selectedGist) {
          settings.defaultGistId = selectedGist.val();

          sendMessage(JSON.stringify(settings));
        }
      });
    });
  };

  // Load gists for the user using the GitHub API
  // and build the list.
  function loadGists(user, callback) {
    getUserGists(user, function(gists, error){
      if (error) {
        $('.gist-list-container').hide();
        $('#error-text').text(JSON.stringify(error, null, 2));
        $('.error-display').show();
        if (callback) callback(false);
      } else {
        $('.error-display').hide();
        buildGistList($('#gist-list'), gists, onGistSelected);
        $('.gist-list-container').show();
        if (callback) callback(true);
      }
    });
  }

  function onGistSelected() {
    $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
    $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
    $('.not-configured-warning').hide();
    $('#settings-done').removeAttr('disabled');
  }

  function sendMessage(message) {
    Office.context.ui.messageParent(message);
  }

  function getParameterByName(name, url) {
    if (!url) {
      url = window.location.href;
    }
    name = name.replace(/[\[\]]/g, "\\$&");
    var regex = new RegExp("[?&]" + name + "(=([^&#]*)|&|#|$)"),
      results = regex.exec(url);
    if (!results) return null;
    if (!results[2]) return '';
    return decodeURIComponent(results[2].replace(/\+/g, " "));
  }
})();
```

#### <a name="update-webpack-config-settings"></a>Обновление настроек конфигурации webpack

Наконец, откройте файл **webpack.config.js**, расположенный в корневом каталоге проекта, и выполните описанные ниже шаги.

1. Найдите объект `entry` в объекте `config` и добавьте новую запись для `dialog`.

    ```js
    dialog: "./src/settings/dialog.js",
    ```

    После этого новый объект `entry` будет выглядеть следующим образом:

    ```js
    entry: {
      polyfill: ["core-js/stable", "regenerator-runtime/runtime"],
      taskpane: "./src/taskpane/taskpane.js",
      commands: "./src/commands/commands.js",
      dialog: "./src/settings/dialog.js",
    },
    ```

1. Найдите массив `plugins` в объекте `config`. В массив `patterns` объекта `new CopyWebpackPlugin` добавьте новые записи для **taskpane.css** и **dialog.css**.

    ```js
    {
      from: "./src/taskpane/taskpane.css",
      to: "taskpane.css",
    },
    {
      from: "./src/settings/dialog.css",
      to: "dialog.css",
    },
    ```

    После этого объект `new CopyWebpackPlugin` будет выглядеть следующим образом:

    ```js
    new CopyWebpackPlugin({
      patterns: [
      {
        from: "./src/taskpane/taskpane.css",
        to: "taskpane.css",
      },
      {
        from: "./src/settings/dialog.css",
        to: "dialog.css",
      },
      {
        from: "assets/*",
        to: "assets/[name][ext][query]",
      },
      {
        from: "manifest*.xml",
        to: "[name]." + buildType + "[ext]",
        transform(content) {
          if (dev) {
            return content;
          } else {
            return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
          }
        },
      },
    ]}),
    ```

1. В том же массиве `plugins` в объекте `config` добавьте этот новый объект в конец массива.

    ```js
    new HtmlWebpackPlugin({
      filename: "dialog.html",
      template: "./src/settings/dialog.html",
      chunks: ["polyfill", "dialog"]
    })
    ```

    После этого новый массив `plugins` будет выглядеть следующим образом:

    ```js
    plugins: [
      new HtmlWebpackPlugin({
        filename: "taskpane.html",
        template: "./src/taskpane/taskpane.html",
        chunks: ["polyfill", "taskpane"],
      }),
      new CopyWebpackPlugin({
        patterns: [
          {
            from: "./src/taskpane/taskpane.css",
            to: "taskpane.css",
          },
          {
            from: "./src/settings/dialog.css",
            to: "dialog.css",
          },
          {
            from: "assets/*",
            to: "assets/[name][ext][query]",
          },
          {
            from: "manifest*.xml",
            to: "[name]." + buildType + "[ext]",
            transform(content) {
              if (dev) {
                return content;
              } else {
                return content.toString().replace(new RegExp(urlDev, "g"), urlProd);
              }
            },
          },
        ],
      }),
      new HtmlWebpackPlugin({
        filename: "commands.html",
        template: "./src/commands/commands.html",
        chunks: ["polyfill", "commands"],
      }),
      new HtmlWebpackPlugin({
        filename: "dialog.html",
        template: "./src/settings/dialog.html",
        chunks: ["polyfill", "dialog"]
      })
    ],
    ```

### <a name="fetch-data-from-github"></a>Получение данных из GitHub

Только что созданный файл **dialog.js** определяет, что надстройка должна загружать элементы gist, если возникает событие **change** для поля имени пользователя GitHub. Для получения элементов gist пользователя из GitHub используется [API элементов gist GitHub](https://developer.github.com/v3/gists/).

Создайте в папке **./src** новую подпапку с именем **helpers**. Создайте в папке **./src/helpers** файл с именем **gist-api.js** и добавьте следующий код, чтобы получить элементы gist пользователя из GitHub и составить список элементов gist.

```js
function getUserGists(user, callback) {
  var requestUrl = 'https://api.github.com/users/' + user + '/gists';

  $.ajax({
    url: requestUrl,
    dataType: 'json'
  }).done(function(gists){
    callback(gists);
  }).fail(function(error){
    callback(null, error);
  });
}

function buildGistList(parent, gists, clickFunc) {
  gists.forEach(function(gist) {

    var listItem = $('<div/>')
      .appendTo(parent);

    var radioItem = $('<input>')
      .addClass('ms-ListItem')
      .addClass('is-selectable')
      .attr('type', 'radio')
      .attr('name', 'gists')
      .attr('tabindex', 0)
      .val(gist.id)
      .appendTo(listItem);

    var desc = $('<span/>')
      .addClass('ms-ListItem-primaryText')
      .text(gist.description)
      .appendTo(listItem);

    var desc = $('<span/>')
      .addClass('ms-ListItem-secondaryText')
      .text(' - ' + buildFileList(gist.files))
      .appendTo(listItem);

    var updated = new Date(gist.updated_at);

    var desc = $('<span/>')
      .addClass('ms-ListItem-tertiaryText')
      .text(' - Last updated ' + updated.toLocaleString())
      .appendTo(listItem);

    listItem.on('click', clickFunc);
  });  
}

function buildFileList(files) {

  var fileList = '';

  for (var file in files) {
    if (files.hasOwnProperty(file)) {
      if (fileList.length > 0) {
        fileList = fileList + ', ';
      }

      fileList = fileList + files[file].filename + ' (' + files[file].language + ')';
    }
  }

  return fileList;
}
```

Выполните указанную ниже команду, чтобы повторно собрать проект.

```command&nbsp;line
npm run build
```

## <a name="implement-a-ui-less-button"></a>Реализация кнопки без пользовательского интерфейса

Эта кнопка надстройки **Insert default gist** (Вставить gist по умолчанию) является кнопкой без пользовательского интерфейса, вызывающей функцию JavaScript вместо открытия области задач, выполняемого многими кнопками надстройки. Если пользователь нажимает кнопку **Insert gist** (Вставить gist), соответствующая функция JavaScript проверяет наличие конфигурации надстройки.

- Если конфигурация надстройки уже выполнена, функция загружает содержимое элемента gist, выбранного пользователем в качестве используемого по умолчанию, и вставляет его в текст сообщения.

- Если конфигурация надстройки еще не выполнена, диалоговое окно параметров предложит пользователю предоставить нужные сведения.

### <a name="update-the-function-file-html"></a>Обновление файла функции (HTML)

Функция, вызываемая кнопкой без пользовательского интерфейса, должна быть определена в файле, указанном в элементе **FunctionFile** манифеста для соответствующего форм-фактора. Этот манифест надстройки указывает `https://localhost:3000/commands.html` в качестве файла функции.

Откройте файл **./src/commands/commands.html** и замените все содержимое приведенной ниже разметкой.

```html
<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

    <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
    <script type="text/javascript" src="../../node_modules/showdown/dist/showdown.min.js"></script>
    <script type="text/javascript" src="../../node_modules/urijs/src/URI.min.js"></script>
    <script type="text/javascript" src="../helpers/addin-config.js"></script>
    <script type="text/javascript" src="../helpers/gist-api.js"></script>
</head>

<body>
  <!-- NOTE: The body is empty on purpose. Since functions in commands.js are
       invoked via a button, there is no UI to render. -->
</body>

</html>
```

Вы могли заметить, что HTML-файл ссылается на файл JavaScript **addin-config.js**, который еще не существует. Этот файл будет создан в этом руководстве в разделе [Создание файла для управления параметрами конфигурации](#create-a-file-to-manage-configuration-settings) ниже.

### <a name="update-the-function-file-javascript"></a>Обновление файла функции (JavaScript)

Откройте файл **./src/commands/commands.js** и замените все содержимое приведенным ниже кодом. Обратите внимание, если функция **insertDefaultGist** определяет, что конфигурация надстройки не выполнена, добавляется параметр `?warn=1` к URL-адресу диалогового окна. Благодаря этому в диалоговом окне параметров отображается панель сообщений, определенная в файле **./src/settings/dialog.html**, которая сообщает пользователю причину появления диалогового окна.

```js
var config;
var btnEvent;

// The initialize function must be run each time a new page is loaded.
Office.initialize = function (reason) {
};

// Add any UI-less function here.
function showError(error) {
  Office.context.mailbox.item.notificationMessages.replaceAsync('github-error', {
    type: 'errorMessage',
    message: error
  }, function(result){
  });
}

var settingsDialog;

function insertDefaultGist(event) {

  config = getConfig();

  // Check if the add-in has been configured.
  if (config && config.defaultGistId) {
    // Get the default gist content and insert.
    try {
      getGist(config.defaultGistId, function(gist, error) {
        if (gist) {
          buildBodyContent(gist, function (content, error) {
            if (content) {
              Office.context.mailbox.item.body.setSelectedDataAsync(content,
                {coercionType: Office.CoercionType.Html}, function(result) {
                  event.completed();
              });
            } else {
              showError(error);
              event.completed();
            }
          });
        } else {
          showError(error);
          event.completed();
        }
      });
    } catch (err) {
      showError(err);
      event.completed();
    }

  } else {
    // Save the event object so we can finish up later.
    btnEvent = event;
    // Not configured yet, display settings dialog with
    // warn=1 to display warning.
    var url = new URI('dialog.html?warn=1').absoluteTo(window.location).toString();
    var dialogOptions = { width: 20, height: 40, displayInIframe: true };

    Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
      settingsDialog = result.value;
      settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
      settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
    });
  }
}

// Register the function.
Office.actions.associate("insertDefaultGist", insertDefaultGist);

function receiveMessage(message) {
  config = JSON.parse(message.message);
  setConfig(config, function(result) {
    settingsDialog.close();
    settingsDialog = null;
    btnEvent.completed();
    btnEvent = null;
  });
}

function dialogClosed(message) {
  settingsDialog = null;
  btnEvent.completed();
  btnEvent = null;
}
```

### <a name="create-a-file-to-manage-configuration-settings"></a>Создание файла для управления параметрами конфигурации

HTML-файл функции ссылается на файл под названием **addin-config.js**, которого еще не существует. В папке **./src/helpers** создайте файл с именем **addin-config.js** и добавьте указанный ниже код. В этом коде используется [объект RoamingSettings](/javascript/api/outlook/office.roamingsettings), позволяющий получать и задавать значения конфигурации.

```js
function getConfig() {
  var config = {};

  config.gitHubUserName = Office.context.roamingSettings.get('gitHubUserName');
  config.defaultGistId = Office.context.roamingSettings.get('defaultGistId');

  return config;
}

function setConfig(config, callback) {
  Office.context.roamingSettings.set('gitHubUserName', config.gitHubUserName);
  Office.context.roamingSettings.set('defaultGistId', config.defaultGistId);

  Office.context.roamingSettings.saveAsync(callback);
}
```

### <a name="create-new-functions-to-process-gists"></a>Создание новых функций для обработки элементов gist

Затем откройте файл **./src/helpers/gist-api.js** и добавьте указанные ниже функции. Обратите внимание на перечисленные ниже аспекты.

- Если элемент gist содержит код HTML, надстройка вставит HTML-код в текст сообщения без изменений.

- Если элемент gist содержит код Markdown, надстройка воспользуется библиотекой [Showdown](https://github.com/showdownjs/showdown), чтобы преобразовать формат Markdown в HTML, и вставит получившийся HTML-код в текст сообщения.

- Если элемент gist содержит любой код, отличный от HTML или Markdown, надстройка вставит его в текст сообщения как фрагмент кода.

```js
function getGist(gistId, callback) {
  var requestUrl = 'https://api.github.com/gists/' + gistId;

  $.ajax({
    url: requestUrl,
    dataType: 'json'
  }).done(function(gist){
    callback(gist);
  }).fail(function(error){
    callback(null, error);
  });
}

function buildBodyContent(gist, callback) {
  // Find the first non-truncated file in the gist
  // and use it.
  for (var filename in gist.files) {
    if (gist.files.hasOwnProperty(filename)) {
      var file = gist.files[filename];
      if (!file.truncated) {
        // We have a winner.
        switch (file.language) {
          case 'HTML':
            // Insert as is.
            callback(file.content);
            break;
          case 'Markdown':
            // Convert Markdown to HTML.
            var converter = new showdown.Converter();
            var html = converter.makeHtml(file.content);
            callback(html);
            break;
          default:
            // Insert contents as a <code> block.
            var codeBlock = '<pre><code>';
            codeBlock = codeBlock + file.content;
            codeBlock = codeBlock + '</code></pre>';
            callback(codeBlock);
        }
        return;
      }
    }
  }
  callback(null, 'No suitable file found in the gist');
}
```

### <a name="test-the-insert-default-gist-button"></a>Тестирование кнопки Insert default gist (Вставить gist по умолчанию)

Сохраните все изменения и выполните в командной строке команду `npm start`, если сервер еще не запущен. Затем выполните указанные ниже действия, чтобы протестировать кнопку **Insert default gist** (Вставить gist по умолчанию).

1. Откройте Outlook и создайте новое сообщение.

1. В окне создания сообщения нажмите кнопку **Insert default gist** (Вставить gist по умолчанию). Вы увидите диалоговое окно, в котором можно настроить надстройку, указав имя пользователя GitHub в диалоговом окне с соответствующим приглашением.

    ![Снимок экрана: диалоговое окно с приглашением настроить надстройку.](../images/addin-prompt-configure.png)

1. В диалоговом окне параметров введите имя пользователя GitHub, а затем нажмите клавишу **TAB** или щелкните в другом месте диалогового окна, чтобы вызвать событие **change**, которое должно загрузить ваш список общедоступных элементов gist. Выберите элемент gist, который будет использоваться по умолчанию, и нажмите **Готово**.

    ![Снимок экрана с диалоговым окном параметров надстройки.](../images/addin-settings.png)

1. Нажмите кнопку **Insert default gist** (Вставить gist по умолчанию) снова. На этот раз содержимое элемента gist должно быть вставлено в текст сообщения.

   > [!NOTE]
   > Outlook для Windows: чтобы применить последние параметры, может потребоваться закрытие и повторное открытие окна создания сообщения.

## <a name="implement-a-task-pane"></a>Реализация области задач

Эта кнопка **Insert gist** (Вставить gist) надстройки открывает область задач и отображает элементы gist пользователя. После этого пользователь сможет выбрать один из элементов gist для вставки в текст сообщения. Если пользователь еще не выполнил конфигурацию надстройки, ему будет предложено сделать это.

### <a name="specify-the-html-for-the-task-pane"></a>Указание HTML для области задач

В созданном вами проекте HTML области задач указан в файле **./src/taskpane/taskpane.html**. Откройте этот файл и замените все содержимое приведенной ниже разметкой.

```html
<!DOCTYPE html>
<html>

<head>
    <meta charset="UTF-8" />
    <meta http-equiv="X-UA-Compatible" content="IE=Edge" />
    <meta name="viewport" content="width=device-width, initial-scale=1">
    <title>Contoso Task Pane Add-in</title>

    <!-- Office JavaScript API -->
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/1.1/hosted/office.js"></script>

   <!-- For more information on Fluent UI, visit https://developer.microsoft.com/fluentui. -->
    <link rel="stylesheet" href="https://static2.sharepointonline.com/files/fabric/office-ui-fabric-core/9.6.1/css/fabric.min.css"/>

    <!-- Template styles -->
    <link href="taskpane.css" rel="stylesheet" type="text/css" />
</head>

<body class="ms-font-l ms-landing-page">
  <main class="ms-landing-page__main">
    <section class="ms-landing-page__content ms-font-m ms-fontColor-neutralPrimary">
      <div id="not-configured" style="display: none;">
        <div class="centered ms-font-xxl ms-u-textAlignCenter">Welcome!</div>
        <div class="ms-font-xl" id="settings-prompt">Please choose the <strong>Settings</strong> icon at the bottom of this window to configure this add-in.</div>
      </div>
      <div id="gist-list-container" style="display: none;">
        <form>
          <div id="gist-list">
          </div>
        </form>
      </div>
      <div id="error-display" style="display: none;" class="ms-u-borderBase ms-fontColor-error ms-font-m ms-bgColor-error ms-borderColor-error">
      </div>
    </section>
    <button class="ms-Button ms-Button--primary" id="insert-button" tabindex=0 disabled>
      <span class="ms-Button-label">Insert</span>
    </button>
  </main>
  <footer class="ms-landing-page__footer ms-bgColor-themePrimary">
    <div class="ms-landing-page__footer--left">
      <img src="../../assets/logo-filled.png" />
      <h1 class="ms-font-xl ms-fontWeight-semilight ms-fontColor-white">Git the gist</h1>
    </div>
    <div id="settings-icon" class="ms-landing-page__footer--right" aria-label="Settings" tabindex=0>
      <i class="ms-Icon enlarge ms-Icon--Settings ms-fontColor-white"></i>
    </div>
  </footer>
  <script type="text/javascript" src="../../node_modules/jquery/dist/jquery.js"></script>
  <script type="text/javascript" src="../../node_modules/showdown/dist/showdown.min.js"></script>
  <script type="text/javascript" src="../../node_modules/urijs/src/URI.min.js"></script>
  <script type="text/javascript" src="../helpers/addin-config.js"></script>
  <script type="text/javascript" src="../helpers/gist-api.js"></script>
  <script type="text/javascript" src="taskpane.js"></script>
</body>

</html>
```

### <a name="specify-the-css-for-the-task-pane"></a>Указание CSS для области задач

В созданном вами проекте CSS области задач указан в файле **./src/taskpane/taskpane.css**. Откройте этот файл и замените все содержимое приведенным ниже кодом.

```css
/* Copyright (c) Microsoft. All rights reserved. Licensed under the MIT license. See full license in root of repo. */
html, body {
  width: 100%;
  height: 100%;
  margin: 0;
  padding: 0;
  overflow: auto; }

body {
  position: relative;
  font-size: 16px; }

main {
  height: 100%;
  overflow-y: auto; }

footer {
  width: 100%;
  position: relative;
  bottom: 0;
  margin-top: 10px;}

p, h1, h2, h3, h4, h5, h6 {
  margin: 0;
  padding: 0; }

ul {
  padding: 0; }

#settings-prompt {
  margin: 10px 0;
}

#error-display {
  padding: 10px;
}

#insert-button {
  margin: 0 10px;
}

.clearfix {
  display: block;
  clear: both;
  height: 0; }

.pointerCursor {
  cursor: pointer; }

.invisible {
  visibility: hidden; }

.undisplayed {
  display: none; }

.ms-Icon.enlarge {
  position: relative;
  font-size: 20px;
  top: 4px; }

.ms-ListItem-secondaryText,
.ms-ListItem-tertiaryText {
  padding-left: 15px;
}

.ms-landing-page {
  display: -webkit-flex;
  display: flex;
  -webkit-flex-direction: column;
          flex-direction: column;
  -webkit-flex-wrap: nowrap;
          flex-wrap: nowrap;
  height: 100%; }
  .ms-landing-page__main {
    display: -webkit-flex;
    display: flex;
    -webkit-flex-direction: column;
            flex-direction: column;
    -webkit-flex-wrap: nowrap;
            flex-wrap: nowrap;
    -webkit-flex: 1 1 0;
            flex: 1 1 0;
    height: 100%; }

  .ms-landing-page__content {
    display: -webkit-flex;
    display: flex;
    -webkit-flex-direction: column;
            flex-direction: column;
    -webkit-flex-wrap: nowrap;
            flex-wrap: nowrap;
    height: 100%;
    -webkit-flex: 1 1 0;
            flex: 1 1 0;
    padding: 20px; }
    .ms-landing-page__content h2 {
      margin-bottom: 20px; }
  .ms-landing-page__footer {
    display: -webkit-inline-flex;
    display: inline-flex;
    -webkit-justify-content: center;
            justify-content: center;
    -webkit-align-items: center;
            align-items: center; }
    .ms-landing-page__footer--left {
      transition: background ease 0.1s, color ease 0.1s;
      display: -webkit-inline-flex;
      display: inline-flex;
      -webkit-justify-content: flex-start;
              justify-content: flex-start;
      -webkit-align-items: center;
              align-items: center;
      -webkit-flex: 1 0 0px;
              flex: 1 0 0px;
      padding: 20px; }
      .ms-landing-page__footer--left:active {
        cursor: default; }
      .ms-landing-page__footer--left--disabled {
        opacity: 0.6;
        pointer-events: none;
        cursor: not-allowed; }
        .ms-landing-page__footer--left--disabled:active, .ms-landing-page__footer--left--disabled:hover {
          background: transparent; }
      .ms-landing-page__footer--left img {
        width: 40px;
        height: 40px; }
      .ms-landing-page__footer--left h1 {
        -webkit-flex: 1 0 0px;
                flex: 1 0 0px;
        margin-left: 15px;
        text-align: left;
        width: auto;
        max-width: auto;
        overflow: hidden;
        white-space: nowrap;
        text-overflow: ellipsis; }
    .ms-landing-page__footer--right {
      transition: background ease 0.1s, color ease 0.1s;
      padding: 29px 20px; }
      .ms-landing-page__footer--right:active, .ms-landing-page__footer--right:hover {
        background: #005ca4;
        cursor: pointer; }
      .ms-landing-page__footer--right:active {
        background: #005ca4; }
      .ms-landing-page__footer--right--disabled {
        opacity: 0.6;
        pointer-events: none;
        cursor: not-allowed; }
        .ms-landing-page__footer--right--disabled:active, .ms-landing-page__footer--right--disabled:hover {
          background: transparent; }
```

### <a name="specify-the-javascript-for-the-task-pane"></a>Указание JavaScript для области задач

В созданном вами проекте область задач JavaScript указана в файле **./src/taskpane/taskpane.js**. Откройте этот файл и замените все содержимое приведенным ниже кодом.

```js
(function(){
  'use strict';

  var config;
  var settingsDialog;

  Office.initialize = function(reason){

    jQuery(document).ready(function(){

      config = getConfig();

      // Check if add-in is configured.
      if (config && config.gitHubUserName) {
        // If configured, load the gist list.
        loadGists(config.gitHubUserName);
      } else {
        // Not configured yet.
        $('#not-configured').show();
      }

      // When insert button is selected, build the content
      // and insert into the body.
      $('#insert-button').on('click', function(){
        var gistId = $('.ms-ListItem.is-selected').val();
        getGist(gistId, function(gist, error) {
          if (gist) {
            buildBodyContent(gist, function (content, error) {
              if (content) {
                Office.context.mailbox.item.body.setSelectedDataAsync(content,
                  {coercionType: Office.CoercionType.Html}, function(result) {
                    if (result.status === Office.AsyncResultStatus.Failed) {
                      showError('Could not insert gist: ' + result.error.message);
                    }
                });
              } else {
                showError('Could not create insertable content: ' + error);
              }
            });
          } else {
            showError('Could not retrieve gist: ' + error);
          }
        });
      });

      // When the settings icon is selected, open the settings dialog.
      $('#settings-icon').on('click', function(){
        // Display settings dialog.
        var url = new URI('dialog.html').absoluteTo(window.location).toString();
        if (config) {
          // If the add-in has already been configured, pass the existing values
          // to the dialog.
          url = url + '?gitHubUserName=' + config.gitHubUserName + '&defaultGistId=' + config.defaultGistId;
        }

        var dialogOptions = { width: 20, height: 40, displayInIframe: true };

        Office.context.ui.displayDialogAsync(url, dialogOptions, function(result) {
          settingsDialog = result.value;
          settingsDialog.addEventHandler(Office.EventType.DialogMessageReceived, receiveMessage);
          settingsDialog.addEventHandler(Office.EventType.DialogEventReceived, dialogClosed);
        });
      })
    });
  };

  function loadGists(user) {
    $('#error-display').hide();
    $('#not-configured').hide();
    $('#gist-list-container').show();

    getUserGists(user, function(gists, error) {
      if (error) {

      } else {
        $('#gist-list').empty();
        buildGistList($('#gist-list'), gists, onGistSelected);
      }
    });
  }

  function onGistSelected() {
    $('#insert-button').removeAttr('disabled');
    $('.ms-ListItem').removeClass('is-selected').removeAttr('checked');
    $(this).children('.ms-ListItem').addClass('is-selected').attr('checked', 'checked');
  }

  function showError(error) {
    $('#not-configured').hide();
    $('#gist-list-container').hide();
    $('#error-display').text(error);
    $('#error-display').show();
  }

  function receiveMessage(message) {
    config = JSON.parse(message.message);
    setConfig(config, function(result) {
      settingsDialog.close();
      settingsDialog = null;
      loadGists(config.gitHubUserName);
    });
  }

  function dialogClosed(message) {
    settingsDialog = null;
  }
})();
```

### <a name="test-the-insert-gist-button"></a>Тестирование кнопки Insert gist (Вставить gist)

Сохраните все изменения и выполните в командной строке команду `npm start`, если сервер еще не запущен. Затем выполните указанные ниже действия, чтобы протестировать кнопку **Insert gist** (Вставить gist).

1. Откройте Outlook и создайте новое сообщение.

1. В окне создания сообщения нажмите кнопку **Insert gist** (Вставить gist). Справа от формы создания сообщения должна открыться область задач.

1. В области задач выберите элемент gist **Hello World Html** и нажмите кнопку **Insert** (Вставить) для вставки этого элемента gist в текст сообщения.

![Снимок экрана: область задач надстройки и выделенное содержимое элемента gist, отображаемое в тексте сообщения.](../images/addin-taskpane.png)

## <a name="next-steps"></a>Дальнейшие действия

С помощью этого руководства вы выполнили сборку надстройки Outlook, которую можно использовать в режиме создания сообщения для вставки содержимого в его текст. Чтобы узнать больше о разработке надстроек Outlook, перейдите к следующей статье.

> [!div class="nextstepaction"]
> [API надстроек Outlook](../outlook/apis.md)
