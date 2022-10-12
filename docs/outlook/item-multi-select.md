---
title: Активация надстройки Outlook для нескольких сообщений (предварительная версия)
description: Узнайте, как активировать надстройку Outlook при выборе нескольких сообщений.
ms.topic: article
ms.date: 10/10/2022
ms.localizationpriority: medium
ms.openlocfilehash: 335ee2303bfff9c5a4193e863c626e11133fa8fb
ms.sourcegitcommit: a2df9538b3deb32ae3060ecb09da15f5a3d6cb8d
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 10/12/2022
ms.locfileid: "68541293"
---
# <a name="activate-your-outlook-add-in-on-multiple-messages-preview"></a>Активация надстройки Outlook для нескольких сообщений (предварительная версия)

Благодаря функции выбора нескольких элементов надстройка Outlook теперь может активировать и выполнять операции с несколькими выбранными сообщениями за один раз. Некоторые операции, такие как отправка сообщений в систему управления отношениями с клиентами (CRM) или классификация множества элементов, теперь можно легко выполнить одним щелчком мыши.

В следующих разделах показано, как настроить надстройку для получения строки темы нескольких сообщений в режиме чтения.

> [!IMPORTANT]
> Функция выбора элементов доступна только в предварительной версии с подпиской Microsoft 365 в Outlook для Windows. Функции в предварительной версии не следует использовать в рабочих надстройки. Мы предлагаем вам протестировать эту функцию в средах тестирования или разработки и оставить отзыв о вашем интерфейсе с помощью GitHub (см. раздел **отзывов** в конце этой страницы).

> [!NOTE]
> Функция выбора нескольких элементов в настоящее время не поддерживается в манифесте [Teams (предварительная версия),](../develop/json-manifest-overview.md) но команда компонента работает над тем, чтобы сделать это доступным.

## <a name="prerequisites-to-preview-item-multi-select"></a>Предварительные требования для предварительного просмотра нескольких элементов

Чтобы просмотреть функцию с несколькими выборами, установите Outlook в Windows, начиная с версии 2209 (сборка 15629.20110). После установки присоединитесь к программе [предварительной](https://insider.office.com/join/windows) оценки Office и выберите параметр **бета-канала** , чтобы получить доступ к бета-сборкам Office.

## <a name="set-up-your-environment"></a>Настройка среды

Краткое [руководство по созданию](../quickstarts/outlook-quickstart.md?tabs=yeomangenerator) проекта надстройки Outlook с помощью генератора [Yeoman для надстроек Office](../develop/yeoman-generator-overview.md).

## <a name="configure-the-manifest"></a>Настройка манифеста

Чтобы включить активацию надстройки для нескольких выбранных сообщений, необходимо добавить дочерний элемент [SupportsMultiSelect](/javascript/api/manifest/action?view=outlook-js-preview&preserve-view=true#supportsmultiselect) **\<Action\>** в элемент и задать его значение `true`. Так как элемент с множественным выделением в данный момент поддерживает только сообщения, **\<ExtensionPoint\>** `xsi:type` значение атрибута элемента должно быть установлено как "или `MessageReadCommandSurface` " `MessageComposeCommandSurface`.

1. В предпочитаемом редакторе кода откройте созданный проект быстрого запуска Outlook.

1. Откройте **manifest.xmlфайл** , расположенный в корне проекта.

1. Назначьте **\<Permissions\>** элементу `ReadWriteMailbox` значение.

    ```xml
    <Permissions>ReadWriteMailbox</Permissions>
    ```

1. Выделите весь **\<VersionOverrides\>** узел и замените его следующим XML-кодом.

    ```xml
    <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides" xsi:type="VersionOverridesV1_0">
        <VersionOverrides xmlns="http://schemas.microsoft.com/office/mailappversionoverrides/1.1" xsi:type="VersionOverridesV1_1">
            <Requirements>
                <bt:Sets DefaultMinVersion="1.12">
                  <bt:Set Name="Mailbox"/>
                </bt:Sets>
            </Requirements>
            <Hosts>
                <Host xsi:type="MailHost">
                    <DesktopFormFactor>
                        <!-- Message Read mode-->
                        <ExtensionPoint xsi:type="MessageReadCommandSurface">
                            <OfficeTab id="TabDefault">
                                <Group id="msgReadGroup">
                                    <Label resid="GroupLabel"/>
                                    <Control xsi:type="Button" id="msgReadOpenPaneButton">
                                        <Label resid="TaskpaneButton.Label"/>
                                        <Supertip>
                                            <Title resid="TaskpaneButton.Label"/>
                                            <Description resid="TaskpaneButton.Tooltip"/>
                                        </Supertip>
                                        <Icon>
                                            <bt:Image size="16" resid="Icon.16x16"/>
                                            <bt:Image size="32" resid="Icon.32x32"/>
                                            <bt:Image size="80" resid="Icon.80x80"/>
                                        </Icon>
                                        <Action xsi:type="ShowTaskpane">
                                            <SourceLocation resid="Taskpane.Url"/>
                                            <!-- Enables your add-in to activate on multiple selected messages. -->
                                            <SupportsMultiSelect>true</SupportsMultiSelect>
                                        </Action>
                                    </Control>
                                </Group>
                            </OfficeTab>
                        </ExtensionPoint>
                    </DesktopFormFactor>
                </Host>
            </Hosts>
            <Resources>
                <bt:Images>
                  <bt:Image id="Icon.16x16" DefaultValue="https://localhost:3000/assets/icon-16.png"/>
                  <bt:Image id="Icon.32x32" DefaultValue="https://localhost:3000/assets/icon-32.png"/>
                  <bt:Image id="Icon.80x80" DefaultValue="https://localhost:3000/assets/icon-80.png"/>
                </bt:Images>
                <bt:Urls>
                  <bt:Url id="Taskpane.Url" DefaultValue="https://localhost:3000/taskpane.html"/>
                </bt:Urls>
                <bt:ShortStrings>
                  <bt:String id="GroupLabel" DefaultValue="Item Multi-select"/>
                  <bt:String id="TaskpaneButton.Label" DefaultValue="Show Taskpane"/>
                </bt:ShortStrings>
                <bt:LongStrings>
                  <bt:String id="TaskpaneButton.Tooltip" DefaultValue="Opens a pane which displays an option to retrieve the subject line of selected messages."/>
                </bt:LongStrings>
            </Resources>
        </VersionOverrides>
    </VersionOverrides>
    ```

1. Сохраните изменения.

## <a name="configure-the-task-pane"></a>Настройка области задач

Выбор нескольких элементов зависит от события [SelectedItemsChanged](/javascript/api/office/office.eventtype?view=outlook-js-preview&preserve-view=true) , чтобы определить, когда выбираются или отменяются выбор сообщений. Для этого события требуется реализация области задач.

1. В **папке ./src/taskpane** откройте **taskpane.html**.

1. В элементе **\<script\>** задайте для атрибута `src` значение . `"https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"`. Он ссылается на бета-версию библиотеки в сети доставки содержимого (CDN).

    ```html
    <script type="text/javascript" src="https://appsforoffice.microsoft.com/lib/beta/hosted/office.js"></script>
    ```

1. В элементе **\<body\>** замените весь элемент **\<main\>** следующей разметкой.

    ```html
    <main id="app-body" class="ms-welcome__main">
        <h2 class="ms-font-xl">Retrieve the subject line of multiple messages with one click!</h2>
        <ul id="selected-items"></ul>
        <div role="button" id="run" class="ms-welcome__action ms-Button ms-Button--hero ms-font-xl">
            <span class="ms-Button-label">Run</span>
        </div>
    </main>
    ```

1. Сохраните изменения.

## <a name="implement-a-handler-for-the-selecteditemschanged-event"></a>Реализация обработчика для события SelectedItemsChanged

Чтобы оповещение надстройки `SelectedItemsChanged` о возникновении события, необходимо зарегистрировать обработчик событий с помощью метода `addHandlerAsync` .

1. В **папке ./src/taskpane** откройте **taskpane.js**.

1. В функции `Office.onReady()` обратного вызова замените существующий код следующим кодом:

    ```javascript
    if (info.host === Office.HostType.Outlook) {
        document.getElementById("sideload-msg").style.display = "none";
        document.getElementById("app-body").style.display = "flex";
        document.getElementById("run").onclick = run;
    
        // Register an event handler to identify when messages are selected.
        Office.context.mailbox.addHandlerAsync(Office.EventType.SelectedItemsChanged, run, asyncResult => {
          if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;
          }
    
          console.log("Event handler added.");
        });
    }
    ```

## <a name="retrieve-the-subject-line-of-selected-messages"></a>Получение строки темы выбранных сообщений

Теперь, когда вы зарегистрировали обработчик событий, вызовите метод [getSelectedItemsAsync](/javascript/api/outlook/office.mailbox?view=outlook-js-preview&preserve-view=true#outlook-office-mailbox-getSelectedItemsAsync-member(1)) , чтобы получить строку темы выбранных сообщений и зарегистрировать их в области задач. Этот `getSelectedItemsAsync` метод также можно использовать для получения других свойств сообщения, таких как идентификатор элемента, тип элемента (`Message` является единственным поддерживаемым типом в настоящее время) и режим элемента (`Read` или `Compose`).

1. В **taskpane.js** перейдите к функции `run` и вставьте следующий код.

    ```javascript
    // Clear list of previously selected messages, if any.
    const list = document.getElementById("selected-items");
    while (list.firstChild) {
        list.removeChild(list.firstChild);
    }

    // Retrieve the subject line of the selected messages and log it to a list in the task pane.
    Office.context.mailbox.getSelectedItemsAsync(asyncResult => {
        if (asyncResult.status === Office.AsyncResultStatus.Failed) {
            console.log(asyncResult.error.message);
            return;      
        }

        asyncResult.value.forEach(item => {
            const listItem = document.createElement("li");
            listItem.textContent = item.subject;
            list.appendChild(listItem);
        });
    });
    ```

1. Сохраните изменения.

## <a name="try-it-out"></a>Проверка

1. В окне терминала выполните следующий код в корневом каталоге проекта. При этом запускается локальный веб-сервер и выполняется загрузка неопубликоваемой надстройки.

    ```command line
    npm start
    ```

    > [!TIP]
    > Если ваша надстройка не выполняет автоматическую загрузку неопубликованных приложений, следуйте инструкциям в статье "Загрузка неопубликованных надстроек [Outlook](sideload-outlook-add-ins-for-testing.md?tabs=windows#outlook-on-the-desktop) " для тестирования, чтобы вручную загрузить неопубликованное приложение в Outlook.

1. В Outlook убедитесь, что область чтения включена. Чтобы включить область чтения, см. раздел "Использование и настройка области чтения для [предварительного просмотра сообщений"](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0).

1. Перейдите в папку "Входящие" и выберите несколько сообщений, удерживая **нажатой клавишу CTRL** при выборе сообщений.

1. Выберите **"Показать область задач** " на ленте.

1. В области задач выберите "Выполнить", чтобы просмотреть список строк тем выбранных сообщений.

    :::image type="content" source="../images/outlook-multi-select.png" alt-text="Пример списка строк темы, полученных из нескольких выбранных сообщений.":::

## <a name="item-multi-select-behavior-and-limitations"></a>Поведение и ограничения для нескольких элементов

Элемент с множественным выделением поддерживает только сообщения в почтовом ящике Exchange в режимах чтения и создания. Надстройка Outlook активируется только для нескольких сообщений, если выполняются следующие условия.

- Сообщения должны быть выбраны из одного почтового ящика Exchange за раз. Почтовые ящики, отличные от Exchange, не поддерживаются.
- Сообщения должны быть выбраны из одной папки почтового ящика за раз. Надстройка не активируется для нескольких сообщений, если они находятся в разных папках, если не включено представление "Беседы". Дополнительные сведения см. в [разделе "Множественный выбор в беседах"](#multi-select-in-conversations).
- Чтобы определить событие `SelectedItemsChanged` , надстройка должна реализовать область задач.
- Область [чтения в](https://support.microsoft.com/office/2fd687ed-7fc4-4ae3-8eab-9f9b8c6d53f0) Outlook должна быть включена.
- Одновременно можно выбрать не более 100 сообщений.

> [!NOTE]
> Приглашения и ответы на собрания считаются сообщениями, а не встречами, поэтому их можно включить в выборку.

### <a name="multi-select-in-conversations"></a>Множественный выбор в беседах

Элемент multi-select поддерживает [представление бесед](https://support.microsoft.com/office/0eeec76c-f59b-4834-98e6-05cfdfa9fb07) , включено ли оно в почтовом ящике или в определенных папках. В следующей таблице описано ожидаемое поведение при развертывании или свертывании бесед, при выборе заголовка беседы и при размещении сообщений беседы в папке, которая находится в другой папке.

|Selection|Развернутое представление беседы|Свернутое представление беседы|
|------|------|------|
|**Выбран заголовок беседы**|Если выбран только заголовок беседы, надстройка, поддерживающие множественный выбор, не активируется. Однако если также выбраны другие сообщения, отличные от заголовков, надстройка активируется только для них, а не для выбранного заголовка.|Самое новое сообщение (то есть первое сообщение в стеке бесед) включается в выбор сообщения.<br><br>Если самое новое сообщение в беседе находится в другой папке из текущего представления, в выбор включается следующее сообщение в стеке, расположенном в текущей папке.|
|**Выбранные сообщения беседы находятся в той же папке, что и в настоящее время в представлении**|Все выбранные сообщения беседы включаются в выборку.|Не применимо. Для выбора в свернутом представлении беседы доступен только заголовок беседы.|
|**Выбранные сообщения беседы находятся в разных папках, чем в настоящее время в представлении** |Все выбранные сообщения беседы включаются в выборку.|Не применимо. Для выбора в свернутом представлении беседы доступен только заголовок беседы.|

## <a name="next-steps"></a>Дальнейшие действия

Теперь, когда вы включили надстройку для работы с несколькими выбранными сообщениями, вы можете расширить возможности надстройки и улучшить взаимодействие с пользователем. Изучите выполнение более сложных операций, используя идентификаторы элементов выбранных сообщений со службами, такими как [веб-службы Exchange (EWS)](web-services.md) и [Microsoft Graph](/graph/overview).

## <a name="see-also"></a>См. также

- [Манифесты надстроек Outlook](manifests.md)
- [Вызов веб-служб из надстройки Outlook](web-services.md)
- [Обзор Microsoft Graph](/graph/overview)
