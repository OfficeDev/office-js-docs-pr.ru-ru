---
title: Использование Office Dialog API в вашей надстройках Office
description: Узнайте основы создания диалоговых окне в Office надстройке.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 03859956fbc005f33212343b83470d4f91ef65ee
ms.sourcegitcommit: 287a58de82a09deeef794c2aa4f32280efbbe54a
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/28/2022
ms.locfileid: "64496833"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Использование Office Dialog API в надстройках Office

Вы можете использовать [Office dialog API](/javascript/api/office/office.ui), чтобы открывать диалоговые окна в надстройке Office. Эта статья содержит инструкции по использованию dialog API в надстройке Office.

> [!NOTE]
> Сведения о том, где поддерживается API диалогов, см. в наборе требований [к API диалогов](/javascript/api/requirement-sets/common/dialog-api-requirement-sets). API диалогов в настоящее время поддерживается для Excel, PowerPoint и Word. Outlook поддержка включается в различные&mdash; наборы требований к почтовым ящикам, дополнительные сведения по ссылке API.

Основной сценарий для Dialog API - включить аутентификацию с помощью таких ресурсов, как Google, Facebook или Microsoft Graph. Дополнительные сведения см. в статье [Проверка подлинности с помощью Office Dialog API](auth-with-office-dialog-api.md) *после* ознакомления с текущей статьей.

Возможность открытия диалогового окна с помощью области задач, контентной надстройки или [команды надстройки](../design/add-in-commands.md) может позволить следующее:

- Отображение страниц входных входов, которые не могут быть открыты непосредственно в области задач.
- предоставить больше места на экране (или даже весь экран) для некоторых задач в надстройке;
- разместить видео, которое будет слишком маленьким в области задач.

> [!NOTE]
> Поскольку перекрывающиеся элементы пользовательского интерфейса не приветствуются, избегайте открытия диалогового окна на панели задач, если это не требуется в сценарий. При планировании контактной зоны помните, что в области задач можно использовать вкладки. Пример области задач на вкладке см. в примере Excel [JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

На приведенном ниже изображении показан пример диалогового окна. 

![Снимок экрана, показывающий диалоговое окно с 3 вариантами входа, отображаемой перед Word.](../images/auth-o-dialog-open.png)

Обратите внимание, что диалоговое окно всегда открывается в центре экрана. Пользователь может перемещать ее и изменять ее размер. Окно является *немодальным*— пользователь может продолжать взаимодействовать как с документом в приложении Office, так и со страницей в области задач, если она есть.

## <a name="open-a-dialog-box-from-a-host-page"></a>Откройте диалоговое окно с главной страницы

Office JavaScript API включает в себя [Диалоговый](/javascript/api/office/office.dialog) объекта и две функции в [Office.context.ui namespace](/javascript/api/office/office.ui).

Чтобы открыть диалоговое окно, ваш код, обычно страница в панели задач, вызывает метод [displayDialogAsync](/javascript/api/office/office.ui) и передает ему URL-адрес ресурса, который вам нужно открыть. Страница, на которой этот метод вызван, называется "главной страницей". Например, если вы вызываете этот метод в скрипте для index.html на панели задач, то index.html - это главная страница диалогового окна, которое открывает метод.

Ресурс, который открывается в диалоговом окне, обычно представляет собой страницу, но это может быть метод контроллера в приложении MVC, маршрут, метод веб-службы или любой другой ресурс. В этой статье "страница" или "веб-сайт" ссылается на ресурс в диалоговом окне. Следующий код — простой пример.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - В случае URL-адреса используется протокол HTTP **S**, обязательный для всех страниц, загружаемых в диалоговом окне, а не только для первой страницы.
> - Домен диалогового окна совпадает с доменом главной страницы, которая может быть страницей в панели задач или [файлом функции](/javascript/api/manifest/functionfile) команды надстройки. Страница, метод контроллера или другой ресурс, передаваемый в метод `displayDialogAsync`, должен быть в том же домене, что и страница ведущего приложения.

> [!IMPORTANT]
> Главная страница и ресурс, который открывается в диалоговом окне, должны иметь один и тот же полный домен. Если вы попробуете передать поддомен домена надстройки в `displayDialogAsync`, ничего не получится. Полные доменные имена, включая поддомены, должны совпадать.

После загрузки первой страницы (или другого ресурса) пользователь может использовать ссылки или другой пользовательский интерфейс для перехода на любой веб-сайт (или другой ресурс), использующий HTTPS. Первая страница также может сразу перенаправлять пользователя на другой сайт.

По умолчанию диалоговое окно занимает 80 % высоты и ширины экрана устройства, но вы можете установить другие соотношения путем передачи объекта конфигурации в метод, как показано в приведенном ниже примере.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Пример надстройки, в которой используется этот метод, см. [здесь](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example). Дополнительные примеры, которые используются `displayDialogAsync`, см. [в примере](#samples).

Установите оба значения равными 100 %, чтобы надстройка открывалась во весь экран. (На самом деле, максимальное значение составляет 99,5 %, возможность перемещать окно и изменять его размер сохраняется.)

> [!NOTE]
> В окне хост можно открыть только одно диалоговое окно. Попытка открыть другое диалоговое окно создает ошибку. Например, если пользователь открывает диалоговое окно из области задач, он не может открыть второе диалоговое окно с другой страницы в области задач. Однако, когда диалоговое окно открывается из команды надстройки, команда открывает новый (но невидимый) [HTML-файл](../design/add-in-commands.md) при каждом выборе. Это создает новое (невидимое) окно хоста, поэтому каждое такое окно может запускать свое собственное диалоговое окно. Дополнительные сведения см. [в дополнительных сведениях об ошибках с displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Использование параметра производительности в Office в Интернете

`displayInIframe` — дополнительное свойство в объекте конфигурации, которое можно передать `displayDialogAsync`. Когда этому свойству присвоено значение `true`, а надстройка запущена для документа в Office в Интернете, диалоговое окно будет открываться быстрее, потому что будет выступать как плавающий фрейм iframe. Ниже приведен пример.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

Значение по умолчанию: `false`. Его использование равнозначно пропуску всего свойства. Если надстройка не работает в Office в Интернете, `displayInIframe` она игнорируется.

> [!NOTE]
> Не следует **использовать** , `displayInIframe: true` если диалоговое окно в любой момент перенаправляется на страницу, которую нельзя открыть в iframe. Например, вход на страницы многих популярных веб-служб, таких как учетные записи Google и Microsoft, нельзя открыть в iframe.

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Отправка сведений из диалогового окна главной странице

> [!NOTE]
>
> - Для ясности в этом разделе мы называем сообщение целевым на хост-странице *, но* строго говоря, сообщения идут на время запуска *JavaScript* в области задач (или время запуска, в которое размещен файл функций [).](/javascript/api/manifest/functionfile) Различие имеет важное значение только в случае меж доменных сообщений. Дополнительные сведения см. в разделе [Междоменные сообщения в основной среде выполнения](#cross-domain-messaging-to-the-host-runtime).
> - Диалоговое окно не может общаться с хост-страницей в области задач, если Office библиотека API JavaScript не загружена на страницу. (Как и на любой странице, Office библиотеке API JavaScript, сценарий страницы должен инициализировать надстройки. Дополнительные сведения см. в [материале Initialize your Office надстройки](initialize-add-in.md).)

Код в диалоговом окне использует [функцию messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) для отправки строки сообщения на хост-страницу. Строка может быть словом, предложением, BLOB XML, строками JSON или другими строками, которые можно сериализировать в строку или отбрасовать в строку. Ниже приведен пример.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
> - Эта `messageParent` функция является *одним из двух* Office API JS, которые можно назвать в диалоговом окне.
> - Другой API JS, который можно назвать в диалоговом окне, — `Office.context.requirements.isSetSupported`. Сведения об этом см. в [Office приложениях и требованиях API](specify-office-hosts-and-api-requirements.md). Однако в диалоговом окне этот API не поддерживается Outlook 2016 одноразовой покупке (то есть версии MSI).

В следующем примере `googleProfile` — это строковое представление профиля Google пользователя.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

Чтобы главная страница получила сообщение, ее необходимо настроить. Для этого добавьте параметр обратного вызова в исходный вызов метода `displayDialogAsync`. Обратный вызов назначает обработчик событию `DialogMessageReceived`. Ниже приведен пример.

```js
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20},
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);
```

> [!NOTE]
>
> - Office передает объект [AsyncResult](/javascript/api/office/office.asyncresult) в функцию обратного вызова. Он представляет результат попытки открыть диалоговое окно. Он не представляет результат событий в диалоговом окне. Подробнее об этом различии см. в [Обработка ошибок и событий](dialog-handle-errors-events.md). 
> - Для свойства `value` объекта `asyncResult` задан объект [Dialog](/javascript/api/office/office.dialog), который существует на главной странице, а не в контексте выполнения диалогового окна.
> - `processMessage` — это функция, которая обрабатывает событие. Вы можете присвоить ей любое имя.
> - Переменная `dialog` объявляется в более широком контексте, чем обратный вызов, так как на нее также ссылается `processMessage`.

Ниже приведен простой пример обработчика для события `DialogMessageReceived`.

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
>
> - Office передает объект `arg` в обработчик. Его `message` свойство — строка, отправленная вызовом `messageParent` в диалоговом окне. В этом примере это строковая репрезентация профиля пользователя из службы, например учетной записи Майкрософт или Google, `JSON.parse`поэтому она десервализована обратно к объекту с .
> - Реализация `showUserName` не отображается. Она может отображать персонализированное приветствие в области задач.

Когда взаимодействие пользователя с диалоговым окном закончится, обработчик сообщений должен закрыть диалоговое окно, как показано в этом примере.

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
>
> - Объект `dialog` должен быть таким же, как объект, который возвращается при вызове `displayDialogAsync`.
> - Вызов метода `dialog.close` дает указание Office немедленно закрыть диалоговое окно.

Пример надстройки, в которой используются эти методы, см. в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Если надстройка должна открыть другую страницу области задач после получения сообщения, можно использовать метод `window.location.replace` (или `window.location.href`) в последней строке обработчика. Ниже приведен пример.

```js
function processMessage(arg) {
    // message processing code goes here;
    window.location.replace("/newPage.html");
    // Alternatively ...
    // window.location.href = "/newPage.html";
}
```

Пример подобной надстройки см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

### <a name="conditional-messaging"></a>Условные сообщения

Так как из диалогового окна можно отправить несколько вызовов `messageParent`, но на главной странице есть только один обработчик для события `DialogMessageReceived`, обработчику необходимо использовать условную логику, чтобы различать сообщения. Например, если диалоговое окно побуждает пользователя войти к поставщику удостоверений, например учетной записи Майкрософт или Google, он отправляет профиль пользователя в качестве сообщения. Если проверка подлинности не удается, диалоговое окно отправляет сведения об ошибках на хост-страницу, как в следующем примере.

```js
if (loginSuccess) {
    var userProfile = getProfile();
    var messageObject = {messageType: "signinSuccess", profile: userProfile};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    var errorDetails = getError();
    var messageObject = {messageType: "signinFailure", error: errorDetails};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
>
> - Переменная `loginSuccess` будет инициализирована после считывания отклика HTTP от поставщика удостоверений.
> - Реализация функций `getProfile` и `getError` не показана. Они получают данные из параметра запроса или ответа HTTP.
> - В зависимости от того, удалось ли выполнить вход, отправляются анонимные объекты различных типов. Оба содержат свойство `messageType`, но один содержит свойство `profile`, а другой — свойство `error`.

Код обработчика на главной странице использует значение свойства `messageType` для разветвления, как показано в приведенном ниже примере. Обратите внимание на то, что здесь используется та же функция `showUserName`, что и в примере выше, а функция `showNotification` отображает сообщение об ошибке в элементе пользовательского интерфейса на главной странице.

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "signinSuccess") {
        dialog.close();
        showUserName(messageFromDialog.profile.name);
        window.location.replace("/newPage.html");
    } else {
        dialog.close();
        showNotification("Unable to authenticate user: " + messageFromDialog.error);
    }
}
```

> [!NOTE]
> Реализация `showNotification` не отображается в примере кода, предоставленного этой статьей. Пример возможного способа реализации этой функции в своей надстройке см. в статье [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

### <a name="cross-domain-messaging-to-the-host-runtime"></a>Меж доменная передача сообщений в хост-время работы

Диалоговое окно или родительское время запуска JavaScript (либо в области задач, либо в пользовательском интерфейсе, в котором размещен файл функций) может перемещаться в стороне от домена надстройки после открытия диалоговое окно. Если что-либо из этих вещей произошло, `messageParent` вызов сбой, если в коде не указан домен родительского времени запуска. Это необходимо, добавив параметр [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) в вызов `messageParent`. Этот объект имеет свойство `targetOrigin` , которое указывает домен, в который должно быть отправлено сообщение. Если параметр не используется, Office предполагает, что целью является тот же домен, что и диалоговое окно.

> [!NOTE]
> Для `messageParent` отправки меж доменного сообщения требуется набор требований [Диалоговое начало 1.1](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets). Параметр `DialogMessageOptions` игнорируется в старых версиях Office, которые не поддерживают набор требований, поэтому поведение метода не влияет на его пропуск.

Ниже приводится пример использования для `messageParent` отправки меж доменного сообщения.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "https://resource.contoso.com" });
```

> [!NOTE]
> Параметр `DialogMessageOptions` был выпущен приблизительно 19 июля 2021 г. Примерно через 30 `messageParent` `DialogMessageOptions` дней после этой даты в Office в Интернете, когда первый раз называется без параметра, а родительский домен отличается от диалогового, пользователю будет предложено утвердить отправку данных в целевой домен. Если пользователь одобряет, ответ пользователя кэшется в течение 24 часов. В этот период, `messageParent` когда он вызван с тем же целевым доменом, пользователю больше не будет предложено.

Если в сообщении не содержатся конфиденциальные данные, `targetOrigin` можно установить "\*", что позволяет отправлять его в любой домен. Ниже приведен пример.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "*" });
```

> [!TIP]
> Параметр `DialogMessageOptions` был добавлен в метод в `messageParent` качестве обязательного параметра в середине 2021 г. Старые надстройки, отправив сообщение с помощью метода, перестают работать до тех пор, пока не будут обновлены для использования нового параметра. Пока надстройка не будет обновлена, только Office для *Windows* пользователи и системные администраторы могут включить эти надстройки для продолжения работы, указав доверенный домен (ы) с параметром реестра: **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Для этого создайте `.reg` файл с расширением, сохраните его на Windows, а затем дважды щелкните его, чтобы запустить его. Ниже приводится пример содержимого такого файла.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="pass-information-to-the-dialog-box"></a>Передача данных диалоговому окну

Ваша надстройка может отправлять сообщения с хост-страницы в диалоговое окно с помощью [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)).[](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page)

### <a name="use-messagechild-from-the-host-page"></a>Использование `messageChild()` на хост-странице

При вызове Office диалогового API для открытия диалоговое окно возвращается объект [Диалог](/javascript/api/office/office.dialog). Она должна быть назначена переменной, которая имеет больше возможностей, чем метод [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) , так как объект будет ссылаться на другие методы. Ниже приведен пример.

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

Этот `Dialog` объект имеет [метод messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) , который отправляет в диалоговое окно любую строку, включая строковые данные. Это вызывает событие `DialogParentMessageReceived` в диалоговом окне. Код должен обрабатывать это событие, как показано в следующем разделе.

Рассмотрим сценарий, в котором пользовательский интерфейс диалогов связан с текущим активным таблицой и положением этого таблицы по отношению к другим таблицам. В следующем примере `sheetPropertiesChanged` отправляет свойства Excel таблицы в диалоговое окно. В этом случае текущий лист называется "Мой лист" и это второй лист в книге. Данные инкапсулированы в объекте и струнные, чтобы они могли быть переданы `messageChild`.

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Обработать диалоговое окно DialogParentMessageReceived

В диалоговом окне JavaScript зарегистрируйте `DialogParentMessageReceived` обработчиватель события методом [UI.addHandlerAsync](/javascript/api/office/office.ui#office-office-ui-addhandlerasync-member(1)) . Обычно это делается в [методах Office.onReady или Office.initialize](initialize-add-in.md), как показано ниже. (Более надежный пример ниже.)

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Затем определите обработник `onMessageFromParent` . Следующий код продолжает пример из предыдущего раздела. Обратите внимание, Office `message` передает аргумент обработнику и что свойство объекта аргумента содержит строку со страницы хост. В этом примере сообщение перенаправляется в объект, а jQuery используется для набора верхнего заголовка диалогов, чтобы соответствовать новому имени таблицы.

```javascript
function onMessageFromParent(arg) {
    var messageFromParent = JSON.parse(arg.message);
    $('h1').text(messageFromParent.name);
}
```

Это лучшая практика, чтобы убедиться, что обработник правильно зарегистрирован. Это можно сделать, передав методу `addHandlerAsync` вызов. Это выполняется при попытке зарегистрировать обработник. Используйте обработник для входа или показа ошибки, если обработник не был успешно зарегистрирован. Ниже приведен пример. Обратите внимание `reportError` , что это функция, не определенная здесь, которая регистрит или отображает ошибку.

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a>Условный обмен сообщениями с родительской страницы на диалоговое окно

Так как `messageChild` на хост-странице можно сделать несколько вызовов, `DialogParentMessageReceived` но в диалоговом окне для события есть только один обработок, обработник должен использовать условную логику, чтобы различать различные сообщения. Это можно сделать точно параллельно структуре условного обмена сообщениями, когда диалоговое окно отправляет сообщение на хост-страницу, как описано в условном [сообщении](#conditional-messaging).

> [!NOTE]
> В некоторых ситуациях `messageChild` API, который входит в набор требований [DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), может не поддерживаться. Некоторые альтернативные способы передачи сообщений из родительского окна в диалоговое окно описаны в альтернативных способах передачи сообщений в диалоговое окно со [своей хост-страницы](parent-to-dialog.md).

> [!IMPORTANT]
> Набор [требований DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets) не может быть указан в разделе Требования манифеста надстройки. Вам придется проверять поддержку DialogApi 1.2 `isSetSupported` во время запуска с помощью метода, описанного в проверках времени запуска для поддержки набора [методов и требований](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). Поддержка требований манифеста находится в стадии разработки.

### <a name="cross-domain-messaging-to-the-dialog-runtime"></a>Меж доменная передача сообщений в диалоговое время работы

Диалоговое окно или родительское время запуска JavaScript (либо в области задач, либо в пользовательском интерфейсе, в котором размещен файл функций) может перемещаться в стороне от домена надстройки после открытия диалоговое окно. Если что-либо из этих вещей произошло, `messageChild` вызов сбой, если в коде не указан домен времени запуска диалогов. Это необходимо, добавив параметр [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) в вызов `messageChild`. Этот объект имеет свойство `targetOrigin` , которое указывает домен, в который должно быть отправлено сообщение. Если параметр не используется, Office предполагает, что целью является тот же домен, что и родительское время запуска. 

> [!NOTE]
> Для `messageChild` отправки меж доменного сообщения требуется набор требований [Диалоговое начало 1.1](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets). Параметр `DialogMessageOptions` игнорируется в старых версиях Office, которые не поддерживают набор требований, поэтому поведение метода не влияет на его пропуск.

Ниже приводится пример использования для `messageChild` отправки меж доменного сообщения.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "https://resource.contoso.com" });
```

Если в сообщении не содержатся конфиденциальные данные, `targetOrigin` можно установить "\*", что позволяет отправлять *его в любой* домен. Ниже приведен пример.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "*" });
```

Так как время работы JavaScript, в котором находится диалоговое окно, не может получить доступ к разделу **AppDomains** манифеста и, таким образом, определить, является ли домен, из которого поступает сообщение, доверенным, `DialogParentMessageReceived` для этого необходимо использовать обработчиватель. Объект, который передается обработителю, содержит домен, который в настоящее время размещен в родительском качестве его `origin` свойства. Ниже приводится пример использования свойства.

```javascript
function onMessageFromParent(arg) {
    if (arg.origin === "https://addin.fabrikam.com") {
        // process message
    } else {
        dialog.close();
        showNotification("Messages from " + arg.origin + " are not accepted.");
    }
}
```

Например, в коде можно [использовать методы Office.onReady или Office.initialize](initialize-add-in.md) для хранения массива доверенных доменов в глобальной переменной. Затем `arg.origin` свойство можно проверить в обработнике с этим списком.

> [!TIP]
> Параметр `DialogMessageOptions` был добавлен в метод в `messageChild` качестве обязательного параметра в середине 2021 г. Старые надстройки, отправив сообщение с помощью метода, перестают работать до тех пор, пока не будут обновлены для использования нового параметра. Пока надстройка не будет обновлена, только Office для *Windows* пользователи и системные администраторы могут включить эти надстройки для продолжения работы, указав доверенный домен (ы) с параметром реестра: **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Для этого создайте `.reg` файл с расширением, сохраните его на Windows, а затем дважды щелкните его, чтобы запустить его. Ниже приводится пример содержимого такого файла.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="close-the-dialog-box"></a>Закройте диалоговое окно

Вы можете добавить в диалоговое окно кнопку, которая будет его закрывать. Для этого обработчик событий кнопки должен использовать функцию `messageParent`, чтобы сообщить главной странице, что кнопка нажата. Ниже приведен пример.

```js
function closeButtonClick() {
    var messageObject = {messageType: "dialogClosed"};
    var jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

Обработчик главной страницы для `DialogMessageReceived` вызовет `dialog.close`, как показано в этом примере. (См. предыдущие примеры, в которых показано, как `dialog` инициализируется объект).

```js
function processMessage(arg) {
    var messageFromDialog = JSON.parse(arg.message);
    if (messageFromDialog.messageType === "dialogClosed") {
       dialog.close();
    }
}
```

Даже если у вас нет собственного пользовательского интерфейса для закрытия диалогового окна, пользователь может закрыть диалоговое окно, выбрав **X** в правом верхнем углу. Это действие запускает событие `DialogEventReceived`. Чтобы главная область могла реагировать на это событие, для нее должен быть объявлен обработчик этого события. Дополнительные сведения см. в разделе [Ошибок и события в диалоговом окне](dialog-handle-errors-events.md#errors-and-events-in-the-dialog-box).

## <a name="advanced-topics-and-special-scenarios"></a>Продвинутые темы и специальные сценарии

### <a name="use-the-dialog-api-to-show-a-video"></a>Используйте Dialog API, чтобы показать видео

См. статью [Использование диалогового окна «Office» для отображения видео](dialog-video.md).

### <a name="use-the-dialog-apis-in-an-authentication-flow"></a>Использование Dialog API в потоке аутентификации

См. статью[ Проверка подлинности с помощью Office Dialog API ](auth-with-office-dialog-api.md).

### <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Используйте API Office диалогов с односпальтными приложениями и маршрутизами на стороне клиента

При использовании Office dialog API, SPA и маршрутизация на стороне клиента должны обрабатываться с осторожностью См. статью[Рекомендации по использованию Office dialog API в SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).

### <a name="error-and-event-handling"></a>Обработка ошибок и событий

См. статью об ошибках и событиях [Обработка ошибок и событий в Office dialog box](dialog-handle-errors-events.md).

## <a name="next-steps"></a>Дальнейшие действия

Узнайте о том, как использовать Office dialog API, в [Рекомендации по использованию Office dialog API](dialog-best-practices.md).

## <a name="samples"></a>Примеры

Все следующие примеры использования `displayDialogAsync`. Некоторые из них имеют серверы на основе NodeJS, а другие имеют ASP.NET/IIS-серверы, но логика использования метода та же, независимо от того, как реализована серверная сторона надстройки.

**Основы:**

- [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [Учебный контент / создание надстроек (несколько примеров)](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**Более сложные примеры:**

- [Office надстройки Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Office в Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [Единый вход с использованием NodeJS для надстройки Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
- [Office надстройки ASPNET SSO](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Office saAS пример монетизации надстройки](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [Outlook надстройки Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Outlook SSO надстройки](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Outlook просмотра маркеров надстройки](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Outlook надстройки Actionable Message](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [Outlook надстройки Общего доступа к OneDrive](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [PowerPoint Надстройка Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Excel общего времени работы](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario)
- [Excel надстройки ASPNET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Word Add-in JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Word Add-in JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [Word Add-in AngularJS Client OAuth](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Надстройка Office Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Office надстройки OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Office кода шаблонов дизайна надстройки UX](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
