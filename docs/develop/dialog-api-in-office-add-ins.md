---
title: Использование Office Dialog API в вашей надстройках Office
description: Основные сведения о создании диалогового окна в надстройке Office.
ms.date: 07/18/2022
ms.localizationpriority: medium
ms.openlocfilehash: 4dc1bc0b45bb41952cd2ab83fcd62633d598ab4e
ms.sourcegitcommit: 3abcf7046446e7b02679c79d9054843088312200
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 11/02/2022
ms.locfileid: "68810017"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Использование Office Dialog API в надстройках Office

Вы можете использовать [Office dialog API](/javascript/api/office/office.ui), чтобы открывать диалоговые окна в надстройке Office. Эта статья содержит инструкции по использованию dialog API в надстройке Office.

> [!NOTE]
> Сведения о том, где в настоящее время поддерживается API диалога, см. в разделе [Наборы обязательных элементов API диалога](/javascript/api/requirement-sets/common/dialog-api-requirement-sets). Api Dialog в настоящее время поддерживается для Excel, PowerPoint и Word. Поддержка Outlook включена в различные наборы&mdash;требований к почтовым ящикам, см. в справочнике по API для получения дополнительных сведений.

Основной сценарий для Dialog API - включить аутентификацию с помощью таких ресурсов, как Google, Facebook или Microsoft Graph. Дополнительные сведения см. в статье [Проверка подлинности с помощью Office Dialog API](auth-with-office-dialog-api.md) *после* ознакомления с текущей статьей.

Возможность открытия диалогового окна с помощью области задач, контентной надстройки или [команды надстройки](../design/add-in-commands.md) может позволить следующее:

- Отображение страниц входа, которые нельзя открыть непосредственно в области задач.
- предоставить больше места на экране (или даже весь экран) для некоторых задач в надстройке;
- разместить видео, которое будет слишком маленьким в области задач.

> [!NOTE]
> Поскольку перекрывающиеся элементы пользовательского интерфейса не приветствуются, избегайте открытия диалогового окна на панели задач, если это не требуется в сценарий. При планировании контактной зоны помните, что в области задач можно использовать вкладки. Пример области задач с вкладками см. [в примере Надстройка Excel JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) .

На приведенном ниже изображении показан пример диалогового окна. 

![Диалоговое окно с 3 параметрами входа, отображаемыми перед Word.](../images/auth-o-dialog-open.png)

Обратите внимание, что диалоговое окно всегда открывается в центре экрана. Пользователь может перемещать ее и изменять ее размер. Окно *немодальное*. Пользователь может продолжать взаимодействовать как с документом в приложении Office, так и со страницей в области задач, если она есть.

## <a name="open-a-dialog-box-from-a-host-page"></a>Откройте диалоговое окно с главной страницы

Office JavaScript API включает в себя [Диалоговый](/javascript/api/office/office.dialog) объекта и две функции в [Office.context.ui namespace](/javascript/api/office/office.ui).

Чтобы открыть диалоговое окно, ваш код, обычно страница в панели задач, вызывает метод [displayDialogAsync](/javascript/api/office/office.ui) и передает ему URL-адрес ресурса, который вам нужно открыть. Страница, на которой этот метод вызван, называется "главной страницей". Например, если вы вызываете этот метод в скрипте для index.html на панели задач, то index.html - это главная страница диалогового окна, которое открывает метод.

Ресурс, который открывается в диалоговом окне, обычно представляет собой страницу, но это может быть метод контроллера в приложении MVC, маршрут, метод веб-службы или любой другой ресурс. В этой статье "страница" или "веб-сайт" ссылается на ресурс в диалоговом окне. Ниже приведен простой пример кода.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
>
> - В случае URL-адреса используется протокол HTTP **S**, Обязательный для всех страниц, загружаемых в диалоговом окне, а не только для первой страницы.
> - Домен диалогового окна совпадает с доменом главной страницы, которая может быть страницей в панели задач или [файлом функции](/javascript/api/manifest/functionfile) команды надстройки. Страница, метод контроллера или другой ресурс, передаваемый в метод `displayDialogAsync`, должен быть в том же домене, что и страница ведущего приложения.

> [!IMPORTANT]
> Главная страница и ресурс, который открывается в диалоговом окне, должны иметь один и тот же полный домен. Если вы попробуете передать поддомен домена надстройки в `displayDialogAsync`, ничего не получится. Полные доменные имена, включая поддомены, должны совпадать.

После загрузки первой страницы (или другого ресурса) пользователь может использовать ссылки или другой пользовательский интерфейс для перехода на любой веб-сайт (или другой ресурс), использующий HTTPS. Первая страница также может сразу перенаправлять пользователя на другой сайт.

По умолчанию диалоговое окно занимает 80 % высоты и ширины экрана устройства, но вы можете установить другие соотношения путем передачи объекта конфигурации в метод, как показано в приведенном ниже примере.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Подобная надстройка приведена в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example). Дополнительные примеры, использующие `displayDialogAsync`, см. [в разделе Примеры](#samples).

Set both values to 100% to get what is effectively a full screen experience. (The effective maximum is 99.5%, and the window is still moveable and resizable.)

> [!NOTE]
> В окне узла можно открыть только одно диалоговое окно. При попытке открыть другое диалоговое окно возникает ошибка. Например, если пользователь открывает диалоговое окно из области задач, он не может открыть второе диалоговое окно с другой страницы в области задач. Однако при открытии диалогового окна из команды надстройки команда открывает новый (но невидимый) [HTML-файл](../design/add-in-commands.md) каждый раз при его выборе. При этом создается новое (невидимое) основное окно, поэтому каждое такое окно может запускать собственное диалоговое окно. Дополнительные сведения см. в разделе [Ошибки из displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Использование параметра производительности в Office в Интернете

`displayInIframe` — дополнительное свойство в объекте конфигурации, которое можно передать `displayDialogAsync`. Когда этому свойству присвоено значение `true`, а надстройка запущена для документа в Office в Интернете, диалоговое окно будет открываться быстрее, потому что будет выступать как плавающий фрейм iframe. Ниже приведен пример.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

Значение по умолчанию: `false`. Его использование равнозначно пропуску всего свойства. Если надстройка не запущена в Office в Интернете, `displayInIframe` она игнорируется.

> [!NOTE]
> **Не** следует использовать, `displayInIframe: true` если диалоговое окно будет в любой момент перенаправлять на страницу, которая не может быть открыта в iframe. Например, страницы входа многих популярных веб-служб, таких как Google и учетная запись Майкрософт, не могут быть открыты в iframe.

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Отправка сведений из диалогового окна главной странице

> [!NOTE]
>
> - Для ясности в этом разделе мы называем сообщение целевой *главной страницей*, но, строго говоря, сообщения отправляются в [среду выполнения](../testing/runtimes.md) в области задач (или среду выполнения, в которую размещается [файл функции](/javascript/api/manifest/functionfile)). Различие имеет важное значение только в случае обмена сообщениями между доменами. Дополнительные сведения см. в разделе [Междоменные сообщения в основной среде выполнения](#cross-domain-messaging-to-the-host-runtime).
> - Диалоговое окно не может взаимодействовать с главной страницей в области задач, если на странице не загружена библиотека API JavaScript для Office. (Как и любая страница, использующая библиотеку API JavaScript для Office, скрипт для страницы должен инициализировать надстройку. Дополнительные сведения см. [в разделе Инициализация надстройки Office](initialize-add-in.md).)

Код в диалоговом окне использует функцию [messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) для отправки строкового сообщения на хост-страницу. Строка может быть словом, предложением, BLOB-объектом XML, строкифицированным JSON или другими объектами, которые можно сериализовать в строку или привести к строке. Ниже приведен пример.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
>
> - Эта `messageParent` функция является *одним из двух* API Office JS, которые можно вызвать в диалоговом окне.
> - Другой API JS, который можно вызвать в диалоговом окне, — .`Office.context.requirements.isSetSupported` Дополнительные сведения см. в статье [Указание приложений Office и требований к API](specify-office-hosts-and-api-requirements.md). Однако в диалоговом окне этот API не поддерживается в корпоративной лицензированной бессрочной Outlook 2016 (то есть в версии MSI).

В следующем примере `googleProfile` — это строковое представление профиля Google пользователя.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

Чтобы главная страница получила сообщение, ее необходимо настроить. Для этого добавьте параметр обратного вызова в исходный вызов метода `displayDialogAsync`. Обратный вызов назначает обработчик событию `DialogMessageReceived`. Ниже приведен пример.

```js
let dialog;
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
> - The `processMessage` is the function that handles the event. You can give it any name you want.
> - Переменная `dialog` объявляется в более широком контексте, чем обратный вызов, так как на нее также ссылается `processMessage`.

Ниже приведен простой пример обработчика для события `DialogMessageReceived`.

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
    showUserName(messageFromDialog.name);
}
```

> [!NOTE]
>
> - Office передает объект `arg` в обработчик. Его `message` свойство представляет собой строку, отправляемую вызовом `messageParent` в диалоговом окне. В этом примере это строковое представление профиля пользователя из такой службы, как учетная запись Майкрософт или Google, поэтому он десериализируется обратно в объект с `JSON.parse`помощью .
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

If the add-in needs to open a different page of the task pane after receiving the message, you can use the `window.location.replace` method (or `window.location.href`) as the last line of the handler. The following is an example.

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

Так как из диалогового окна можно отправить несколько вызовов `messageParent`, но на главной странице есть только один обработчик для события `DialogMessageReceived`, обработчику необходимо использовать условную логику, чтобы различать сообщения. Например, если диалоговое окно предлагает пользователю войти в поставщик удостоверений, например учетную запись Майкрософт или Google, оно отправляет профиль пользователя в виде сообщения. В случае сбоя проверки подлинности диалоговое окно отправляет сведения об ошибке на хост-страницу, как показано в следующем примере.

```js
if (loginSuccess) {
    const userProfile = getProfile();
    const messageObject = {messageType: "signinSuccess", profile: userProfile};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
} else {
    const errorDetails = getError();
    const messageObject = {messageType: "signinFailure", error: errorDetails};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

> [!NOTE]
>
> - Переменная `loginSuccess` будет инициализирована после считывания отклика HTTP от поставщика удостоверений.
> - Реализация `getProfile` функций и `getError` не отображается. Они получают данные из параметра запроса или ответа HTTP.
> - Anonymous objects of different types are sent depending on whether the sign in was successful. Both have a `messageType` property, but one has a `profile` property and the other has an `error` property.

The handler code in the host page uses the value of the `messageType` property to branch as shown in the following example. Note that the `showUserName` function is the same as in the previous example and `showNotification` function displays the error in the host page's UI.

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
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
> Реализация `showNotification` не показана в примере кода, приведенном в этой статье. Пример возможного способа реализации этой функции в своей надстройке см. в статье [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

### <a name="cross-domain-messaging-to-the-host-runtime"></a>Обмен сообщениями между доменами в среду выполнения узла

После открытия диалогового окна диалоговое окно или родительская среда выполнения может уйти из домена надстройки. Если произойдет одно из этих действий, вызов `messageParent` завершится ошибкой, если в коде не указан домен родительской среды выполнения. Для этого добавьте параметр [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) в вызов `messageParent`. Этот объект имеет `targetOrigin` свойство, указывающее домен, в который должно быть отправлено сообщение. Если параметр не используется, Office предполагает, что целевой объект является тем же доменом, который сейчас размещается в диалоговом окне.

> [!NOTE]
> Для `messageParent` отправки междоменного сообщения требуется [набор обязательных элементов Dialog Origin 1.1](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets). Параметр `DialogMessageOptions` игнорируется в более старых версиях Office, которые не поддерживают набор требований, поэтому поведение метода не влияет на его передачу.

Ниже приведен пример использования `messageParent` для отправки междоменного сообщения.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "https://resource.contoso.com" });
```

> [!NOTE]
> Параметр `DialogMessageOptions` был выпущен примерно 19 июля 2021 г. Примерно через 30 дней после этой даты в Office в Интернете при `messageParent` первом вызове без `DialogMessageOptions` параметра и родительском домене, отличном от диалогового окна, пользователю будет предложено утвердить отправку данных в целевой домен. Если пользователь одобряет, ответ пользователя кэшируется в течение 24 часов. Пользователь не будет повторно запрашивать в течение этого периода, когда `messageParent` вызывается с тем же целевым доменом.

Если сообщение не содержит конфиденциальные данные, можно задать `targetOrigin` для параметра значение "\*", что позволяет отправлять его в любой домен. Ниже приведен пример.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "*" });
```

> [!TIP]
> Параметр `DialogMessageOptions` был добавлен в метод в качестве обязательного `messageParent` параметра в середине 2021 года. Старые надстройки, отправляющие междоменные сообщения с помощью метода, больше не работают, пока не будут обновлены для использования нового параметра. Пока надстройка не будет обновлена, *только в Office в Windows* пользователи и системные администраторы могут разрешить этим надстройкам продолжать работу, указав доверенные домены с параметром реестра **:HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Для этого создайте файл с расширением `.reg` , сохраните его на компьютере с Windows, а затем дважды щелкните его, чтобы запустить его. Ниже приведен пример содержимого такого файла.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="pass-information-to-the-dialog-box"></a>Передача данных диалоговому окну

Надстройка может отправлять сообщения с [главной страницы](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) в диалоговое окно с помощью [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)).

### <a name="use-messagechild-from-the-host-page"></a>Использование `messageChild()` с главной страницы

При вызове API диалога Office для открытия диалогового окна возвращается объект [Dialog](/javascript/api/office/office.dialog) . Он должен быть назначен переменной с большей областью, чем метод [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) , так как на объект будут ссылаться другие методы. Ниже приведен пример.

```javascript
let dialog;
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

Этот `Dialog` объект имеет метод [messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) , который отправляет в диалоговое окно любую строку, включая строкифицированные данные. В диалоговом окне возникает `DialogParentMessageReceived` событие. Код должен обрабатывать это событие, как показано в следующем разделе.

Рассмотрим сценарий, в котором пользовательский интерфейс диалогового окна связан с текущим активным листом и положением этого листа относительно других листов. В следующем примере `sheetPropertiesChanged` отправляет свойства листа Excel в диалоговое окно. В этом случае текущий лист называется "Мой лист" и является вторым листом в книге. Данные инкапсулируются в объект и строкируются, чтобы их можно было передать в `messageChild`.

```javascript
function sheetPropertiesChanged() {
    const messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Дескриптор DialogParentMessageReceived в диалоговом окне

В JavaScript диалогового окна зарегистрируйте обработчик для `DialogParentMessageReceived` события с помощью метода [UI.addHandlerAsync](/javascript/api/office/office.ui#office-office-ui-addhandlerasync-member(1)) . Обычно это выполняется с помощью [функции Office.onReady или Office.initialize](initialize-add-in.md), как показано ниже. (Более надежный пример приведен далее в этой статье.)

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Затем определите `onMessageFromParent` обработчик. Следующий код продолжает пример из предыдущего раздела. Обратите внимание, что Office передает аргумент обработчику, `message` а свойство объекта argument содержит строку с главной страницы. В этом примере сообщение перевернуто в объект, а jQuery используется для задания верхнего заголовка диалогового окна в соответствии с новым именем листа.

```javascript
function onMessageFromParent(arg) {
    const messageFromParent = JSON.parse(arg.message);
    $('h1').text(messageFromParent.name);
}
```

Рекомендуется проверить правильность регистрации обработчика. Это можно сделать, передав обратный вызов методу `addHandlerAsync` . Он выполняется по завершении попытки регистрации обработчика. Используйте обработчик для регистрации или отображения ошибки, если обработчик не был успешно зарегистрирован. Ниже приведен пример. Обратите внимание, что `reportError` это функция, не определенная здесь, которая регистрирует или отображает ошибку.

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

### <a name="conditional-messaging-from-parent-page-to-dialog-box"></a>Условное обмен сообщениями с родительской страницы в диалоговое окно

Так как можно выполнить несколько `messageChild` вызовов с главной страницы, но в диалоговом окне для `DialogParentMessageReceived` события есть только один обработчик, обработчик должен использовать условную логику для различения разных сообщений. Это можно сделать таким образом, который точно соответствует структуре условного обмена сообщениями, когда диалоговое окно отправляет сообщение на хост-страницу, как описано в разделе [Условные сообщения](#conditional-messaging).

> [!NOTE]
> В некоторых ситуациях `messageChild` API, который входит в [набор требований DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), может не поддерживаться. Некоторые альтернативные способы обмена сообщениями между родительскими окнами описаны в разделе [Альтернативные способы передачи сообщений в диалоговое окно с его главной страницы](parent-to-dialog.md).

> [!IMPORTANT]
> [Набор обязательных элементов DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets) нельзя указать в **\<Requirements\>** разделе манифеста надстройки. Необходимо проверить поддержку DialogApi 1.2 во время выполнения с помощью `isSetSupported` метода, как описано в разделе [Проверка среды выполнения для поддержки метода и набора требований](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). Поддержка требований манифеста находится в разработке.

### <a name="cross-domain-messaging-to-the-dialog-runtime"></a>Обмен сообщениями между доменами в среде выполнения диалога

После открытия диалогового окна диалоговое окно или родительская среда выполнения может уйти из домена надстройки. Если произойдет одно из этих действий, вызовы к `messageChild` завершатся ошибкой, если код не указывает домен среды выполнения диалога. Для этого добавьте параметр [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) в вызов `messageChild`. Этот объект имеет `targetOrigin` свойство, указывающее домен, в который должно быть отправлено сообщение. Если параметр не используется, Office предполагает, что целевой домен является тем же доменом, который сейчас размещает родительская среда выполнения.

> [!NOTE]
> Для `messageChild` отправки междоменного сообщения требуется [набор обязательных элементов Dialog Origin 1.1](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets). Параметр `DialogMessageOptions` игнорируется в более старых версиях Office, которые не поддерживают набор требований, поэтому поведение метода не влияет на его передачу.

Ниже приведен пример использования `messageChild` для отправки междоменного сообщения.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "https://resource.contoso.com" });
```

Если сообщение не содержит конфиденциальные данные, можно задать `targetOrigin` для параметра значение "\*", что позволяет *отправлять* его в любой домен. Ниже приведен пример.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "*" });
```

Так как среда выполнения, в которой размещено диалоговое окно, не может получить доступ к разделу **\<AppDomains\>** манифеста и тем самым определить, является ли домен, *из которого поступает сообщение* , является доверенным, необходимо использовать `DialogParentMessageReceived` обработчик, чтобы определить это. Объект, передаваемый обработчику, содержит домен, который в настоящее время размещен в родительском объекте в качестве свойства `origin` . Ниже приведен пример использования свойства .

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

Например, код может использовать [функцию Office.onReady или Office.initialize](initialize-add-in.md) для хранения массива доверенных доменов в глобальной переменной. Затем `arg.origin` свойство можно проверить с этим списком в обработчике.

> [!TIP]
> Параметр `DialogMessageOptions` был добавлен в метод в качестве обязательного `messageChild` параметра в середине 2021 года. Старые надстройки, отправляющие междоменные сообщения с помощью метода, больше не работают, пока не будут обновлены для использования нового параметра. Пока надстройка не будет обновлена, *только в Office в Windows* пользователи и системные администраторы могут разрешить этим надстройкам продолжать работу, указав доверенные домены с параметром реестра **:HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Для этого создайте файл с расширением `.reg` , сохраните его на компьютере с Windows, а затем дважды щелкните его, чтобы запустить его. Ниже приведен пример содержимого такого файла.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="close-the-dialog-box"></a>Закрытие диалогового окна

You can implement a button in the dialog box that will close it. To do this, the click event handler for the button should use `messageParent` to tell the host page that the button has been clicked. The following is an example.

```js
function closeButtonClick() {
    const messageObject = {messageType: "dialogClosed"};
    const jsonMessage = JSON.stringify(messageObject);
    Office.context.ui.messageParent(jsonMessage);
}
```

Обработчик главной страницы для `DialogMessageReceived` вызовет `dialog.close`, как показано в этом примере. (См. предыдущие примеры, в которых показано, как `dialog` инициализируется объект).

```js
function processMessage(arg) {
    const messageFromDialog = JSON.parse(arg.message);
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

### <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Использование API диалогового окна Office с одностраничными приложениями и маршрутизацией на стороне клиента

При использовании Office dialog API, SPA и маршрутизация на стороне клиента должны обрабатываться с осторожностью См. статью[Рекомендации по использованию Office dialog API в SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).

### <a name="error-and-event-handling"></a>Обработка ошибок и событий

См. статью об ошибках и событиях [Обработка ошибок и событий в Office dialog box](dialog-handle-errors-events.md).

## <a name="next-steps"></a>Дальнейшие действия

Узнайте о том, как использовать Office dialog API, в [Рекомендации по использованию Office dialog API](dialog-best-practices.md).

## <a name="samples"></a>Примеры

Во всех следующих примерах используется `displayDialogAsync`. Некоторые из них имеют серверы на основе NodeJS, а другие — серверы ASP.NET/IIS-based, но логика использования метода одинакова независимо от того, как реализована серверная часть надстройки.

**Основы:**

- [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [Содержимое для обучения и создание надстроек (несколько примеров)](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**Более сложные примеры:**

- [Надстройка Office Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Office в Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [Единый вход с использованием NodeJS для надстройки Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
- [Единый вход ASPNET надстройки Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Пример монетизации SAAS надстроек Office](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [Надстройка Outlook Microsoft Graph ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Единый вход надстроек Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Средство просмотра маркеров надстройки Outlook](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Сообщение с действиями надстройки Outlook](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [Общий доступ к надстройке Outlook в OneDrive](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [Надстройка PowerPoint Microsoft Graph ASPNET InsertChart](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Сценарий общей среды выполнения Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario)
- [Надстройка Excel ASPNET QuickBooks](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Надстройка Word JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Надстройка Word JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [OAuth клиента AngularJS для надстройки Word](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Надстройка Office Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Надстройка Office OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Код шаблонов оформления пользовательского интерфейса надстройки Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)

** См. также**

- [Среды выполнения в надстройках Office](../testing/runtimes.md)