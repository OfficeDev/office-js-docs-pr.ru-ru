---
title: Использование Office Dialog API в вашей надстройках Office
description: Основные сведения о создании диалогового окна в надстройке Office.
ms.date: 01/22/2022
ms.localizationpriority: medium
ms.openlocfilehash: 72b5fa49a4667a0333ebe04114820912db3c15a7
ms.sourcegitcommit: 4ba5f750358c139c93eb2170ff2c97322dfb50df
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 07/06/2022
ms.locfileid: "66659964"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Использование Office Dialog API в надстройках Office

Вы можете использовать [Office dialog API](/javascript/api/office/office.ui), чтобы открывать диалоговые окна в надстройке Office. Эта статья содержит инструкции по использованию dialog API в надстройке Office.

> [!NOTE]
> Сведения о том, где сейчас поддерживается API диалогов, см. в разделе наборов обязательных [элементов API диалоговых окон](/javascript/api/requirement-sets/common/dialog-api-requirement-sets). API диалоговых окон в настоящее время поддерживается для Excel, PowerPoint и Word. Поддержка Outlook включена в различные наборы обязательных&mdash;элементов почтовых ящиков. Дополнительные сведения см. в справочнике по API.

Основной сценарий для Dialog API - включить аутентификацию с помощью таких ресурсов, как Google, Facebook или Microsoft Graph. Дополнительные сведения см. в статье [Проверка подлинности с помощью Office Dialog API](auth-with-office-dialog-api.md) *после* ознакомления с текущей статьей.

Возможность открытия диалогового окна с помощью области задач, контентной надстройки или [команды надстройки](../design/add-in-commands.md) может позволить следующее:

- Отображение страниц входа, которые невозможно открыть непосредственно в области задач.
- предоставить больше места на экране (или даже весь экран) для некоторых задач в надстройке;
- разместить видео, которое будет слишком маленьким в области задач.

> [!NOTE]
> Поскольку перекрывающиеся элементы пользовательского интерфейса не приветствуются, избегайте открытия диалогового окна на панели задач, если это не требуется в сценарий. При планировании контактной зоны помните, что в области задач можно использовать вкладки. Пример области задач с вкладками см. в примере [надстройки Excel Для JavaScript SalesTracker](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker) .

На приведенном ниже изображении показан пример диалогового окна. 

![Снимок экрана: диалоговое окно с тремя параметрами входа перед Word.](../images/auth-o-dialog-open.png)

Обратите внимание, что диалоговое окно всегда открывается в центре экрана. Пользователь может перемещать ее и изменять ее размер. Окно не является *модальное*— пользователь может продолжать взаимодействовать как с документом в приложении Office, так и со страницей в области задач, если таковой имеется.

## <a name="open-a-dialog-box-from-a-host-page"></a>Откройте диалоговое окно с главной страницы

Office JavaScript API включает в себя [Диалоговый](/javascript/api/office/office.dialog) объекта и две функции в [Office.context.ui namespace](/javascript/api/office/office.ui).

Чтобы открыть диалоговое окно, ваш код, обычно страница в панели задач, вызывает метод [displayDialogAsync](/javascript/api/office/office.ui) и передает ему URL-адрес ресурса, который вам нужно открыть. Страница, на которой этот метод вызван, называется "главной страницей". Например, если вы вызываете этот метод в скрипте для index.html на панели задач, то index.html - это главная страница диалогового окна, которое открывает метод.

Ресурс, который открывается в диалоговом окне, обычно представляет собой страницу, но это может быть метод контроллера в приложении MVC, маршрут, метод веб-службы или любой другой ресурс. В этой статье "страница" или "веб-сайт" ссылается на ресурс в диалоговом окне. Ниже приведен простой пример кода.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - В случае URL-адреса используется протокол HTTP **S**, Обязательный для всех страниц, загружаемых в диалоговом окне, а не только для первой страницы.
> - Домен диалогового окна совпадает с доменом главной страницы, которая может быть страницей в панели задач или [файлом функции](/javascript/api/manifest/functionfile) команды надстройки. Страница, метод контроллера или другой ресурс, передаваемый в метод `displayDialogAsync`, должен быть в том же домене, что и страница ведущего приложения.

> [!IMPORTANT]
> Главная страница и ресурс, который открывается в диалоговом окне, должны иметь один и тот же полный домен. Если вы попробуете передать поддомен домена надстройки в `displayDialogAsync`, ничего не получится. Полные доменные имена, включая поддомены, должны совпадать.

После загрузки первой страницы (или другого ресурса) пользователь может использовать ссылки или другой пользовательский интерфейс для перехода на любой веб-сайт (или другой ресурс), использующий HTTPS. Первая страница также может сразу перенаправлять пользователя на другой сайт.

По умолчанию диалоговое окно занимает 80 % высоты и ширины экрана устройства, но вы можете установить другие соотношения путем передачи объекта конфигурации в метод, как показано в приведенном ниже примере.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Пример надстройки, в которой используется этот метод, см. [здесь](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example). Дополнительные примеры см. в `displayDialogAsync`[разделе "Примеры"](#samples).

Установите оба значения равными 100 %, чтобы надстройка открывалась во весь экран. (На самом деле, максимальное значение составляет 99,5 %, возможность перемещать окно и изменять его размер сохраняется.)

> [!NOTE]
> В окне узла можно открыть только одно диалоговое окно. Попытка открыть другое диалоговое окно вызывает ошибку. Например, если пользователь открывает диалоговое окно из области задач, он не может открыть второе диалоговое окно с другой страницы в области задач. Однако при открытии диалогового окна из команды надстройки при каждом выборе команды открывается новый (но невидимый) [HTML-файл](../design/add-in-commands.md). При этом создается новое (невидимое) окно узла, поэтому каждое такое окно может запускать собственное диалоговое окно. Дополнительные сведения см. в [разделе "Ошибки из displayDialogAsync"](dialog-handle-errors-events.md#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Использование параметра производительности в Office в Интернете

`displayInIframe` — дополнительное свойство в объекте конфигурации, которое можно передать `displayDialogAsync`. Когда этому свойству присвоено значение `true`, а надстройка запущена для документа в Office в Интернете, диалоговое окно будет открываться быстрее, потому что будет выступать как плавающий фрейм iframe. Ниже приведен пример.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

Значение по умолчанию: `false`. Его использование равнозначно пропуску всего свойства. Если надстройка не работает в Office в Интернете, `displayInIframe` она игнорируется.

> [!NOTE]
> Не следует **использовать** , `displayInIframe: true` если диалоговое окно будет перенаправлено на страницу, которую невозможно открыть в iframe. Например, страницы входа многих популярных веб-служб, таких как Учетная запись Google и Майкрософт, не могут быть открыты в iframe.

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Отправка сведений из диалогового окна главной странице

> [!NOTE]
>
> - Для ясности в этом разделе мы называем сообщение целевой страницей *узла, но* строго говоря, сообщения будут отправляться в среду выполнения *JavaScript* в области задач (или в среду выполнения, в которой размещается файл функции [).](/javascript/api/manifest/functionfile) Это различие важно только в случае обмена сообщениями между доменами. Дополнительные сведения см. в разделе [Междоменные сообщения в основной среде выполнения](#cross-domain-messaging-to-the-host-runtime).
> - Диалоговое окно не может взаимодействовать со страницей узла в области задач, если на нее не загружена библиотека API JavaScript для Office. (Как и любая страница, использующая библиотеку API JavaScript для Office, скрипт для страницы должен инициализировать надстройку. Дополнительные сведения см [. в разделе "Инициализация надстройки Office"](initialize-add-in.md).)

Код в диалоговом окне использует [функцию messageParent](/javascript/api/office/office.ui#office-office-ui-messageparent-member(1)) для отправки строкового сообщения на страницу узла. Строка может быть словом, предложением, BLOB-объектом XML, строковым JSON или любым другим объектом, который можно сериализовать в строку или привести к строке. Ниже приведен пример.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true.toString());
}
```

> [!IMPORTANT]
> - Эта `messageParent` функция является одним из *двух* API-интерфейсов Office JS, которые можно вызвать в диалоговом окне.
> - Другой API JS, который можно вызвать в диалоговом окне, — это `Office.context.requirements.isSetSupported`. Дополнительные сведения см. в [разделе "Указание приложений Office и требований К API"](specify-office-hosts-and-api-requirements.md). Однако в диалоговом окне этот API не поддерживается при Outlook 2016 однофакторной покупки (то есть версии MSI).

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
> - Office передает объект `arg` в обработчик. Его `message` свойством является строка, отправляемая вызовом `messageParent` диалогового окна. В этом примере это строковое представление профиля пользователя из службы, такой как учетная запись Майкрософт или Google, поэтому она десериализуется обратно в объект `JSON.parse`с помощью .
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

Так как из диалогового окна можно отправить несколько вызовов `messageParent`, но на главной странице есть только один обработчик для события `DialogMessageReceived`, обработчику необходимо использовать условную логику, чтобы различать сообщения. Например, если диалоговое окно предлагает пользователю войти в поставщик удостоверений, например учетную запись Майкрософт или Google, оно отправляет профиль пользователя в виде сообщения. Если проверка подлинности завершается сбоем, диалоговое окно отправляет сведения об ошибке на страницу узла, как показано в следующем примере.

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
> Реализация `showNotification` не показана в примере кода, предоставленном в этой статье. Пример возможного способа реализации этой функции в своей надстройке см. в статье [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

### <a name="cross-domain-messaging-to-the-host-runtime"></a>Междоменной обмен сообщениями в среду выполнения узла

Диалоговое окно или родительская среда выполнения JavaScript (либо в области задач, либо в среде выполнения без пользовательского интерфейса, в которой размещается файл функции) может быть перемещена из домена надстройки после открытия диалогового окна. Если что-то из этого произошло, вызов завершится ошибкой, `messageParent` если только в коде не указан домен родительской среды выполнения. Для этого добавьте параметр [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) в вызов метода `messageParent`. Этот объект имеет свойство `targetOrigin` , указывающее домен, в который должно быть отправлено сообщение. Если параметр не используется, Office предполагает, что целевой объект — это тот же домен, что и в этом диалоговом окне.

> [!NOTE]
> Для `messageParent` отправки междоменного сообщения требуется набор обязательных элементов [dialog Origin 1.1](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets). Этот `DialogMessageOptions` параметр игнорируется в более старых версиях Office, которые не поддерживают набор обязательных элементов, поэтому поведение метода не влияет на его передачу.

Ниже приведен пример использования для отправки `messageParent` междоменного сообщения.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "https://resource.contoso.com" });
```

> [!NOTE]
> Параметр `DialogMessageOptions` был выпущен приблизительно 19 июля 2021 г. Примерно через 30 дней после этой даты в Office в Интернете `messageParent` `DialogMessageOptions` при первом вызове без параметра, а родительский домен отличается от домена диалогового окна, пользователю будет предложено утвердить отправку данных в целевой домен. Если пользователь утверждает, ответ пользователя кэшируется в течение 24 часов. В течение этого периода `messageParent` при вызове с тем же целевым доменом пользователю больше не будет предложено.

Если сообщение не содержит конфиденциальные данные, можно задать значение ""\*, `targetOrigin` которое позволяет отправлять его в любой домен. Ниже приведен пример.

```js
Office.context.ui.messageParent("Some message", { targetOrigin: "*" });
```

> [!TIP]
> Этот `DialogMessageOptions` параметр был добавлен в метод `messageParent` в качестве обязательного параметра в середине 2021 года. Старые надстройки, которые отправляют междомовое сообщение с методом, больше не работают, пока не будут обновлены для использования нового параметра. Пока надстройка не будет обновлена, только в *Office для Windows* пользователи и системные администраторы могут разрешить этим надстройкам продолжать работу, указав доверенные домены с параметром реестра: **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Для этого создайте файл `.reg` с расширением, сохраните его на компьютере с Windows, а затем дважды щелкните его, чтобы запустить. Ниже приведен пример содержимого такого файла.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="pass-information-to-the-dialog-box"></a>Передача данных диалоговому окну

Надстройка может отправлять сообщения со страницы узла [](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) в диалоговое окно с помощью [Dialog.messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)).

### <a name="use-messagechild-from-the-host-page"></a>Использование `messageChild()` со страницы узла

При вызове API диалогового окна Office для открытия диалогового окна возвращается объект [Dialog](/javascript/api/office/office.dialog) . Его следует назначить переменной с большей областью действия, чем метод [displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)) , так как на объект будут ссылаться другие методы. Ниже приведен пример.

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

Этот `Dialog` объект имеет метод [messageChild](/javascript/api/office/office.dialog#office-office-dialog-messagechild-member(1)) , который отправляет в диалоговое окно любую строку, включая строковые данные. Это вызывает событие `DialogParentMessageReceived` в диалоговом окне. Код должен обработать это событие, как показано в следующем разделе.

Рассмотрим сценарий, в котором пользовательский интерфейс диалогового окна связан с текущим активным листом и положением этого листа относительно других листов. В следующем примере свойства `sheetPropertiesChanged` листа Excel отправляются в диалоговое окно. В этом случае текущий лист называется "Мой лист" и является вторым листом в книге. Данные инкапсулируются в объект и строковые, чтобы их можно было передать `messageChild`.

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

### <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Обработка DialogParentMessageReceived в диалоговом окне

В JavaScript `DialogParentMessageReceived` диалогового окна зарегистрируйте обработчик события с помощью метода [UI.addHandlerAsync](/javascript/api/office/office.ui#office-office-ui-addhandlerasync-member(1)) . Обычно это делается в методах [Office.onReady или Office.initialize](initialize-add-in.md), как показано ниже. (Ниже приведен более надежный пример.)

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Затем определите обработчик `onMessageFromParent` . Следующий код продолжает пример из предыдущего раздела. Обратите внимание, что Office передает `message` аргумент обработчику и свойство объекта аргумента содержит строку со страницы узла. В этом примере сообщение возвращается к объекту, а jQuery используется для задания верхнего заголовка диалогового окна в соответствии с новым именем листа.

```javascript
function onMessageFromParent(arg) {
    var messageFromParent = JSON.parse(arg.message);
    $('h1').text(messageFromParent.name);
}
```

Рекомендуется проверить правильность регистрации обработчика. Это можно сделать, передав методу обратный `addHandlerAsync` вызов. Это выполняется после завершения попытки регистрации обработчика. Используйте обработчик для регистрации или отображения ошибки, если обработчик не был успешно зарегистрирован. Ниже приведен пример. Обратите внимание `reportError` , что это функция, не определенная здесь, которая регистрирует или отображает ошибку.

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

Так как `messageChild` на странице узла можно выполнить несколько вызовов, `DialogParentMessageReceived` но в диалоговом окне события есть только один обработчик, обработчик должен использовать условную логику для различения разных сообщений. Это можно сделать точно так же, как и структурировать условный обмен сообщениями, когда диалоговое окно отправляет сообщение на страницу узла, как описано в разделе "Условные сообщения ["](#conditional-messaging).

> [!NOTE]
> В некоторых ситуациях `messageChild` API, который входит в набор обязательных элементов [DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets), может не поддерживаться. Некоторые альтернативные способы обмена сообщениями из родительского диалогового окна описаны в разделе "Альтернативные способы передачи сообщений в диалоговое окно с [хост-страницы"](parent-to-dialog.md).

> [!IMPORTANT]
> Набор [обязательных элементов DialogApi 1.2](/javascript/api/requirement-sets/common/dialog-api-requirement-sets) **\<Requirements\>** нельзя указать в разделе манифеста надстройки. Вам потребуется проверить поддержку DialogApi 1.2 `isSetSupported` во время выполнения с помощью метода, как описано в проверках среды выполнения на наличие поддержки методов и [наборов требований](../develop/specify-office-hosts-and-api-requirements.md#runtime-checks-for-method-and-requirement-set-support). Поддержка требований манифеста находится на этапе разработки.

### <a name="cross-domain-messaging-to-the-dialog-runtime"></a>Междоменной обмен сообщениями в среду выполнения диалоговых окон

Диалоговое окно или родительская среда выполнения JavaScript (либо в области задач, либо в среде выполнения без пользовательского интерфейса, в которой размещается файл функции) может быть перемещена из домена надстройки после открытия диалогового окна. Если что-то из этого произошло, вызов завершится ошибкой, `messageChild` если только в коде не указан домен среды выполнения диалогового окна. Для этого добавьте параметр [DialogMessageOptions](/javascript/api/office/office.dialogmessageoptions) в вызов метода `messageChild`. Этот объект имеет свойство `targetOrigin` , указывающее домен, в который должно быть отправлено сообщение. Если параметр не используется, Office предполагает, что целевой объект — это тот же домен, что и родительская среда выполнения. 

> [!NOTE]
> Для `messageChild` отправки междоменного сообщения требуется набор обязательных элементов [dialog Origin 1.1](/javascript/api/requirement-sets/common/dialog-origin-requirement-sets). Этот `DialogMessageOptions` параметр игнорируется в более старых версиях Office, которые не поддерживают набор обязательных элементов, поэтому поведение метода не влияет на его передачу.

Ниже приведен пример использования для отправки `messageChild` междоменного сообщения.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "https://resource.contoso.com" });
```

Если сообщение не содержит конфиденциальные данные, можно задать значение ""\*, `targetOrigin` которое позволяет отправлять *его в любой* домен. Ниже приведен пример.

```js
dialog.messageChild(messageToDialog, { targetOrigin: "*" });
```

Так как среда выполнения JavaScript **\<AppDomains\>**, в которой размещается диалоговое окно, не может получить доступ к разделу манифеста и тем самым определить,  является ли домен, из которого поступает сообщение, доверенным, `DialogParentMessageReceived` необходимо использовать обработчик для определения этого. Объект, передаваемый обработчику, содержит домен, который в настоящее время размещен в родительском объекте в качестве его `origin` свойства. Ниже приведен пример использования свойства.

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

Например, код может использовать методы [Office.onReady или Office.initialize](initialize-add-in.md) для хранения массива доверенных доменов в глобальной переменной. Затем `arg.origin` свойство может быть проверено для этого списка в обработчике.

> [!TIP]
> Этот `DialogMessageOptions` параметр был добавлен в метод `messageChild` в качестве обязательного параметра в середине 2021 года. Старые надстройки, которые отправляют междомовое сообщение с методом, больше не работают, пока не будут обновлены для использования нового параметра. Пока надстройка не будет обновлена, только в *Office для Windows* пользователи и системные администраторы могут разрешить этим надстройкам продолжать работу, указав доверенные домены с параметром реестра: **HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains**. Для этого создайте файл `.reg` с расширением, сохраните его на компьютере с Windows, а затем дважды щелкните его, чтобы запустить. Ниже приведен пример содержимого такого файла.
>
> ```
> Windows Registry Editor Version 5.00
> 
> [HKEY_CURRENT_USER\SOFTWARE\Microsoft\Office\16.0\WEF\AllowedDialogCommunicationDomains]
> "My trusted domain"="https://www.contoso.com"
> "Another trusted domain"="https://fabrikam.com"
> ```

## <a name="close-the-dialog-box"></a>Закрытие диалогового окна

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

### <a name="use-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Использование API диалогового окна Office с одностраничного приложениями и маршрутиза запросом на стороне клиента

При использовании Office dialog API, SPA и маршрутизация на стороне клиента должны обрабатываться с осторожностью См. статью[Рекомендации по использованию Office dialog API в SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).

### <a name="error-and-event-handling"></a>Обработка ошибок и событий

См. статью об ошибках и событиях [Обработка ошибок и событий в Office dialog box](dialog-handle-errors-events.md).

## <a name="next-steps"></a>Дальнейшие действия

Узнайте о том, как использовать Office dialog API, в [Рекомендации по использованию Office dialog API](dialog-best-practices.md).

## <a name="samples"></a>Примеры

Во всех следующих примерах используется `displayDialogAsync`. Некоторые серверы на основе NodeJS, а другие имеют ASP.NET/IIS-based серверы, но логика использования метода та же независимо от способа реализации надстройки на стороне сервера.

**Основы:**

- [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example)
- [Обучающее содержимое и сборка надстроек (несколько примеров)](https://github.com/OfficeDev/TrainingContent/tree/2db14a16774e1539a3eebae7dada4798142b8493/OfficeAddin)

**Более сложные примеры:**

- [ASPNET надстройки Microsoft Graph](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-ASPNET)
- [Надстройка Office в Microsoft Graph React](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-Microsoft-Graph-React)
- [Единый вход с использованием NodeJS для надстройки Office](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-NodeJS-SSO)
- [Единый вход надстройки Office ASPNET](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Office-Add-in-ASPNET-SSO)
- [Пример монетизации SAAS надстройки Office](https://github.com/OfficeDev/office-add-in-saas-monetization-sample)
- [AsPNET надстройки Microsoft Graph для Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-Microsoft-Graph-ASPNET)
- [Единый вход надстройки Outlook](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/auth/Outlook-Add-in-SSO)
- [Средство просмотра маркеров надстроек Outlook](https://github.com/OfficeDev/Outlook-Add-In-Token-Viewer)
- [Сообщение с действиями надстройки Outlook](https://github.com/OfficeDev/Outlook-Add-In-Actionable-Message)
- [Общий доступ к надстройке Outlook в OneDrive](https://github.com/OfficeDev/Outlook-Add-in-Sharing-to-OneDrive)
- [Вставка asPNET надстройки PowerPoint в Microsoft Graph](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart)
- [Сценарий общей среды выполнения Excel](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Samples/excel-shared-runtime-scenario)
- [Краткие книги по надстройке Excel ASPNET](https://github.com/OfficeDev/Excel-Add-in-ASPNET-QuickBooks)
- [Надстройка Word JS Redact](https://github.com/OfficeDev/Word-Add-in-JS-Redact)
- [Надстройка Word JS SpecKit](https://github.com/OfficeDev/Word-Add-in-JS-SpecKit)
- [Клиент OAuth надстройки Word AngularJS](https://github.com/OfficeDev/Word-Add-in-AngularJS-Client-OAuth)
- [Надстройка Office Auth0](https://github.com/OfficeDev/Office-Add-in-Auth0)
- [Надстройки Office OAuth.io](https://github.com/OfficeDev/Office-Add-in-OAuth.io)
- [Код шаблонов оформления надстроек Office](https://github.com/OfficeDev/Office-Add-in-UX-Design-Patterns-Code)
