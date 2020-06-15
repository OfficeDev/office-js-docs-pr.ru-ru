---
title: Использование Office Dialog API в вашей надстройках Office
description: Общие сведения о создании диалогового окна в надстройке Office.
ms.date: 06/10/2020
localization_priority: Normal
ms.openlocfilehash: 749fd6041c2ef60a4d766e865e25d53e97298d01
ms.sourcegitcommit: 449a728118db88dea22a44f83728d21604d6ee8c
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/12/2020
ms.locfileid: "44719072"
---
# <a name="use-the-office-dialog-api-in-office-add-ins"></a>Использование Office Dialog API в надстройках Office

Вы можете использовать [Office dialog API](/javascript/api/office/office.ui), чтобы открывать диалоговые окна в надстройке Office. Эта статья содержит инструкции по использованию dialog API в надстройке Office.

> [!NOTE]
> Сведения о поддержке Dialog API см. в статье [Наборы обязательных элементов Dialog API](../reference/requirement-sets/dialog-api-requirement-sets.md). В настоящее время Dialog API поддерживается для Word, Excel, PowerPoint и Outlook.

Основной сценарий для Dialog API - включить аутентификацию с помощью таких ресурсов, как Google, Facebook или Microsoft Graph. Дополнительные сведения см. в статье [Проверка подлинности с помощью Office Dialog API](auth-with-office-dialog-api.md) *после* ознакомления с текущей статьей.

Возможность открытия диалогового окна с помощью области задач, контентной надстройки или [команды надстройки](../design/add-in-commands.md) может позволить следующее:

- отобразить страницу входа, которую невозможно открыть непосредственно в области задач;
- предоставить больше места на экране (или даже весь экран) для некоторых задач в надстройке;
- разместить видео, которое будет слишком маленьким в области задач.

> [!NOTE]
> Поскольку перекрывающиеся элементы пользовательского интерфейса не приветствуются, избегайте открытия диалогового окна на панели задач, если это не требуется в сценарий. При планировании контактной зоны помните, что в области задач можно использовать вкладки. Например, как в [надстройке JavaScript SalesTracker для Excel](https://github.com/OfficeDev/Excel-Add-in-JavaScript-SalesTracker).

На приведенном ниже изображении показан пример диалогового окна. 

![Команды надстроек](../images/auth-o-dialog-open.png)

Обратите внимание, что диалоговое окно всегда открывается в центре экрана. Пользователь может перемещать ее и изменять ее размер. Окно является *не модальным*: пользователь может продолжать взаимодействовать как с документом в главном приложении Office, так и со страницей на панели задач, если она есть.

## <a name="open-a-dialog-box-from-a-host-page"></a>Откройте диалоговое окно с главной страницы

Office JavaScript API включает в себя [Диалоговый](/javascript/api/office/office.dialog) объекта и две функции в [Office.context.ui namespace](/javascript/api/office/office.ui).

Чтобы открыть диалоговое окно, ваш код, обычно страница в панели задач, вызывает метод [displayDialogAsync](/javascript/api/office/office.ui) и передает ему URL-адрес ресурса, который вам нужно открыть. Страница, на которой этот метод вызван, называется "главной страницей". Например, если вы вызываете этот метод в скрипте для index.html на панели задач, то index.html - это главная страница диалогового окна, которое открывает метод.

Ресурс, который открывается в диалоговом окне, обычно представляет собой страницу, но это может быть метод контроллера в приложении MVC, маршрут, метод веб-службы или любой другой ресурс. В этой статье "страница" или "веб-сайт" ссылается на ресурс в диалоговом окне. Ниже приведен простой пример кода.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html');
```

> [!NOTE]
> - В случае URL-адреса используется протокол HTTP**S**, Обязательный для всех страниц, загружаемых в диалоговом окне, а не только для первой страницы.
> - Домен диалогового окна совпадает с доменом главной страницы, которая может быть страницей в панели задач или [файлом функции](../reference/manifest/functionfile.md) команды надстройки. Страница, метод контроллера или другой ресурс, передаваемый в метод `displayDialogAsync`, должен быть в том же домене, что и страница ведущего приложения.

> [!IMPORTANT]
> Главная страница и ресурс, который открывается в диалоговом окне, должны иметь один и тот же полный домен. Если вы попробуете передать поддомен домена надстройки в `displayDialogAsync`, ничего не получится. Полные доменные имена, включая поддомены, должны совпадать.

После загрузки первой страницы (или другого ресурса) пользователь может использовать ссылки или другой пользовательский интерфейс для перехода на любой веб-сайт (или другой ресурс), использующий HTTPS. Первая страница также может сразу перенаправлять пользователя на другой сайт.

По умолчанию диалоговое окно занимает 80 % высоты и ширины экрана устройства, но вы можете установить другие соотношения путем передачи объекта конфигурации в метод, как показано в приведенном ниже примере.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20});
```

Подобная надстройка приведена в статье [Пример надстройки Office с Dialog API](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

Установите оба значения равными 100 %, чтобы надстройка открывалась во весь экран. (На самом деле, максимальное значение составляет 99,5 %, возможность перемещать окно и изменять его размер сохраняется.)

> [!NOTE]
> Из главного окна можно открыть только одно диалоговое окно. При попытке открыть еще одно диалоговое окно произойдет ошибка. Поэтому если пользователь, например, откроет диалоговое окно из области задач, он не сможет открыть второе диалоговое окно на другой странице в области задач. Но при открытии диалогового окна с помощью [команды надстройки](../design/add-in-commands.md) каждый раз открывается новый (невидимый) HTML-файл. При этом создается новое (невидимое) главное окно, которое может запускать собственное диалоговое окно. Дополнительные сведения см. в разделе [Ошибки метода displayDialogAsync](dialog-handle-errors-events.md#errors-from-displaydialogasync).

### <a name="take-advantage-of-a-performance-option-in-office-on-the-web"></a>Использование параметра производительности в Office в Интернете

`displayInIframe` — дополнительное свойство в объекте конфигурации, которое можно передать `displayDialogAsync`. Когда этому свойству присвоено значение `true`, а надстройка запущена для документа в Office в Интернете, диалоговое окно будет открываться быстрее, потому что будет выступать как плавающий фрейм iframe. Ниже приведен пример.

```js
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html', {height: 30, width: 20, displayInIframe: true});
```

Значение по умолчанию: `false`. Его использование равнозначно пропуску всего свойства. Если надстройка не работает в Office в Интернете, `displayInIframe` игнорируется.

> [!NOTE]
> Вам **не** следует `displayInIframe: true`использовать, если диалоговое окно будет выполнять перенаправление на страницу, которую невозможно открыть в элементе iframe. Например, страницы входа многих популярных веб-служб, таких как Google или учетной записи Майкрософт, невозможно открыть в iframe.

## <a name="send-information-from-the-dialog-box-to-the-host-page"></a>Отправка сведений из диалогового окна главной странице

Диалоговое окно может взаимодействовать с главной страницей в области задач, если:

- Текущая страница в диалоговом окне не находится в том же домене, что и главная страница.
- На странице загружается библиотека API JavaScript для Office. (Как и любая страница, использующая библиотеку API JavaScript для Office, сценарий для страницы должен назначить метод `Office.initialize` свойству, хотя это может быть пустой метод. Дополнительные сведения см. [в статье Initialize Your надстройка Office](initialize-add-in.md).

Код в диалоговом окне использует функцию [messageParent](/javascript/api/office/office.ui#messageparent-message-), чтобы отправить на главную страницу логическое значение или строковое сообщение. Строка может быть словом, предложением, большим двоичным объектом XML, преобразованными данными JSON или любыми другими объектами, которые можно сериализовать в строку. Ниже приведен пример.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(true);
}
```

> [!IMPORTANT]
> - Функцию `messageParent` можно вызывать только на странице, которая относится к тому же домену (включая протокол и порт), что и главная страница.
> - `messageParent`Функция является одним из двух *only* API Office JS, которые можно вызывать в диалоговом окне. 
> - Другой API JS, который может вызываться в диалоговом окне, — это `Office.context.requirements.isSetSupported` . Сведения о нем: [Указание ведущих приложений Office и требований к API](specify-office-hosts-and-api-requirements.md). Однако в диалоговом окне этот API не поддерживается в Outlook 2016 1-Time Purchase (версия MSI).

В следующем примере `googleProfile` — это строковое представление профиля Google пользователя.

```js
if (loginSuccess) {
    Office.context.ui.messageParent(googleProfile);
}
```

Чтобы главная страница получила сообщение, ее необходимо настроить. Для этого добавьте параметр обратного вызова в исходный вызов `displayDialogAsync`. Обратный вызов назначает событию `DialogMessageReceived` обработчик. Ниже приведен пример.

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
> - Office передает объект `arg` в обработчик. Его `message` свойством является логическое значение или строка, отправляемая при вызове `messageParent` в диалоговом окне. В данном примере это профиль пользователя из учетной записи Майкрософт, Google или другой службы, представленный в виде строки, поэтому он десериализируется обратно в объект с помощью метода `JSON.parse`.
> - Функция `showUserName` не показана. Она может отображать персонализированное приветствие в области задач.

Когда взаимодействие пользователя с диалоговым окном закончится, обработчик сообщений должен закрыть диалоговое окно, как показано в этом примере.

```js
function processMessage(arg) {
    dialog.close();
    // message processing code goes here;
}
```

> [!NOTE]
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

Так как из диалогового окна можно отправить несколько вызовов `messageParent`, но на главной странице есть только один обработчик для события `DialogMessageReceived`, обработчику необходимо использовать условную логику, чтобы различать сообщения. Например, если диалоговое окно предлагает пользователю войти в учетную запись Майкрософт, Google или другого поставщика удостоверений, оно отправляет профиль пользователя в виде сообщения. Если выполнить аутентификацию не удается, диалоговое окно отправляет сведения об ошибке на главную страницу, как показано в приведенном ниже примере.

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
> Реализация функции `showNotification` не показана в примере кода, представленном в этой статье. Пример возможного способа реализации этой функции в своей надстройке см. в статье [Пример использования API диалоговых окон в надстройке Office](https://github.com/OfficeDev/Office-Add-in-Dialog-API-Simple-Example).

## <a name="pass-information-to-the-dialog-box"></a>Передача данных диалоговому окну

Иногда главной странице нужно передать данные в диалоговое окно. Есть два основных способа обеспечить эту возможность:

- Добавьте параметры запроса в URL-адрес, который передается в метод `displayDialogAsync`.
- Храните информацию в месте, доступном как для главного, так и для диалогового окна. Два окна не разделяют общее хранилище сеансов, но *если они имеют один и тот же домен* (включая номер порта, если таковой имеется), они совместно используют общее [Локальное хранилище](https://www.w3schools.com/html/html5_webstorage.asp).\*

> [!NOTE]
> \* Существует ошибка, влияющая на вашу стратегию обработки маркеров. Если надстройка работает в **Office в Интернете** с использованием браузера Safari или Microsoft Edge, у диалогового окна и области задач нет одного общего локального хранилища, поэтому его нельзя использовать для связи между ними.

### <a name="use-local-storage"></a>Использование локального хранилища

Чтобы использовать локальное хранилище, код вызывает метод `setItem` объекта `window.localStorage` на главной странице перед вызовом `displayDialogAsync`, как показано в приведенном ниже примере.

```js
localStorage.setItem("clientID", "15963ac5-314f-4d9b-b5a1-ccb2f1aea248");
```

Код в диалоговом окне считывает элемент, когда он необходим, как в следующем примере:

```js
var clientID = localStorage.getItem("clientID");
// You can also use property syntax:
// var clientID = localStorage.clientID;
```

### <a name="use-query-parameters"></a>Использование параметров запроса

В приведенном ниже примере показано, как передавать данные с помощью параметра запроса.

```js
Office.context.ui.displayDialogAsync('https://myAddinDomain/myDialog.html?clientID=15963ac5-314f-4d9b-b5a1-ccb2f1aea248');
```

Пример, в котором используется эта техника, см. в статье [Вставка диаграмм Excel с помощью Microsoft Graph в надстройке PowerPoint](https://github.com/OfficeDev/PowerPoint-Add-in-Microsoft-Graph-ASPNET-InsertChart).

Код в вашем диалоговом окне может проанализировать URL-адрес и прочитать значение параметра.

> [!IMPORTANT]
> Office автоматически добавляет параметр запроса `_host_info` в URL-адрес, который передается `displayDialogAsync`. (Этот параметр добавляется после пользовательских параметров запроса, если они есть. Он не добавляется в последующие URL-адреса, которые открываются в диалоговом окне.) Корпорация Майкрософт может изменить содержимое этого значения или удалить его полностью, поэтому ваш код не должен его считывать. То же значение добавляется в хранилище сеанса диалогового окна. *Ваш код не должен ни считывать это значение, ни записывать в него данные*.

> [!NOTE]
> Теперь вы `messageChild` можете просмотреть API, который родительская страница может использовать для отправки сообщений в диалоговое окно, так как `messageParent` описанный выше API отправляет сообщения из диалогового окна. Дополнительные сведения см. в статье [Передача данных и сообщений в диалоговое окно с главной страницы](parent-to-dialog.md). Мы рекомендуем испытать ее, но для рабочих надстроек мы рекомендуем использовать методы, описанные в этом разделе.

## <a name="closing-the-dialog-box"></a>Закрытие диалогового окна

Вы можете добавить в диалоговое окно кнопку, которая будет его закрывать. Для этого обработчик событий кнопки должен использовать `messageParent`, чтобы сообщить главной странице, что кнопка нажата. Ниже приведен пример.

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

### <a name="using-the-office-dialog-api-with-single-page-applications-and-client-side-routing"></a>Использование Office dialog API с одностраничными приложениями и клиентской маршрутизацией

При использовании Office dialog API, SPA и маршрутизация на стороне клиента должны обрабатываться с осторожностью См. статью[Рекомендации по использованию Office dialog API в SPA](dialog-best-practices.md#best-practices-for-using-the-office-dialog-api-in-an-spa).

### <a name="error-and-event-handling"></a>Обработка ошибок и событий

См. статью об ошибках и событиях [Обработка ошибок и событий в Office dialog box](dialog-handle-errors-events.md).

## <a name="next-steps"></a>Дальнейшие действия

Узнайте о том, как использовать Office dialog API, в [Рекомендации по использованию Office dialog API](dialog-best-practices.md).
