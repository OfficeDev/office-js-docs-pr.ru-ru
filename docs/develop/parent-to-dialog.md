---
title: Передача данных и сообщений в диалоговое окно с главной страницы
description: Узнайте, как передавать данные в диалоговое окно с главной страницы с помощью API Мессажечилд и Диалогпарентмессажерецеивед.
ms.date: 03/11/2020
localization_priority: Normal
ms.openlocfilehash: 03d89a2e5ffb9060edb25dd8e0c3c71c0dd274eb
ms.sourcegitcommit: 153576b1efd0234c6252433e22db213238573534
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/07/2020
ms.locfileid: "42561875"
---
# <a name="passing-data-and-messages-to-a-dialog-box-from-its-host-page-preview"></a>Передача данных и сообщений в диалоговое окно с главной страницы (Предварительная версия)

Надстройка может отправлять сообщения с [главной страницы](dialog-api-in-office-add-ins.md#open-a-dialog-box-from-a-host-page) в диалоговое окно с помощью метода [мессажечилд](/javascript/api/office/office.dialog#messagechild-message-) объекта [DIALOG](/javascript/api/office/office.dialog) .

> [!Important]
>
> - API, описанные в этой статье, доступны в предварительной версии. Они доступны разработчикам для экспериментов; но его не следует использовать в рабочей надстройке. Пока этот API не будет выпущен, используйте методы, описанные в статье [Передача сведений в диалоговое окно](dialog-api-in-office-add-ins.md#pass-information-to-the-dialog-box) для рабочих надстроек.
> - Для интерфейсов API, описанных в этой статье, требуется Office 365 (версия подписки Office). Следует использовать последнюю версию для текущего месяца и сборку из канала для участников программы предварительной оценки. Чтобы получить эту версию, необходимо быть участником программы предварительной оценки Office. Дополнительные сведения см. на странице [Примите участие в программе предварительной оценки Office](https://products.office.com/office-insider?tab=tab-1). Обратите внимание на то, что при построении градуатес к производственному каналу поддержка предварительных функций для этой сборки отключена.
> - На начальном этапе предварительной версии API поддерживаются в Excel, PowerPoint и Word; но не в Outlook.
>
> [!INCLUDE [Information about using preview APIs](../includes/using-preview-apis.md)]

## <a name="use-messagechild-from-the-host-page"></a>Использование `messageChild()` с главной страницы

Когда вы вызываете API диалоговых окон Office для открытия диалогового окна, возвращается объект [DIALOG](/javascript/api/office/office.dialog) . Она должна быть назначена переменной, которая, как правило, имеет больший объем, чем метод [displayDialogAsync](/javascript/api/office/office.ui#displaydialogasync-startaddress--callback-) , так как на объект будут ссылаться другие методы. Ниже приведен пример.

```javascript
var dialog;
Office.context.ui.displayDialogAsync('https://myDomain/myDialog.html',
    function (asyncResult) {
        dialog = asyncResult.value;
        dialog.addEventHandler(Office.EventType.DialogMessageReceived, processMessage);
    }
);

function processMessage(arg) {
    dialog.close();

  // message processing code goes here;

}
```

Этот `Dialog` объект содержит метод [мессажечилд](/javascript/api/office/office.dialog#messagechild-message-) , который отправляет любую строку или данные преобразованного в диалоговое окно. Это вызывает `DialogParentMessageReceived` событие в диалоговом окне. Код должен обрабатывать это событие, как показано в следующем разделе.

Рассмотрим сценарий, в котором пользовательский интерфейс диалогового окна должен сопоставляться с текущим активным листом и положением листа относительно других листов. В следующем примере в диалоговое окно `sheetPropertiesChanged` отправляются свойства листа Excel. В этом случае текущий лист называется "Мой лист" и является 2-м листом книги. Данные инкапсулируются в объекте, который является преобразованного, чтобы его можно было передать `messageChild`.

```javascript
function sheetPropertiesChanged() {
    var messageToDialog = JSON.stringify({
                               name: "My Sheet",
                               position: 2
                           });

    dialog.messageChild(messageToDialog);
}
```

## <a name="handle-dialogparentmessagereceived-in-the-dialog-box"></a>Обработка Диалогпарентмессажерецеивед в диалоговом окне

В JavaScript диалогового окна Зарегистрируйте обработчик для `DialogParentMessageReceived` события с помощью метода [UI. addHandlerAsync](/javascript/api/office/office.ui#addhandlerasync-eventtype--handler--options--callback-) . Как правило, это выполняется в [методах Office. onread или Office. Initialize](initialize-add-in.md). Ниже приведен пример.

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent);
    });
```

Затем определите `onMessageFromParent` обработчик. Приведенный ниже код продолжает пример из предыдущего раздела. Обратите внимание, что Office передает аргумент обработчику и что `message` свойство объекта Argument содержит строку со страницы узла. В этом примере сообщение переводится в объект, а jQuery используется для установки верхнего заголовка диалогового окна в соответствующее имя нового листа.

```javascript
function onMessageFromParent(event) {
    var messageFromParent = JSON.parse(event.message);
    $('h1').text(messageFromParent.name);
}
```

Рекомендуется проверить правильность регистрации обработчика. Для этого можно передать обратный вызов `addHandlerAsync` методу, который выполняется при завершении попытки регистрации обработчика. Используйте обработчик для записи или отображения ошибки, если обработчик не был успешно зарегистрирован. Ниже приведен пример. Обратите `reportError` внимание, что это функция, не определенная здесь, записывает или отображает сообщение об ошибке.

```javascript
Office.onReady()
    .then(function() {
        Office.context.ui.addHandlerAsync(
            Office.EventType.DialogParentMessageReceived,
            onMessageFromParent,
            onRegisterMessageComplete);
    });

function onRegisterMessageComplete(asyncResult) {
    if (asyncResult.status !== Office.AsyncResultStatus.Succeeded) {
        reportError(asyncResult.error.message);
    }
}
```

## <a name="conditional-messaging"></a>Условные сообщения

Так как вы можете выполнять `messageChild` несколько вызовов со страницы узла, но у вас есть только один обработчик в диалоговом окне для `DialogParentMessageReceived` события, обработчик должен использовать условную логику для различения разных сообщений. Это можно сделать точно так же, как при структурировании условной передачи сообщений, когда диалоговое окно отправляет сообщение на страницу узла, как описано в [условной системе обмена сообщениями](dialog-api-in-office-add-ins.md#conditional-messaging).
