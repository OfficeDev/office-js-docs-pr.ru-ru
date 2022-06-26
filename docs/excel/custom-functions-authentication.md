---
ms.date: 06/15/2022
description: Проверка подлинности пользователей с помощью пользовательских функций, которые не используют общую среду выполнения.
title: Проверка подлинности для пользовательских функций без общей среды выполнения
ms.localizationpriority: medium
ms.openlocfilehash: 0f4493f9cf68236a9d9d83ebd3299c9ce3371560
ms.sourcegitcommit: d8fbe472b35c758753e5d2e4b905a5973e4f7b52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 06/25/2022
ms.locfileid: "66229682"
---
# <a name="authentication-for-custom-functions-without-a-shared-runtime"></a>Проверка подлинности для пользовательских функций без общей среды выполнения

В некоторых сценариях пользовательская функция, которая не использует общую среду выполнения, потребуется выполнить проверку подлинности пользователя, чтобы получить доступ к защищенным ресурсам. Пользовательские функции, которые не используют общую среду выполнения, выполняются в среде выполнения только для JavaScript. По этой причине, если надстройка имеет область задач, необходимо передавать данные между средой выполнения, доступной только для JavaScript, и средой выполнения, поддерживающей HTML, используемой областью задач. Это можно сделать с помощью [объекта OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) и специального API диалогового окна.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>Объект OfficeRuntime.storage

Среда выполнения только для JavaScript `localStorage` не имеет объекта, доступного в глобальном окне, где обычно хранятся данные. Вместо этого код должен совместно использовать данные между `OfficeRuntime.storage` пользовательскими функциями и области задач, используя для задания и получения данных.

### <a name="suggested-usage"></a>Рекомендуемое использование

Если необходимо выполнить проверку подлинности из пользовательской надстройки функции, которая не использует общую среду выполнения, `OfficeRuntime.storage` код должен проверить, был ли уже получен маркер доступа. В противном случае используйте [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) для проверки подлинности пользователя, получения маркера доступа и `OfficeRuntime.storage` последующего хранения маркера.

## <a name="dialog-api"></a>API диалогового окна

Если маркер не существует, `OfficeRuntime.dialog` следует использовать API, чтобы попросить пользователя выполнить вход. После того как пользователь введет свои учетные данные, полученный маркер доступа можно сохранить как элемент в `OfficeRuntime.storage`.

> [!NOTE]
> Среда выполнения только для JavaScript использует объект диалога, который немного отличается от объекта диалогового окна в среде выполнения ядра браузера, используемой в области задач. Они оба называются Dialog API, но используют [OfficeRuntime.displayWebDialog](/javascript/api/office-runtime#office-runtime-officeruntime-displaywebdialog-function(1)) для проверки подлинности пользователей в среде выполнения только javaScript *,* [а не Office.ui.displayDialogAsync](/javascript/api/office/office.ui#office-office-ui-displaydialogasync-member(1)).

На следующей схеме показан этот основной процесс. Пунктирная линия указывает, что пользовательские функции и область задач надстройки являются частью надстройки в целом, хотя они используют отдельные среды выполнения.

1. Вы выполняете вызов пользовательской функции из ячейки книги Excel.
2. Пользовательская функция использует `OfficeRuntime.dialog` для передачи учетных данных пользователя на веб-сайт.
3. Этот веб-сайт возвращает маркер доступа для пользовательской функции.
4. Затем пользовательская функция задает этот маркер доступа для элемента в .`OfficeRuntime.storage`
5. Область задач надстройки получает доступ к маркеру из объекта `OfficeRuntime.storage`.

![Схема пользовательской функции с помощью API диалогового окна для получения маркера доступа, а затем предоставление общего доступа маркеру в области задач через API OfficeRuntime.storage.](../images/authentication-diagram.png "Схема проверки подлинности.")

## <a name="storing-the-token"></a>Хранение маркера

Следующие примеры взяты из примера кода [Использование OfficeRuntime.storage в пользовательских функциях](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AsyncStorage). Полный пример совместного использования данных между пользовательскими функциями и областью задач в надстройки, не использующие общую среду выполнения, см. в этом примере кода.

При проверке подлинности пользовательской функцией она получает маркер доступа, который должен храниться в объекте `OfficeRuntime.storage`. В следующем примере кода показано, как вызвать метод `storage.setItem` чтобы сохранить значение. Эта `storeValue` функция представляет собой пользовательскую функцию, которая хранит значение от пользователя. Можно внести изменение, чтобы сохранять любые нужные значения маркеров.

```js
/**
 * Stores a key-value pair into OfficeRuntime.storage.
 * @customfunction
 * @param {string} key Key of item to put into storage.
 * @param {*} value Value of item to put into storage.
 */
function storeValue(key, value) {
  return OfficeRuntime.storage.setItem(key, value).then(function (result) {
      return "Success: Item with key '" + key + "' saved to storage.";
  }, function (error) {
      return "Error: Unable to save item with key '" + key + "' to storage. " + error;
  });
}
```

Если области задач требуется маркер доступа, она может получить маркер из `OfficeRuntime.storage` элемента. В следующем примере кода показано, как использовать метод `storage.getItem` чтобы извлечь маркер.

```js
/**
 * Read a token from storage.
 * @customfunction GETTOKEN
 */
function receiveTokenFromCustomFunction() {
  var key = "token";
  var tokenSendStatus = document.getElementById('tokenSendStatus');
  OfficeRuntime.storage.getItem(key).then(function (result) {
     tokenSendStatus.value = "Success: Item with key '" + key + "' read from storage.";
     document.getElementById('tokenTextBox2').value = result;
  }, function (error) {
     tokenSendStatus.value = "Error: Unable to read item with key '" + key + "' from storage. " + error;
  });
}
```

## <a name="general-guidance"></a>Общие рекомендации

Надстройки Office являются веб-надстройками, и вы можете использовать любой способ веб-проверки подлинности. При реализации своей собственной проверки подлинности с использованием пользовательских функций отсутствует определенный шаблон или метод. Рекомендуется ознакомиться с документацией по различным шаблонам проверки подлинности, начиная с [этой статьи об авторизации через внешние службы](../develop/auth-external-add-ins.md).  

Избегайте использования следующих расположений для хранения данных при разработке пользовательских функций.

- `localStorage`: пользовательские функции `window` , которые не используют общую среду выполнения, не имеют доступа к глобальному объекту и, следовательно, не имеют доступа к данным, хранящимся в `localStorage`.
- `Office.context.document.settings`: это расположение не является безопасным, и информацию может извлечь любой пользователь, использующий надстройку.

## <a name="dialog-box-api-example"></a>Пример API диалогового окна

В следующем примере кода функция использует `getTokenViaDialog` функцию `OfficeRuntime.displayWebDialog` для отображения диалогового окна. Этот пример предоставляется для демонстрации возможностей метода, а не для демонстрации способа проверки подлинности.

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this isn't a sufficient example of authentication but is intended to show the capabilities of the displayWebDialog method.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once. Wait for previous dialog box's token.
      let timeout = 5;
      let count = 0;
      var intervalId = setInterval(function () {
        count++;
        if(_cachedToken) {
          resolve(_cachedToken);
          clearInterval(intervalId);
        }
        if(count >= timeout) {
          reject("Timeout while waiting for token");
          clearInterval(intervalId);
        }
      }, 1000);
    } else {
      _dialogOpen = true;
      OfficeRuntime.displayWebDialog(url, {
        height: '50%',
        width: '50%',
        onMessage: function (message, dialog) {
          _cachedToken = message;
          resolve(message);
          dialog.close();
          return;
        },
        onRuntimeError: function(error, dialog) {
          reject(error);
        },
      }).catch(function (e) {
        reject(e);
      });
    }
  });
}
```

## <a name="next-steps"></a>Дальнейшие действия

Узнайте, как [выполнять отладку пользовательских функций](custom-functions-debugging.md).

## <a name="see-also"></a>См. также

* [Среда выполнения только для JavaScript для пользовательских функций](custom-functions-runtime.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)