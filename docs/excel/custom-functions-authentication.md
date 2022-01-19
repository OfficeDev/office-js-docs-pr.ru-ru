---
ms.date: 05/17/2020
description: Проверка подлинности пользователей с помощью настраиваемой Excel, которые не используют области задач.
title: Проверка подлинности для пользовательских функций без пользовательского интерфейса
ms.localizationpriority: medium
ms.openlocfilehash: 57a003dbcf3c36842c2b5c98aba7844c9e53e012
ms.sourcegitcommit: 45f7482d5adcb779a9672669360ca4d8d5c85207
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/19/2022
ms.locfileid: "62074240"
---
# <a name="authentication-for-ui-less-custom-functions"></a>Проверка подлинности для пользовательских функций без пользовательского интерфейса

В некоторых сценариях вашей настраиваемой функции, которая не использует области задач или других элементов пользовательского интерфейса (настраиваемая функция без пользовательского интерфейса), потребуется проверить подлинность пользователя, чтобы получить доступ к защищенным ресурсам. Следует помнить, что пользовательские функции без пользовательского интерфейса выполняются только для JavaScript. Из-за этого необходимо передавать данные между временем запуска только JavaScript и типичным временем запуска браузера, используемым большинством надстройок с помощью объекта и `OfficeRuntime.storage` API диалогов.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>Объект OfficeRuntime.storage

Время запуска только на JavaScript, используемое пользовательскими функциями без пользовательского интерфейса, не имеет объекта, доступного в глобальном окне, где обычно `localStorage` хранятся данные. Вместо этого следует обмениваться данными между пользовательскими функциями и области задач без пользовательского интерфейса, используя [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) для настройки и получения данных.

### <a name="suggested-usage"></a>Рекомендуемое использование

При необходимости проверки подлинности из настраиваемой функции, не входя в пользовательский интерфейс, проверьте, был ли уже приобретен маркер `storage` доступа. Если нет, используйте API диалоговых окон, чтобы проверить подлинность пользователя, извлечь маркер доступа и сохранить его в объекте `storage` для дальнейшего использования.

## <a name="dialog-api"></a>API диалоговых окон

Если маркер не существует, следует использовать API диалоговых окон, чтобы попросить пользователя войти в систему. После ввода пользователем своих учетных данных итоговый маркер доступа можно сохранить в объекте `storage`.

> [!NOTE]
> Время запуска только на JavaScript использует объект Dialog, который немного отличается от объекта Dialog в времени запуска двигателя браузера, используемого в области задач. Они оба называются "API диалогов", но используются для проверки подлинности пользователей в время запуска только `OfficeRuntime.Dialog` javaScript.

На следующей схеме показан этот основной процесс. Пунктирная строка указывает на то, что пользовательские функции без пользовательского интерфейса и области задач надстройки являются частью надстройки в целом, хотя они используют отдельные время запуска.

1. Вы выдает пользовательский вызов функции без пользовательского интерфейса из ячейки в Excel книге.
2. Настраиваемая функция, не использующая пользовательский интерфейс, используется для того, чтобы передать учетные данные пользователя `Dialog` веб-сайту.
3. Затем этот веб-сайт возвращает маркер доступа к настраиваемой функции без пользовательского интерфейса.
4. Настраиваемая функция без пользовательского интерфейса задает этот маркер доступа к `storage` .
5. Область задач надстройки получает доступ к маркеру из объекта `storage`.

![Схема настраиваемой функции с помощью диалогового API для получения маркера доступа, а затем обмена маркером с области задач через API OfficeRuntime.storage.](../images/authentication-diagram.png "Схема проверки подлинности.")

## <a name="storing-the-token"></a>Хранение маркера

Следующие примеры взяты из примера кода [Использование OfficeRuntime.storage в пользовательских функциях](https://github.com/OfficeDev/Office-Add-in-samples/tree/main/Excel-custom-functions/AsyncStorage). Обратитесь к этому примеру кода для полного примера обмена данными между пользовательскими функциями без пользовательского интерфейса и области задач.

Если настраиваемая функция без пользовательского интерфейса подает проверку подлинности, она получает маркер доступа и должна будет хранить его `storage` в . В следующем примере кода показано, как вызвать метод `storage.setItem` чтобы сохранить значение. Функция — это настраиваемая функция без пользовательского интерфейса, которая, например, `storeValue` сохраняет значение от пользователя. Можно внести изменение, чтобы сохранять любые нужные значения маркеров.

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

Когда области задач требуется маркер доступа, она может извлечь его из объекта `storage`. В следующем примере кода показано, как использовать метод `storage.getItem` чтобы извлечь маркер.

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

Надстройки Office являются веб-надстройками, и вы можете использовать любой способ веб-проверки подлинности. Для реализации собственной проверки подлинности с помощью пользовательских функций не существует определенного шаблона или метода, который необходимо выполнить. Рекомендуется ознакомиться с документацией по различным шаблонам проверки подлинности, начиная с [этой статьи об авторизации через внешние службы](../develop/auth-external-add-ins.md).  

Избегайте использования следующих местоположений для хранения данных при разработке настраиваемой функции: .

- `localStorage`. Пользовательские функции, не влияемые на пользовательский интерфейс, не имеют доступа к глобальному объекту и поэтому не имеют доступа к данным, `window` хранимым в `localStorage` .
- `Office.context.document.settings`: это расположение не защищено, и сведения могут быть извлечены любым пользователем с помощью надстройки.

## <a name="dialog-box-api-example"></a>Пример API диалоговых полей

В следующем примере кода функция использует функцию API для отображения `getTokenViaDialog` `Dialog` `displayWebDialogOptions` диалогового окна. Этот пример предоставляется для демонстрации возможностей объекта, а не для проверки `Dialog` подлинности.

```JavaScript
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
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
Узнайте, как [отламыть пользовательские функции без пользовательского интерфейса.](custom-functions-debugging.md)

## <a name="see-also"></a>См. также

* [Время запуска для пользовательских Excel пользовательских функций](custom-functions-runtime.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)