---
ms.date: 05/17/2020
description: Проверка подлинности пользователей с помощью пользовательских функций в Excel, которые не используют области задач.
title: Проверка подлинности для пользовательских функций без пользовательского интерфейса
localization_priority: Normal
ms.openlocfilehash: bca3cd422330b6499e18c31ef8d7da6def81b546
ms.sourcegitcommit: d28392721958555d6edea48cea000470bd27fcf7
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 01/13/2021
ms.locfileid: "49839861"
---
# <a name="authentication-for-ui-less-custom-functions"></a>Проверка подлинности для пользовательских функций без пользовательского интерфейса

В некоторых случаях настраиваемая функция, которая не использует области задач или другие элементы пользовательского интерфейса (пользовательская функция без пользовательского интерфейса), должна будет проверить подлинность пользователя, чтобы получить доступ к защищенным ресурсам. Следует помнить, что пользовательские функции без пользовательского интерфейса выполняются в среде только JavaScript. В связи с этим вам потребуется передавать данные между обычной временем работы javaScript и типичной среде запуска браузера, используемой большинством надстройки с помощью объекта и `OfficeRuntime.storage` Dialog API.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>Объект OfficeRuntime.storage

В среде запуска только javaScript, используемой пользовательскими функциями без пользовательского интерфейса, нет объекта, доступного в глобальном окне, где обычно `localStorage` хранятся данные. Вместо этого следует обмениваться данными между пользовательскими функциями и области задач, используя [OfficeRuntime.storage](/javascript/api/office-runtime/officeruntime.storage) для настройки и получения данных.

### <a name="suggested-usage"></a>Рекомендуемое использование

Если вам нужно проверить подлинность из пользовательской функции без пользовательского интерфейса, проверьте, был ли уже получен маркер `storage` доступа. Если нет, используйте API диалоговых окон, чтобы проверить подлинность пользователя, извлечь маркер доступа и сохранить его в объекте `storage` для дальнейшего использования.

## <a name="dialog-api"></a>API диалоговых окон

Если маркер не существует, следует использовать API диалоговых окон, чтобы попросить пользователя войти в систему. После ввода пользователем своих учетных данных итоговый маркер доступа можно сохранить в объекте `storage`.

> [!NOTE]
> В среде запуска только javaScript используется объект Dialog, который немного отличается от объекта Dialog в среде запуска браузера, используемой в области задач. Они оба называются Dialog API, но используются для проверки подлинности пользователей в среде только `OfficeRuntime.Dialog` JavaScript.

На следующей схеме показан этот основной процесс. Пунктирная строка указывает, что пользовательские функции без пользовательского интерфейса и области задач надстройки являются частью надстройки в целом, хотя они используют отдельные точки работы.

1. Вызов пользовательской функции без пользовательского интерфейса из ячейки в книге Excel.
2. Пользовательская функция, не использующая пользовательский интерфейс, передает учетные данные пользователя `Dialog` на веб-сайт.
3. Затем этот веб-сайт возвращает маркер доступа к пользовательской функции без пользовательского интерфейса.
4. Ваша пользовательская функция без пользовательского интерфейса затем устанавливает этот маркер доступа для `storage` .
5. Область задач надстройки получает доступ к маркеру из объекта `storage`.

![Схема пользовательской функции с помощью dialog API для получения маркера доступа, а затем поделиться маркером с области задач через API OfficeRuntime.storage.](../images/authentication-diagram.png "Схема проверки подлинности.")

## <a name="storing-the-token"></a>Хранение маркера

Следующие примеры взяты из примера кода [Использование OfficeRuntime.storage в пользовательских функциях](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage). Полный пример общего доступа к данным между пользовательскими функциями без пользовательского интерфейса и области задач можно найти в этом примере кода.

Если пользовательская функция без пользовательского интерфейса пройдет проверку подлинности, она получит маркер доступа и потребуется сохранить его `storage` в . В следующем примере кода показано, как вызвать метод `storage.setItem` чтобы сохранить значение. Эта функция — это пользовательская функция без пользовательского интерфейса, которая, например, сохраняет `storeValue` значение от пользователя. Можно внести изменение, чтобы сохранять любые нужные значения маркеров.

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

Надстройки Office являются веб-надстройками, и вы можете использовать любой способ веб-проверки подлинности. Для реализации собственной проверки подлинности с помощью пользовательских функций не существует определенного шаблона или метода. Рекомендуется ознакомиться с документацией по различным шаблонам проверки подлинности, начиная с [этой статьи об авторизации через внешние службы](../develop/auth-external-add-ins.md).  

Избегайте использования следующих расположений для хранения данных при разработке пользовательских функций.  

- `localStorage`: пользовательские функции без пользовательского интерфейса не имеют доступа к глобальному объекту и поэтому не имеют доступа к `window` данным, хранимым в `localStorage` .
- `Office.context.document.settings`: это расположение не защищено, и сведения могут быть извлечены любым пользователем с помощью надстройки.

## <a name="dialog-box-api-example"></a>Пример API диалоговых окной

В следующем примере кода функция использует `getTokenViaDialog` функцию API для `Dialog` отображения `displayWebDialogOptions` диалоговых окно. Этот пример предоставляется для демонстрации возможностей объекта, а не для демонстрации `Dialog` проверки подлинности.

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
Узнайте, как [выполнять отлажку пользовательских функций без пользовательского интерфейса.](custom-functions-debugging.md)

## <a name="see-also"></a>См. также

* [Runtime для пользовательских функций Excel без пользовательского интерфейса](custom-functions-runtime.md)
* [Руководство по пользовательским функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)