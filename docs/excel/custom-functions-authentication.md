---
ms.date: 05/17/2020
description: Проверка подлинности пользователей с помощью пользовательских функций в Excel, не использующих область задач.
title: Проверка подлинности для пользовательских функций без пользовательского интерфейса
localization_priority: Normal
ms.openlocfilehash: 93073fb23f3f4d30c36faf4927a3aebdafbc887d
ms.sourcegitcommit: f62d9630de69c5c070e3d4048205f5cc654db7e4
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/18/2020
ms.locfileid: "44278380"
---
# <a name="authentication-for-ui-less-custom-functions"></a>Проверка подлинности для пользовательских функций без пользовательского интерфейса

В некоторых сценариях пользовательская функция, не использующая область задач или другие элементы пользовательского интерфейса (настраиваемая функция без ПОЛЬЗОВАТЕЛЬСКОГО интерфейса пользователя), должна выполнять проверку подлинности пользователя для доступа к защищенным ресурсам. Имейте в виду, что пользовательские функции без пользовательского интерфейса выполняются в среде выполнения с поддержкой JavaScript. Поэтому необходимо передавать данные между средой выполнения JavaScript и обычной средой выполнения, используемой большинством надстроек, с помощью `OfficeRuntime.storage` объекта и API диалоговых окон.

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

[!include[Shared runtime note](../includes/shared-runtime-note.md)]

## <a name="officeruntimestorage-object"></a>Объект OfficeRuntime.storage

В среде выполнения с поддержкой JavaScript, используемой пользовательскими функциями без пользовательского интерфейса `localStorage` , отсутствует объект, доступный в глобальном окне, где обычно хранятся данные. Вместо этого необходимо обмениваться данными между пользовательскими функциями и областями задач с помощью [оффицерунтиме. Storage](/javascript/api/office-runtime/officeruntime.storage) для задания и получения данных.

### <a name="suggested-usage"></a>Рекомендуемое использование

Если вам нужно выполнить проверку подлинности в пользовательской функции без пользовательского интерфейса, проверьте, `storage` был ли уже получен маркер доступа. Если нет, используйте API диалоговых окон, чтобы проверить подлинность пользователя, извлечь маркер доступа и сохранить его в объекте `storage` для дальнейшего использования.

## <a name="dialog-api"></a>API диалоговых окон

Если маркер не существует, следует использовать API диалоговых окон, чтобы попросить пользователя войти в систему. После ввода пользователем своих учетных данных итоговый маркер доступа можно сохранить в объекте `storage`.

> [!NOTE]
> В среде выполнения, предназначенной только для JavaScript, используется объект Dialog, который немного отличается от объекта Dialog в среде выполнения модуля браузера, используемого панелями задач. Они обе называются "диалоговым API", но используются `OfficeRuntime.Dialog` для проверки подлинности пользователей в среде выполнения только JavaScript.

На следующей схеме показан этот основной процесс. Пунктирная линия указывает на то, что пользовательские функции без пользовательского интерфейса и область задач надстройки являются частью надстройки в целом, хотя в них используются отдельные среды выполнения.

1. Вы выдаете вызов пользовательской функции без пользовательского интерфейса из ячейки в книге Excel.
2. Пользовательская функция без пользовательского интерфейса использует `Dialog` для передачи учетных данных пользователя на веб-сайт.
3. Затем этот веб-сайт возвращает маркер доступа для пользовательской функции без пользовательского интерфейса.
4. Пользовательская функция без пользовательского интерфейса устанавливает для маркера доступа значение `storage` .
5. Область задач надстройки получает доступ к маркеру из объекта `storage`.

![Схема пользовательской функции с помощью API диалога для получения маркера доступа, а затем совместного использования маркера с областью задач с помощью API Оффицерунтиме. Storage.](../images/authentication-diagram.png "Схема проверки подлинности.")

## <a name="storing-the-token"></a>Хранение маркера

Следующие примеры взяты из примера кода [Использование OfficeRuntime.storage в пользовательских функциях](https://github.com/OfficeDev/PnP-OfficeAddins/tree/master/Excel-custom-functions/AsyncStorage). В этом примере кода приведен полный пример общего доступа к данным для пользовательских функций без пользовательского интерфейса и области задач.

При проверке подлинности настраиваемой функции, не требующей пользовательского интерфейса, она получает маркер доступа и должна храниться в `storage` . В следующем примере кода показано, как вызвать метод `storage.setItem` чтобы сохранить значение. `storeValue`Функция — это пользовательская функция без пользовательского интерфейса, в которой в качестве примера хранится значение пользователя. Можно внести изменение, чтобы сохранять любые нужные значения маркеров.

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

Надстройки Office являются веб-надстройками, и вы можете использовать любой способ веб-проверки подлинности. Для реализации собственной проверки подлинности с пользовательскими функциями без пользовательского интерфейса нет определенного шаблона или метода. Рекомендуется ознакомиться с документацией по различным шаблонам проверки подлинности, начиная с [этой статьи об авторизации через внешние службы](../develop/auth-external-add-ins.md).  

Избегайте использования следующих расположений для хранения данных при разработке пользовательских функций.  

- `localStorage`: Пользовательские функции без пользовательского интерфейса не имеют доступа к глобальному `window` объекту и поэтому не имеют доступа к данным, хранящимся в `localStorage` .
- `Office.context.document.settings`: это расположение не защищено, и сведения могут быть извлечены любым пользователем с помощью надстройки.

## <a name="dialog-box-api-example"></a>Пример API диалогового окна

В следующем примере кода функция `getTokenViaDialog` использует `Dialog` `displayWebDialogOptions` функцию API для отображения диалогового окна. Этот пример предоставляется для отображения возможностей `Dialog` объекта, не демонстрирующи способов проверки подлинности.

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
Узнайте, как [отлаживать пользовательские функции без пользовательского интерфейса](custom-functions-debugging.md).

## <a name="see-also"></a>См. также

* [Среда выполнения для пользовательских функций Excel без пользовательского интерфейса](custom-functions-runtime.md)
* [Руководство по пользовательским функциям в Excel](excel-tutorial-custom-functions.md)
