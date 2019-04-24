---
ms.date: 03/21/2019
description: Создание диалоговых окон пользовательских функций в Excel с помощью JavaScript.
title: Диалоговые окна пользовательских функций (предварительная версия)
localization_priority: Priority
ms.openlocfilehash: 0f596825a7a32525a68ef45656f1390196146706
ms.sourcegitcommit: 9e7b4daa8d76c710b9d9dd4ae2e3c45e8fe07127
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 04/24/2019
ms.locfileid: "32449262"
---
# <a name="display-a-dialog-box-in-custom-functions"></a>Отображение диалогового окна в пользовательских функциях

Если пользовательская функция должна взаимодействовать с пользователем, можно создать диалоговое окно с помощью объекта `OfficeRuntime.Dialog`. Распространенным сценарием использования диалогового окна является проверка подлинности пользователя, чтобы пользовательская функция могла обращаться к веб-службе. Дополнительные сведения о проверке подлинности с помощью пользовательских функций см. в статье [Проверка подлинности пользовательских функций](./custom-functions-authentication.md).

Примечание. Объект `OfficeRuntime.Dialog` является частью среды выполнения пользовательских функций. Его нельзя использовать в контексте области задач. Сведения о создании диалогового окна из области задач см. в статье [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).

## <a name="dialog-api-example"></a>Пример Dialog API

В приведенном ниже примере кода функция `getTokenViaDialog` использует функцию `displayWebDialog` Dialog API для отображения диалогового окна.

```js
// Get auth token before calling my service, a hypothetical API that will deliver a stock price based on stock ticker string, such as "MSFT"

function getStock (ticker) {
  return new Promise(function (resolve, reject) {
    // Get a token
    getToken("https://www.contoso.com/auth")
    .then(function (token) {

      // Use token to get stock price
      fetch("https://www.contoso.com/?token=token&ticker= + ticker")
      .then(function (result) {

        // Return stock price to cell
        resolve(result);
      });
    })
    .catch(function (error) {
      reject(error);
    });
  });
  
  //Helper
  function getToken(url) {
    return new Promise(function (resolve,reject) {
      if(_cachedToken) {
        resolve(_cachedToken);
      } else {
        getTokenViaDialog(url)
        .then(function (result) {
          resolve(result);
        })
        .catch(function (result) {
          reject(result);
        });
      }
    });
  }

  function getTokenViaDialog(url) {
    return new Promise (function (resolve, reject) {
      if (_dialogOpen) {
        // Can only have one dialog open at once, wait for previous dialog's token
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
}
```

## <a name="see-also"></a>См. также

* [Метаданные пользовательских функций](custom-functions-json.md)
* [Среда выполнения для пользовательских функций Excel](custom-functions-runtime.md)
* [Рекомендации по пользовательским функциям](custom-functions-best-practices.md)
* [Журнал изменений пользовательских функций](custom-functions-changelog.md)
* [Руководство по настраиваемым функциям в Excel](../tutorials/excel-tutorial-create-custom-functions.md)
