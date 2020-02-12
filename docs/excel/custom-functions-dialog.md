---
ms.date: 06/18/2019
description: Создание диалогового окна пользовательских функций в Excel с помощью JavaScript.
title: Вызов диалогового окна из пользовательской функции
localization_priority: Normal
ms.openlocfilehash: 8db5034cf9079ac5cd05654614087882ed1a8d52
ms.sourcegitcommit: d15bca2c12732f8599be2ec4b2adc7c254552f52
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 02/12/2020
ms.locfileid: "41950770"
---
# <a name="display-a-dialog-box-from-a-custom-function"></a>Вызов диалогового окна из пользовательской функции

Если пользовательская функция должна взаимодействовать с пользователем, можно создать диалоговое окно с помощью объекта [`Office.Dialog`](/javascript/api/office-runtime/officeruntime.dialog). Распространенным сценарием использования диалогового окна является проверка подлинности пользователя, чтобы пользовательская функция могла обращаться к веб-службе. Дополнительные сведения о проверке подлинности с помощью пользовательских функций см. в статье [Проверка подлинности пользовательских функций](./custom-functions-authentication.md).

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> Объект `Office.Dialog` является частью среды выполнения пользовательских функций. Объект `Dialog` не используется в областях задач. Сведения о создании диалогового окна из области задач см. в статье [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).

## <a name="dialog-box-api-example"></a>Пример API диалогового окна

В приведенном ниже примере кода функция `getTokenViaDialog` использует функцию `Dialog`API`displayWebDialogOptions` для отображения диалогового окна.

```js
/**
 * Function retrieves a cached token or opens a dialog box if there is no saved token. Note that this is not a sufficient example of authentication but is intended to show the capabilities of the Dialog object.
 * @param {string} url URL for a stored token.
 */
function getTokenViaDialog(url) {
  return new Promise (function (resolve, reject) {
    if (_dialogOpen) {
      // Can only have one dialog box open at once, wait for previous dialog box's token
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
Узнайте, как [создавать пользовательские функции, совместимые с функциями XLL, определенными пользователями](make-custom-functions-compatible-with-xll-udf.md).

## <a name="see-also"></a>Дополнительные ресурсы

* [Проверка подлинности пользовательских функций](custom-functions-authentication.md)
* [Получение и обработка данных с помощью пользовательских функций](custom-functions-web-reqs.md)
* [Создание пользовательских функций в Excel](custom-functions-overview.md)
