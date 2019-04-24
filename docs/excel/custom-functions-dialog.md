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
# <a name="display-a-dialog-box-in-custom-functions"></a><span data-ttu-id="9b366-103">Отображение диалогового окна в пользовательских функциях</span><span class="sxs-lookup"><span data-stu-id="9b366-103">Display a dialog box in custom functions</span></span>

<span data-ttu-id="9b366-104">Если пользовательская функция должна взаимодействовать с пользователем, можно создать диалоговое окно с помощью объекта `OfficeRuntime.Dialog`.</span><span class="sxs-lookup"><span data-stu-id="9b366-104">If your custom function needs to interact with the user, you can create a dialog box using the `OfficeRuntime.Dialog` object.</span></span> <span data-ttu-id="9b366-105">Распространенным сценарием использования диалогового окна является проверка подлинности пользователя, чтобы пользовательская функция могла обращаться к веб-службе.</span><span class="sxs-lookup"><span data-stu-id="9b366-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="9b366-106">Дополнительные сведения о проверке подлинности с помощью пользовательских функций см. в статье [Проверка подлинности пользовательских функций](./custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="9b366-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

<span data-ttu-id="9b366-107">Примечание. Объект `OfficeRuntime.Dialog` является частью среды выполнения пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="9b366-107">Note: The `OfficeRuntime.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="9b366-108">Его нельзя использовать в контексте области задач.</span><span class="sxs-lookup"><span data-stu-id="9b366-108">It cannot be used from the context of a task pane.</span></span> <span data-ttu-id="9b366-109">Сведения о создании диалогового окна из области задач см. в статье [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span><span class="sxs-lookup"><span data-stu-id="9b366-109">To create a dialog from a task pane, see [Dialog API](/office/dev/add-ins/develop/dialog-api-in-office-add-ins).</span></span>

## <a name="dialog-api-example"></a><span data-ttu-id="9b366-110">Пример Dialog API</span><span class="sxs-lookup"><span data-stu-id="9b366-110">Dialog API example</span></span>

<span data-ttu-id="9b366-111">В приведенном ниже примере кода функция `getTokenViaDialog` использует функцию `displayWebDialog` Dialog API для отображения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="9b366-111">In the following code sample, the function `getTokenViaDialog` uses the Dialog API’s `displayWebDialog` function to display a dialog box.</span></span>

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

## <a name="see-also"></a><span data-ttu-id="9b366-112">См. также</span><span class="sxs-lookup"><span data-stu-id="9b366-112">See also</span></span>

* [<span data-ttu-id="9b366-113">Метаданные пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="9b366-113">Custom functions metadata</span></span>](custom-functions-json.md)
* [<span data-ttu-id="9b366-114">Среда выполнения для пользовательских функций Excel</span><span class="sxs-lookup"><span data-stu-id="9b366-114">Runtime for Excel custom functions</span></span>](custom-functions-runtime.md)
* [<span data-ttu-id="9b366-115">Рекомендации по пользовательским функциям</span><span class="sxs-lookup"><span data-stu-id="9b366-115">Custom functions best practices</span></span>](custom-functions-best-practices.md)
* [<span data-ttu-id="9b366-116">Журнал изменений пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="9b366-116">Custom functions changelog</span></span>](custom-functions-changelog.md)
* [<span data-ttu-id="9b366-117">Руководство по настраиваемым функциям в Excel</span><span class="sxs-lookup"><span data-stu-id="9b366-117">Excel custom functions tutorial</span></span>](../tutorials/excel-tutorial-create-custom-functions.md)
