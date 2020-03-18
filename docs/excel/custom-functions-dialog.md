---
ms.date: 06/18/2019
description: Создание диалогового окна пользовательских функций в Excel с помощью JavaScript.
title: Вызов диалогового окна из пользовательской функции
localization_priority: Normal
ms.openlocfilehash: a2ef005f4c1519228f114dbd671d689807e5914c
ms.sourcegitcommit: fa4e81fcf41b1c39d5516edf078f3ffdbd4a3997
ms.translationtype: MT
ms.contentlocale: ru-RU
ms.lasthandoff: 03/17/2020
ms.locfileid: "42718749"
---
# <a name="display-a-dialog-box-from-a-custom-function"></a><span data-ttu-id="6a045-103">Вызов диалогового окна из пользовательской функции</span><span class="sxs-lookup"><span data-stu-id="6a045-103">Display a dialog box from a custom function</span></span>

<span data-ttu-id="6a045-104">Если пользовательская функция должна взаимодействовать с пользователем, можно создать диалоговое окно с помощью объекта [`Office.Dialog`](/javascript/api/office-runtime/officeruntime.dialog).</span><span class="sxs-lookup"><span data-stu-id="6a045-104">If your custom function needs to interact with the user, you can create a dialog box using the [`Office.Dialog` object](/javascript/api/office-runtime/officeruntime.dialog).</span></span> <span data-ttu-id="6a045-105">Распространенным сценарием использования диалогового окна является проверка подлинности пользователя, чтобы пользовательская функция могла обращаться к веб-службе.</span><span class="sxs-lookup"><span data-stu-id="6a045-105">A common scenario for using the dialog box is to authenticate a user so that your custom function can access a web service.</span></span> <span data-ttu-id="6a045-106">Дополнительные сведения о проверке подлинности с помощью пользовательских функций см. в статье [Проверка подлинности пользовательских функций](./custom-functions-authentication.md).</span><span class="sxs-lookup"><span data-stu-id="6a045-106">For more details about authentication with custom functions, see [Custom functions authentication](./custom-functions-authentication.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

>[!NOTE]
> <span data-ttu-id="6a045-107">Объект `Office.Dialog` является частью среды выполнения пользовательских функций.</span><span class="sxs-lookup"><span data-stu-id="6a045-107">The `Office.Dialog` object is part of the custom functions runtime.</span></span> <span data-ttu-id="6a045-108">Объект `Dialog` не используется в областях задач.</span><span class="sxs-lookup"><span data-stu-id="6a045-108">Task panes don't use the `Dialog` object.</span></span> <span data-ttu-id="6a045-109">Сведения о создании диалогового окна из области задач см. в статье [Dialog API](../develop/dialog-api-in-office-add-ins.md).</span><span class="sxs-lookup"><span data-stu-id="6a045-109">To create a dialog box from a task pane, see [Dialog API](../develop/dialog-api-in-office-add-ins.md).</span></span>

## <a name="dialog-box-api-example"></a><span data-ttu-id="6a045-110">Пример API диалогового окна</span><span class="sxs-lookup"><span data-stu-id="6a045-110">dialog box API example</span></span>

<span data-ttu-id="6a045-111">В следующем примере кода функция `getTokenViaDialog` использует `Dialog` `displayWebDialogOptions` функцию API для отображения диалогового окна.</span><span class="sxs-lookup"><span data-stu-id="6a045-111">In the following code sample, the function `getTokenViaDialog` uses the `Dialog` API's `displayWebDialogOptions` function to display a dialog box.</span></span>

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

## <a name="next-steps"></a><span data-ttu-id="6a045-112">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="6a045-112">Next steps</span></span>
<span data-ttu-id="6a045-113">Узнайте, как [создавать пользовательские функции, совместимые с функциями XLL, определенными пользователями](make-custom-functions-compatible-with-xll-udf.md).</span><span class="sxs-lookup"><span data-stu-id="6a045-113">Learn how to [make your custom functions compatible with XLL user-defined functions](make-custom-functions-compatible-with-xll-udf.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="6a045-114">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="6a045-114">See also</span></span>

* [<span data-ttu-id="6a045-115">Проверка подлинности пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="6a045-115">Custom functions authentication</span></span>](custom-functions-authentication.md)
* [<span data-ttu-id="6a045-116">Получение и обработка данных с помощью пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="6a045-116">Receive and handle data with custom functions</span></span>](custom-functions-web-reqs.md)
* [<span data-ttu-id="6a045-117">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="6a045-117">Create custom functions in Excel</span></span>](custom-functions-overview.md)
