---
ms.date: 05/03/2019
description: Обработка ошибок в пользовательских функциях Excel.
title: Обработка ошибок в пользовательских функциях Excel
localization_priority: Priority
ms.openlocfilehash: 188ece6c77bc2cafad6f22448fb698e0c0370ef8
ms.sourcegitcommit: ff73cc04e5718765fcbe74181505a974db69c3f5
ms.translationtype: HT
ms.contentlocale: ru-RU
ms.lasthandoff: 05/06/2019
ms.locfileid: "33628160"
---
# <a name="error-handling-within-custom-functions"></a><span data-ttu-id="3d921-103">Обработка ошибок в пользовательских функциях</span><span class="sxs-lookup"><span data-stu-id="3d921-103">Error handling within custom functions</span></span>

<span data-ttu-id="3d921-104">При создании надстройки, которая определяет пользовательские функции, не забудьте включить логику для обработки ошибок, возникающих в среде выполнения.</span><span class="sxs-lookup"><span data-stu-id="3d921-104">When you build an add-in that defines custom functions, be sure to include error handling logic to account for runtime errors.</span></span> <span data-ttu-id="3d921-105">Обработка ошибок для пользовательских функций в значительной степени совпадает с [обработкой ошибок для API JavaScript в Excel](excel-add-ins-error-handling.md).</span><span class="sxs-lookup"><span data-stu-id="3d921-105">Error handling for custom functions is the same as [error handling for the Excel JavaScript API at large](excel-add-ins-error-handling.md).</span></span>

[!include[Excel custom functions note](../includes/excel-custom-functions-note.md)]

<span data-ttu-id="3d921-106">В следующем примере кода `.catch` будет обрабатывать любые ошибки, возникающие ранее в коде.</span><span class="sxs-lookup"><span data-stu-id="3d921-106">In the following code sample, `.catch` will handle any errors that occur previously in the code.</span></span>

```js
/**
 * Gets a comment from the hypothetical contoso.com/comments API.
 * @customfunction
 * @param {number} commentID ID of a comment.
 */
function getComment(commentID) {
  let url = "https://www.contoso.com/comments/" + x;

  return fetch(url)
    .then(function (data) {
      return data.json();
    })
    .then(function (json) {
      return json.body;
    })
    .catch(function (error) {
      throw error;
    })
}
```

## <a name="next-steps"></a><span data-ttu-id="3d921-107">Дальнейшие действия</span><span class="sxs-lookup"><span data-stu-id="3d921-107">Next steps</span></span>
<span data-ttu-id="3d921-108">Узнайте, как [устранять проблемы с пользовательскими функциями](custom-functions-troubleshooting.md).</span><span class="sxs-lookup"><span data-stu-id="3d921-108">Learn how to [troubleshoot problems with your custom functions](custom-functions-troubleshooting.md).</span></span>

## <a name="see-also"></a><span data-ttu-id="3d921-109">Дополнительные ресурсы</span><span class="sxs-lookup"><span data-stu-id="3d921-109">See also</span></span>

* [<span data-ttu-id="3d921-110">Отладка пользовательских функций</span><span class="sxs-lookup"><span data-stu-id="3d921-110">Custom functions debugging</span></span>](custom-functions-debugging.md)
* [<span data-ttu-id="3d921-111">Требования к настраиваемым функциям</span><span class="sxs-lookup"><span data-stu-id="3d921-111">Custom functions requirements</span></span>](custom-functions-requirements.md)
* [<span data-ttu-id="3d921-112">Создание пользовательских функций в Excel</span><span class="sxs-lookup"><span data-stu-id="3d921-112">Create custom functions in Excel</span></span>](custom-functions-overview.md)
